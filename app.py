
import streamlit as st
import base64
import os
import zipfile
import posixpath
import xml.etree.ElementTree as ET
import re
import html
import hashlib
import hmac
import traceback
from io import BytesIO
from datetime import date

from pptx import Presentation

st.set_page_config(
    page_title="Material Scanner Agent",
    page_icon="🔎",
    layout="wide",
)

APP_NAME = "Material Scanner Agent"
APP_SUBTITLE = "AI-powered Material Pre-Review Prototype for AZ All Hackathon"

PROMOMAT = "PROMOMAT"
NON_PROMOMAT = "NON-PROMOMAT"
ENGINE_PROFILE = {
    PROMOMAT: "PROMOMAT",
    NON_PROMOMAT: "NON_PROMOMAT",
}

# ------------------------------------------------------------
# Project modules
# ------------------------------------------------------------
try:
    from pptx_extractor import extract_pptx
except Exception:
    extract_pptx = None

try:
    from msa_engine import analyze_material
except Exception:
    analyze_material = None

try:
    from evidence_layer import attach_evidence
except Exception:
    attach_evidence = None

try:
    from ai_context_analysis import analyze_slides
except Exception:
    analyze_slides = None

try:
    from rule_manager import (
        load_rules,
        get_rule,
        add_rule,
        update_rule,
        delete_rule,
        enable_rule,
        disable_rule,
        count_rules,
        count_active_rules,
        upsert_rules,
    )
except Exception:
    load_rules = get_rule = add_rule = update_rule = delete_rule = None
    enable_rule = disable_rule = None
    count_rules = count_active_rules = upsert_rules = None


# ------------------------------------------------------------
# Session state
# ------------------------------------------------------------
DEFAULTS = {
    "review_profile": None,
    "uploaded_material": None,
    "uploaded_filename": None,
    "uploaded_hash": None,
    "slides": None,
    "analysis_result": None,
    "evidence_result": None,
    "ai_result": None,
    "review_result": None,
    "review_started": False,
    "rule_admin_authenticated": False,
    "rule_admin_login_version": 0,
    "rule_admin_login_error": "",
    "rule_management_flash": None,
    "uploader_version": 0,
    "page": "Review",
    "show_new_review_confirm": False,
    "result_section": "Summary",
    "cover_generated": False,
}
for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ------------------------------------------------------------
# CSS
# ------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
css_path = os.path.join(BASE_DIR, "styles.css")
banner_path = os.path.join(BASE_DIR, "assets", "msa-banner.png")

if os.path.exists(css_path):
    with open(css_path, "r", encoding="utf-8") as f:
        css = f.read()

    # Embed the local hero banner into the CSS so it works reliably in
    # Streamlit both locally and on Streamlit Community Cloud. CSS-relative
    # asset URLs may resolve against the Streamlit page rather than the repo.
    if os.path.isfile(banner_path):
        with open(banner_path, "rb") as f:
            banner_base64 = base64.b64encode(f.read()).decode("ascii")

        banner_data_uri = f"url('data:image/png;base64,{banner_base64}')"
        css, replacements = re.subn(
                r"""url\(\s*(['"]?)assets/msa-banner\.png\1\s*\)""",
                lambda _match: banner_data_uri,
                css,
                flags=re.IGNORECASE,
            )


        # Keep the app running if the CSS path changes unexpectedly, while
        # surfacing a clear diagnostic in the terminal logs.
        if replacements == 0:
            print(
                "MSA banner image exists, but its URL was not "
                "matched in styles.css. Check the .msa-hero-glass rule."
            )
    else:
        print(f"MSA banner image not found: {banner_path}")

    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------
def esc(value):
    return html.escape(str(value if value is not None else ""))

def clean_rule_description(value):
    """Keep prototype/test markers out of the end-user review experience.

    Rule data remains unchanged in Rule Management; only the display text is
    normalized for Findings/Evidence.
    """
    text = str(value or "").strip()
    prefixes = (
        "DEMO RULE ONLY - Used to test MSA rule processing.",
        "DEMO RULE ONLY - Used to test MSA context processing.",
        "DEMO RULE ONLY - Configure with approved rule content before real use.",
    )
    for prefix in prefixes:
        if text.lower() == prefix.lower():
            return "Illustrative screening rule used for prototype demonstration."
    return text




def render_html(raw):
    """Render HTML with no leading whitespace on any line."""
    cleaned = "\n".join(line.lstrip() for line in str(raw).splitlines()).strip()
    st.markdown(cleaned, unsafe_allow_html=True)


def status_class(value):
    normalized = str(value or "").strip().lower().replace("_", "-").replace(" ", "-")
    return {
        "high": "msa-status-high",
        "medium": "msa-status-medium",
        "low": "msa-status-low",
        "review-required": "msa-status-medium",
        "review": "msa-status-medium",
        "info": "msa-status-info",
        "ai-assisted": "msa-status-info",
    }.get(normalized, "msa-status-info")


def normalize_slides(slides):
    """Normalize all extractor outputs to the canonical engine format.

    The core engine expects:
        {"slide_number": 1, "text": "..."}

    The UI also uses the richer `texts` list for Extracted Content, so we keep
    both representations in each slide record.
    """
    if not slides:
        return []

    normalized = []

    for idx, slide in enumerate(slides, 1):
        if isinstance(slide, dict):
            number = slide.get("slide_number", slide.get("slide", idx))

            raw_texts = slide.get("texts")
            if raw_texts is None:
                raw_texts = slide.get("text", [])

            if isinstance(raw_texts, str):
                texts = [raw_texts]
            elif isinstance(raw_texts, (list, tuple)):
                texts = list(raw_texts)
            else:
                texts = []

            texts = [str(x).strip() for x in texts if str(x).strip()]

            # If a caller supplied both text and texts, retain the explicit
            # text only when it contains information not already represented.
            explicit_text = str(slide.get("text", "") or "").strip()
            if explicit_text and not texts:
                texts = [explicit_text]

            combined_text = explicit_text if explicit_text else "\n".join(texts)

            normalized.append({
                "slide_number": number,
                "slide": number,
                "text": combined_text,
                "texts": texts,
            })
        else:
            text = str(slide).strip()
            normalized.append({
                "slide_number": idx,
                "slide": idx,
                "text": text,
                "texts": [text] if text else [],
            })

    return normalized


def extract_material(data):
    if extract_pptx is not None:
        try:
            return extract_pptx(BytesIO(data))
        except Exception:
            try:
                return extract_pptx(data)
            except Exception:
                pass

    prs = Presentation(BytesIO(data))
    result = []
    for number, slide in enumerate(prs.slides, 1):
        texts = []
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                text = str(shape.text or "").strip()
                if text:
                    texts.append(text)
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        text = str(cell.text or "").strip()
                        if text:
                            texts.append(text)
        result.append({"slide": number, "texts": list(dict.fromkeys(texts))})
    return result


def reset_review_state(clear_material=False):
    st.session_state.slides = None
    st.session_state.analysis_result = None
    st.session_state.evidence_result = None
    st.session_state.ai_result = None
    st.session_state.review_result = None
    st.session_state.review_started = False
    st.session_state.result_section = "Summary"
    st.session_state.cover_generated = False

    if clear_material:
        st.session_state.uploaded_material = None
        st.session_state.uploaded_filename = None
        st.session_state.uploaded_hash = None
        st.session_state.uploader_version += 1


def counts(analysis):
    """Read authoritative totals from msa_engine.py output."""
    if not isinstance(analysis, dict):
        return {"slides": 0, "findings": 0, "high": 0, "medium": 0, "low": 0}

    slides = analysis.get("slides", [])
    summary = analysis.get("summary", {})
    if not isinstance(summary, dict):
        summary = {}

    slide_count = int(summary.get("slides_analyzed", len(slides) if isinstance(slides, list) else 0))
    finding_count = int(summary.get("rule_findings", 0))
    high = int(summary.get("high", 0))
    medium = int(summary.get("medium", 0))
    low = int(summary.get("low", 0))

    # Fallback for older engine output without summary.
    if not summary:
        finding_count = high = medium = low = 0
        for slide in slides if isinstance(slides, list) else []:
            if not isinstance(slide, dict):
                continue
            for finding in slide.get("rule_findings", []):
                if not isinstance(finding, dict):
                    continue
                finding_count += 1
                severity = str(finding.get("severity", "")).upper()
                if severity == "HIGH":
                    high += 1
                elif severity == "MEDIUM":
                    medium += 1
                elif severity == "LOW":
                    low += 1

    return {
        "slides": slide_count,
        "findings": finding_count,
        "high": high,
        "medium": medium,
        "low": low,
    }


def get_profile_mapping(profile):
    """Read active rules configured for the selected Review Profile."""
    canonical = ENGINE_PROFILE.get(profile, profile)
    mapping = {
        "profile": canonical,
        "rule_count": 0,
        "rules": [],
    }
    if load_rules is None:
        return mapping

    try:
        rules = load_rules()
    except Exception:
        return mapping

    for rule in rules if isinstance(rules, list) else []:
        if not isinstance(rule, dict):
            continue
        status = str(rule.get("status", "")).strip().lower()
        approval = str(rule.get("approval_status", "")).strip().upper()
        applies_to = rule.get("applies_to", rule.get("category", ""))
        if isinstance(applies_to, str):
            applies = {x.strip().upper().replace("-", "_") for x in re.split(r"[;,|]", applies_to) if x.strip()}
        elif isinstance(applies_to, (list, tuple, set)):
            applies = {str(x).strip().upper().replace("-", "_") for x in applies_to if str(x).strip()}
        else:
            applies = set()
        applies.add(str(rule.get("category", "")).strip().upper().replace("-", "_"))
        is_approved = not approval or approval in {"APPROVED", "NS_APPROVED", "APPROVED_BY_NS", "APPROVED_BY_SOP_OWNER"}
        if str(canonical).upper().replace("-", "_") in applies and status == "active" and is_approved:
            mapping["rules"].append(rule)

    mapping["rule_count"] = len(mapping["rules"])
    return mapping


def render_profile_mapping(profile, mapping=None):
    """Render the selected profile and the active rules assigned to it."""
    mapping = mapping or get_profile_mapping(profile)
    rules = mapping.get("rules", []) if isinstance(mapping, dict) else []

    label = "NON-PROMOMAT" if profile == NON_PROMOMAT else "PROMOMAT"
    description = (
        "Applies active rules configured for promotional-material screening."
        if profile == PROMOMAT
        else "Applies active rules configured for non-promotional-material screening."
    )

    header("◇", "Rule Profile Mapping", "The selected profile determines which active screening rules are applied.")
    render_html(
        f'<div class="msa-profile-map-card">'
        f'<div class="msa-profile-map-top">'
        f'<div><div class="msa-profile-map-label">SELECTED PROFILE</div>'
        f'<div class="msa-profile-map-title">{esc(label)}</div>'
        f'<div class="msa-profile-map-description">{esc(description)}</div></div>'
        f'<div class="msa-profile-map-count"><strong>{esc(len(rules))}</strong><span>Active Rules</span></div>'
        f'</div>'
    )

    if rules:
        rows = []
        for rule in rules:
            rows.append(
                f'<div class="msa-profile-rule-row">'
                f'<div class="msa-profile-rule-id">{esc(rule.get("rule_id", ""))}</div>'
                f'<div class="msa-profile-rule-name">{esc(rule.get("rule_name", ""))}</div>'
                f'<div class="msa-profile-rule-meta">'
                f'{esc(rule.get("rule_type", ""))} · {esc(rule.get("severity", ""))} · '
                f'{esc(rule.get("action", ""))} · v{esc(rule.get("version", "1.0"))}'
                f'</div></div>'
            )
        render_html(
            '<div class="msa-profile-rule-list">' + ''.join(rows) + '</div></div>'
        )
    else:
        render_html(
            '<div class="msa-profile-map-empty">'
            'No active rules are currently mapped to this profile. '
            'Configure rules in Rule Management before running a review.'
            '</div></div>'
        )


def run_engine(slides, profile):
    if analyze_material is None:
        raise RuntimeError("msa_engine.py could not be loaded.")
    return analyze_material(slides, profile)


def run_evidence(analysis):
    if attach_evidence is None:
        return []
    try:
        return attach_evidence(analysis)
    except Exception:
        return []


def run_ai(slides):
    """
    Normalize ai_context_analysis.py output for the UI.

    Current module output is a list of:
    slide_number, classification, confidence, reason, matched_terms.
    """
    if analyze_slides is None:
        return {
            "status": "AI NOT CONFIGURED",
            "overall_summary": "AI Context Analysis module is not connected to an approved AI endpoint.",
            "items": [],
        }

    try:
        result = analyze_slides(slides)

        if isinstance(result, list):
            items = []
            for raw in result:
                if hasattr(raw, "__dict__"):
                    raw = raw.__dict__
                if not isinstance(raw, dict):
                    continue

                slide_number = raw.get("slide_number", raw.get("slide", "?"))
                classification = raw.get(
                    "classification",
                    raw.get("category", "UNCERTAIN"),
                )
                confidence = raw.get("confidence", 0)
                reason = raw.get(
                    "reason",
                    raw.get("context", "No contextual interpretation available."),
                )
                matched_terms = raw.get(
                    "matched_terms",
                    raw.get("keywords", []),
                )
                if isinstance(matched_terms, str):
                    matched_terms = [matched_terms]
                matched_terms = [str(x) for x in (matched_terms or []) if str(x).strip()]

                try:
                    confidence_value = float(confidence)
                    confidence_percent = (
                        round(confidence_value * 100)
                        if confidence_value <= 1
                        else round(confidence_value)
                    )
                except (TypeError, ValueError):
                    confidence_percent = 0

                items.append({
                    "slide": slide_number,
                    "slide_number": slide_number,
                    "category": classification,
                    "classification": classification,
                    "confidence": confidence_percent,
                    "reason": reason,
                    "context": reason,
                    "keywords": matched_terms,
                    "matched_terms": matched_terms,
                    "risk_level": "Medium" if classification == "PROMOTIONAL_CONTEXT" else "Info",
                    "assessment": classification,
                    "recommendation": (
                        "Review the slide context and confirm the applicable "
                        "review profile and reference."
                    ),
                })

            return {
                "status": "COMPLETED",
                "overall_summary": "Context analysis completed.",
                "items": items,
            }

        if isinstance(result, dict):
            normalized = dict(result)
            raw_items = normalized.get("items", [])
            if not isinstance(raw_items, list):
                raw_items = []

            items = []
            for raw in raw_items:
                if not isinstance(raw, dict):
                    continue
                item = dict(raw)
                slide_number = item.get("slide_number", item.get("slide", "?"))
                classification = item.get(
                    "classification",
                    item.get("category", "UNCERTAIN"),
                )
                confidence = item.get("confidence", 0)
                try:
                    confidence_value = float(confidence)
                    confidence_percent = (
                        round(confidence_value * 100)
                        if confidence_value <= 1
                        else round(confidence_value)
                    )
                except (TypeError, ValueError):
                    confidence_percent = 0

                item["slide"] = slide_number
                item["slide_number"] = slide_number
                item["category"] = classification
                item["classification"] = classification
                item["confidence"] = confidence_percent
                item["context"] = item.get(
                    "context",
                    item.get("reason", "No contextual interpretation available."),
                )
                item["reason"] = item.get("reason", item["context"])
                item["keywords"] = item.get(
                    "keywords",
                    item.get("matched_terms", []),
                )
                item["matched_terms"] = item.get(
                    "matched_terms",
                    item["keywords"],
                )
                item.setdefault("assessment", classification)
                item.setdefault("risk_level", "Info")
                item.setdefault(
                    "recommendation",
                    "Review the slide context and confirm the applicable review profile and reference.",
                )
                items.append(item)

            return {
                "status": normalized.get("status", "COMPLETED"),
                "overall_summary": normalized.get(
                    "overall_summary",
                    "Context analysis completed.",
                ),
                "items": items,
                "error": normalized.get("error"),
            }

        return {
            "status": "COMPLETED",
            "overall_summary": "Context analysis completed.",
            "items": [],
        }

    except Exception as error:
        return {
            "status": "AI ERROR",
            "overall_summary": "Context analysis could not be completed.",
            "items": [],
            "error": str(error),
        }


def unique_findings(analysis):
    """
    Flatten rule findings from the current msa_engine structure.

    msa_engine.py stores findings under:
        analysis["slides"][n]["rule_findings"]

    Older versions may also expose a top-level rule_findings/findings list.
    """
    if not isinstance(analysis, dict):
        return []

    candidates = []

    slide_results = analysis.get("slides", [])
    if isinstance(slide_results, list):
        for slide in slide_results:
            if not isinstance(slide, dict):
                continue

            slide_number = slide.get(
                "slide_number",
                slide.get("slide", "?"),
            )
            slide_text = slide.get("text", "")

            for raw in slide.get("rule_findings", []) or []:
                if not isinstance(raw, dict):
                    continue
                item = dict(raw)
                item.setdefault("slide_number", slide_number)
                item.setdefault("slide", slide_number)
                item.setdefault("text", slide_text)
                candidates.append(item)

    # Backward compatibility for engines that expose top-level findings.
    for key in ("rule_findings", "findings"):
        raw_items = analysis.get(key, [])
        if isinstance(raw_items, list):
            for raw in raw_items:
                if isinstance(raw, dict):
                    candidates.append(dict(raw))

    output = []
    seen = set()

    for item in candidates:
        key = (
            item.get("slide_number", item.get("slide", "")),
            item.get("rule_id", ""),
            item.get("rule_name", item.get("category", "")),
            item.get("severity", ""),
            tuple(item.get("matched_keywords", item.get("keywords", [])))
            if isinstance(item.get("matched_keywords", item.get("keywords", [])), list)
            else str(item.get("matched_keywords", item.get("keywords", ""))),
        )

        if key in seen:
            continue

        seen.add(key)
        output.append(item)

    return output


def normalize_evidence_items(evidence):
    if evidence is None:
        return []
    if isinstance(evidence, dict):
        for key in ("items", "evidence", "findings"):
            if isinstance(evidence.get(key), list):
                return evidence[key]
        return [evidence]
    return evidence if isinstance(evidence, list) else []


def evidence_for_finding(evidence, finding):
    items = normalize_evidence_items(evidence)
    slide = finding.get("slide_number", finding.get("slide", ""))
    rule_id = finding.get("rule_id", "")
    matches = []
    for item in items:
        if not isinstance(item, dict):
            continue
        item_slide = item.get("slide_number", item.get("slide", ""))
        item_rule = item.get("rule_id", "")
        # Evidence must match both slide and rule ID. Never borrow evidence
        # from another rule on the same slide when the finding has no rule ID.
        if rule_id and str(item_slide) == str(slide) and str(item_rule) == str(rule_id):
            matches.append(item)
    return matches


def detected_values(finding):
    values = (
        finding.get("matched_keywords")
        or finding.get("detected_signals")
        or finding.get("keywords")
        or []
    )
    if isinstance(values, str):
        values = [values]
    return [str(v) for v in values if str(v).strip()]


def ai_item_for_finding(ai_result, finding):
    if not isinstance(ai_result, dict):
        return None
    items = ai_result.get("items", [])
    if not isinstance(items, list):
        return None
    slide = finding.get("slide_number", finding.get("slide", ""))
    for item in items:
        if not isinstance(item, dict):
            continue
        if str(item.get("slide", "")) == str(slide):
            return item
    return None


def build_ai_explanation(finding, ai_result):
    """Build a finding-specific explanation, using slide context only as secondary evidence.

    This prototype currently uses transparent rule-family reasoning, not a live generative
    model call. It avoids presenting generic slide-level context text as if it explained
    every rule finding.
    """
    item = ai_item_for_finding(ai_result, finding)
    slide = finding.get("slide_number", finding.get("slide", "?"))
    rule_id = str(finding.get("rule_id", "") or "").strip()
    rule_name = str(finding.get("rule_name", finding.get("category", "configured rule")) or "configured rule").strip()
    signals = detected_values(finding)
    signal_text = ", ".join(signals[:5]) if signals else "the matched slide content"
    description = str(finding.get("description", finding.get("rule_description", "")) or "").strip()
    identity = f"{rule_id} {rule_name} {' '.join(signals)}".casefold()

    # Rule-specific explanation and next step: do not reuse one generic AI sentence for all findings.
    if any(x in identity for x in ("comparative", "more effective", "better than", "superior", "inferior", "compared")):
        explanation = (
            f"On Slide {slide}, the wording '{signal_text}' may be read as a comparison or superiority statement. "
            "The key question is whether the comparator, outcome and supporting evidence align with the claim as presented."
        )
        next_step = (
            "Check the exact comparator and outcome, then compare the wording with the cited evidence. Confirm that the comparison is direct, fair and balanced, "
            "including any important limitations or context."
        )
        context = "A comparative phrase was matched. The phrase alone cannot confirm whether the comparison is supported or appropriate."
    elif any(x in identity for x in ("off-label", "indication", "dosage", "dose", "population", "route of administration", "prescribing information")):
        explanation = (
            f"On Slide {slide}, the matched wording ({signal_text}) relates to product use or prescribing information. "
            "Whether it is appropriate depends on the current locally approved product information and the context in which it is used."
        )
        next_step = "Compare the exact statement with the current locally approved product information for indication, population, dose and route as applicable; escalate any mismatch."
        context = "The match identifies a product-use topic for verification; it is not, by itself, evidence of off-label promotion."
    elif any(x in identity for x in ("approval", "approved material", "expiry", "expiration", "expired", "renewal", "re-approval", "reapproval", "version")):
        explanation = (
            f"On Slide {slide}, the match ({signal_text}) may relate to approval status, validity or version control. "
            "The review depends on the authoritative record for this specific material and version."
        )
        next_step = "Check the material ID/version and its current approval record, including validity or expiry details where applicable, before the material is used."
        context = "This is a document-control signal; verify it against the source record rather than treating the text match as a confirmed issue."
    elif any(x in identity for x in ("privacy", "personal data", "cybersecurity", "it security", "copyright", "consent", "permission", "image rights")):
        explanation = (
            f"On Slide {slide}, the match ({signal_text}) points to a possible privacy, security, consent or rights consideration. "
            "The actual requirement depends on the data, asset and way it is used."
        )
        next_step = "Check the relevant permission/consent and any required notice; involve the privacy, IT security or rights owner if clearance is unclear."
        context = "The keyword highlights an area to check, not a finding that a breach or rights violation has occurred."
    elif any(x in identity for x in ("reference", "citation", "source", "evidence", "substantiat", "claim", "efficacy", "safety", "rule-0001")):
        explanation = (
            f"On Slide {slide}, the rule matched {signal_text}. This may be part of a product-related or factual claim, "
            "so the important check is whether the actual statement is supported by the right source—not simply whether a reference is present."
        )
        next_step = "Trace the exact statement to its cited approved source. Confirm the source supports the same product, population, outcome and wording, and check that important qualifications are retained."
        context = "The matched terms are screening evidence only. Review the surrounding sentence and slide context before deciding what action is needed."
    else:
        explanation = (
            f"On Slide {slide}, the configured rule '{rule_name}' matched {signal_text}. "
            f"The significance depends on the full statement and the rule's intended scope{': ' + description if description else ''}."
        )
        next_step = "Review the matched wording in its full slide context, verify the applicable source or process record, and refer unresolved questions to the responsible reviewer/NS."
        context = "A rule match is a prompt for focused review, not a final compliance decision."

    # Add the classifier's view only when it contributes useful, non-duplicative context.
    if item:
        classification = str(item.get("classification", item.get("category", "")) or "").replace("_", " ").strip().lower()
        context_reason = str(item.get("reason", item.get("context", "")) or "").strip()
        if classification and classification not in {"uncertain", "not classified"}:
            context += f" Slide-level context was classified as {classification}; treat this as supporting context only."
        elif context_reason and "not enough evidence" in context_reason.lower():
            context += " Slide-level context is inconclusive, so the rule-specific evidence should be assessed on its own."

    # Prefer the more detailed evidence-layer recommendation if present; otherwise use the rule-specific next step.
    recommendation = str(finding.get("recommendation", "") or "").strip() or next_step
    return {
        "explanation": explanation,
        "context": context,
        "recommendation": recommendation,
        "suggestion": next_step,
        "risk": finding.get("severity", "Review"),
    }




def translate_ai_explanation(finding, ai):
    """Curated Thai localization for the full AI Explanation panel."""
    slide = finding.get("slide_number", finding.get("slide", "?"))
    signals = detected_values(finding)
    signal_text = ", ".join(signals[:5]) if signals else "เนื้อหาที่ตรงกับกฎ"
    identity = " ".join([
        str(finding.get("rule_id", "") or ""),
        str(finding.get("rule_name", finding.get("category", "")) or ""),
        " ".join(signals),
        str(finding.get("description", finding.get("rule_description", "")) or ""),
    ]).casefold()

    if any(x in identity for x in ("comparative", "more effective", "better than", "superior", "inferior", "compared")):
        explanation = f"ในสไลด์ {slide} พบถ้อยคำ “{signal_text}” ซึ่งอาจสื่อถึงการเปรียบเทียบหรือการกล่าวอ้างความเหนือกว่า ควรตรวจสอบว่าตัวเปรียบเทียบ ผลลัพธ์ และหลักฐานสนับสนุนสอดคล้องกับข้อความที่นำเสนอหรือไม่"
        context = "ระบบตรวจพบถ้อยคำที่ใช้ในการเปรียบเทียบ การพบถ้อยคำเพียงอย่างเดียวยังไม่ยืนยันว่าการเปรียบเทียบนั้นมีหลักฐานรองรับหรือเหมาะสม"
    elif any(x in identity for x in ("off-label", "indication", "dosage", "dose", "population", "route of administration", "prescribing information")):
        explanation = f"ในสไลด์ {slide} พบข้อความ “{signal_text}” ที่เกี่ยวข้องกับการใช้ผลิตภัณฑ์หรือข้อมูลการสั่งใช้ยา ความเหมาะสมต้องพิจารณาเทียบกับข้อมูลผลิตภัณฑ์ที่ได้รับอนุมัติในประเทศไทยฉบับปัจจุบันและบริบทของข้อความ"
        context = "ระบบระบุหัวข้อเกี่ยวกับการใช้ผลิตภัณฑ์ที่ควรตรวจสอบ ไม่ได้หมายความว่าพบการส่งเสริมการใช้ยานอกข้อบ่งใช้แล้ว"
    elif any(x in identity for x in ("approval", "approved material", "expiry", "expiration", "expired", "renewal", "re-approval", "reapproval", "version")):
        explanation = f"ในสไลด์ {slide} พบข้อความ “{signal_text}” ซึ่งอาจเกี่ยวข้องกับสถานะการอนุมัติ อายุการใช้งาน หรือการควบคุมเวอร์ชันของเอกสาร ควรตรวจสอบกับระเบียนทางการของเอกสารและเวอร์ชันนี้"
        context = "นี่เป็นสัญญาณให้ตรวจสอบการควบคุมเอกสาร ควรยืนยันกับแหล่งข้อมูลต้นทางก่อนสรุปว่ามีปัญหา"
    elif any(x in identity for x in ("privacy", "personal data", "cybersecurity", "it security", "copyright", "consent", "permission", "image rights")):
        explanation = f"ในสไลด์ {slide} พบข้อความ “{signal_text}” ที่อาจเกี่ยวข้องกับความเป็นส่วนตัว ความปลอดภัย ความยินยอม หรือสิทธิ์การใช้งาน ข้อกำหนดที่ต้องปฏิบัติตามขึ้นอยู่กับข้อมูล สื่อ และลักษณะการนำไปใช้จริง"
        context = "คำหรือวลีที่ตรวจพบเป็นเพียงจุดที่ควรตรวจสอบ ไม่ใช่ข้อสรุปว่ามีการละเมิดหรือเกิดเหตุด้านความปลอดภัยแล้ว"
    elif any(x in identity for x in ("reference", "citation", "source", "evidence", "substantiat", "claim", "efficacy", "safety", "rule-0001")):
        explanation = f"ในสไลด์ {slide} กฎตรวจพบคำหรือวลี “{signal_text}” ซึ่งอาจเป็นส่วนหนึ่งของข้อความกล่าวอ้างเกี่ยวกับผลิตภัณฑ์หรือข้อเท็จจริง ประเด็นสำคัญคือข้อความนั้นได้รับการสนับสนุนจากแหล่งอ้างอิงที่ถูกต้องหรือไม่ ไม่ใช่เพียงมีการใส่แหล่งอ้างอิงไว้หรือไม่"
        context = "คำที่ตรวจพบใช้เป็นหลักฐานเบื้องต้นสำหรับคัดกรองเท่านั้น ควรอ่านประโยคโดยรอบและพิจารณาบริบทของสไลด์ก่อนตัดสินใจดำเนินการ"
    else:
        rule_name = str(finding.get("rule_name", finding.get("category", "กฎที่ตั้งค่าไว้")) or "กฎที่ตั้งค่าไว้").strip()
        explanation = f"ในสไลด์ {slide} กฎ “{rule_name}” ตรงกับข้อความ “{signal_text}” ความสำคัญของข้อความนี้ขึ้นอยู่กับเนื้อหาโดยรวมและขอบเขตที่กฎดังกล่าวมีไว้ตรวจสอบ"
        context = "การตรงกับกฎเป็นเพียงสัญญาณให้ตรวจสอบเพิ่มเติม ไม่ใช่ข้อสรุปด้าน Compliance ขั้นสุดท้าย"
    return explanation, context

def translate_recommendation_suggestion(finding, recommendation, suggestion):
    """Return curated Thai text for recommendation/suggestion; no external translation service."""
    identity = " ".join([
        str(finding.get("rule_id", "") or ""),
        str(finding.get("rule_name", finding.get("category", "")) or ""),
        " ".join(detected_values(finding)),
        str(finding.get("description", finding.get("rule_description", "")) or ""),
    ]).casefold()

    if any(x in identity for x in ("comparative", "more effective", "better than", "superior", "inferior", "compared")):
        rec_th = "ตรวจสอบว่าข้อความเปรียบเทียบมีเอกสารอ้างอิงที่เหมาะสม และหลักฐานสนับสนุนข้อความ ผลลัพธ์ และผลิตภัณฑ์ที่กล่าวถึงตรงกันหรือไม่"
        sug_th = "1. ตรวจสอบข้อความเปรียบเทียบและตัวเปรียบเทียบที่ระบุ\n2. ตรวจสอบว่าหลักฐานรองรับผลลัพธ์และบริบทที่กล่าวอ้าง\n3. ให้ NS พิจารณาตามกระบวนการอนุมัติที่เกี่ยวข้อง"
    elif any(x in identity for x in ("off-label", "indication", "dosage", "dose", "population", "route of administration", "prescribing information")):
        rec_th = "ตรวจสอบข้อความกับข้อมูลผลิตภัณฑ์ที่ได้รับอนุมัติในประเทศไทยฉบับปัจจุบัน โดยพิจารณาข้อบ่งใช้ กลุ่มผู้ป่วย ขนาดยา และวิธีให้ยาตามความเกี่ยวข้อง"
        sug_th = "1. เทียบข้อความกับข้อมูลผลิตภัณฑ์ที่ได้รับอนุมัติฉบับปัจจุบัน\n2. ตรวจสอบข้อบ่งใช้ กลุ่มผู้ป่วย ขนาดยา หรือวิธีให้ยาในส่วนที่เกี่ยวข้อง\n3. หากพบความไม่สอดคล้อง ให้ส่งต่อผู้รับผิดชอบและ NS พิจารณา"
    elif any(x in identity for x in ("approval", "approved material", "expiry", "expiration", "expired", "renewal", "re-approval", "reapproval", "version")):
        rec_th = "ตรวจสอบรหัสเอกสาร รุ่น และสถานะการอนุมัติจากแหล่งข้อมูลทางการ รวมถึงวันหมดอายุหรือการต่ออายุถ้ามี"
        sug_th = "1. ยืนยันรหัสเอกสารและหมายเลขเวอร์ชัน\n2. ตรวจสอบบันทึกการอนุมัติและช่วงเวลาที่มีผล\n3. หากข้อมูลไม่ชัดเจน ให้ยืนยันกับเจ้าของกระบวนการก่อนนำเอกสารไปใช้"
    elif any(x in identity for x in ("privacy", "personal data", "cybersecurity", "it security", "copyright", "consent", "permission", "image rights")):
        rec_th = "ตรวจสอบสิทธิ์การใช้ข้อมูลหรือสื่อ ความยินยอม และข้อกำหนดด้านความเป็นส่วนตัวหรือความปลอดภัยที่เกี่ยวข้อง"
        sug_th = "1. ตรวจสอบว่ามีความยินยอมหรือสิทธิ์การใช้งานที่จำเป็น\n2. ตรวจสอบประกาศหรือเงื่อนไขการใช้ข้อมูลที่เกี่ยวข้อง\n3. หากยังไม่แน่ใจ ให้ปรึกษาทีม Privacy, IT Security หรือเจ้าของสิทธิ์ที่เกี่ยวข้อง"
    elif any(x in identity for x in ("reference", "citation", "source", "evidence", "substantiat", "claim", "efficacy", "safety", "rule-0001")):
        rec_th = "ตรวจสอบว่าข้อความกล่าวอ้างมีแหล่งอ้างอิงที่เหมาะสม และหลักฐานสนับสนุนผลิตภัณฑ์ กลุ่มผู้ป่วย ผลลัพธ์ และถ้อยคำที่ใช้จริง"
        sug_th = "1. เชื่อมโยงข้อความที่กล่าวอ้างกับแหล่งอ้างอิงโดยตรง\n2. ตรวจสอบว่าหลักฐานตรงกับผลิตภัณฑ์ กลุ่มผู้ป่วย และผลลัพธ์ที่ระบุ\n3. ตรวจสอบว่าข้อจำกัดหรือเงื่อนไขสำคัญยังคงอยู่ และให้ NS พิจารณา"
    else:
        rec_th = "ตรวจสอบข้อความที่ระบบระบุในบริบทของสไลด์ทั้งหมด และยืนยันกับเอกสารอ้างอิงหรือกระบวนการที่เกี่ยวข้อง"
        sug_th = "1. อ่านข้อความรอบคำหรือวลีที่ตรวจพบให้ครบถ้วน\n2. ตรวจสอบเอกสารอ้างอิงหรือบันทึกกระบวนการที่เกี่ยวข้อง\n3. ส่งคำถามที่ยังไม่ชัดเจนให้ผู้รับผิดชอบหรือ NS พิจารณา"

    # A high-level severity cue is retained without turning a rule match into a compliance conclusion.
    severity = str(finding.get("severity", "")).casefold()
    if severity == "high":
        rec_th = "ควรตรวจสอบก่อนดำเนินการต่อ: " + rec_th
    return rec_th, sug_th

def render_app_chrome():
    """Premium AstraZeneca glass product chrome used across both pages."""
    render_html(
        '<div class="msa-topbar">'
        '<div class="msa-brand-lockup">'
        '<div class="msa-brand-mark">⌕</div>'
        '<div><div class="msa-brand-name">Material Scanner Agent</div>'
        '<div class="msa-brand-subtitle">AstraZeneca · Compliance Suite</div></div>'
        '</div>'
        '<div class="msa-topbar-meta">AZ Hackathon 2026 <span class="msa-avatar">NS</span></div>'
        '</div>'
    )


def render_application_nav():
    """Render the Application switcher as a custom glass button group.

    This intentionally uses native Streamlit buttons rather than segmented_control
    so the visual treatment is stable across Streamlit versions.
    """
    current = st.session_state.page

    with st.container(key="application-nav-shell"):
        label_col, nav_col = st.columns([0.30, 2.15], vertical_alignment="center")

        with label_col:
            render_html('<div class="msa-application-label msa-application-label-inline">APPLICATION</div>')

        with nav_col:
            with st.container(key="application-pill"):
                review_col, rules_col = st.columns([1, 1.72], gap="small")

                with review_col:
                    review_clicked = st.button(
                        "⌕  Review",
                        key="application_review_button",
                        type="primary" if current == "Review" else "secondary",
                        use_container_width=True,
                    )

                with rules_col:
                    rules_clicked = st.button(
                        "♙  Rule Management",
                        key="application_rules_button",
                        type="primary" if current == "Rule Management" else "secondary",
                        use_container_width=True,
                    )

    if review_clicked and current != "Review":
        # Lock Rule Management immediately when leaving it.
        # Do this before changing the page because the navigation handler
        # updates st.session_state.page directly.
        if current == "Rule Management":
            st.session_state.rule_admin_authenticated = False
            st.session_state.rule_admin_login_error = ""
            st.session_state.rule_admin_login_version = (
                st.session_state.get("rule_admin_login_version", 0) + 1
            )
            st.session_state.rule_management_flash = None
        st.session_state.page = "Review"
        st.rerun()

    if rules_clicked and current != "Rule Management":
        # Every entry into Rule Management must pass the password gate.
        st.session_state.rule_admin_authenticated = False
        st.session_state.rule_admin_login_error = ""
        st.session_state.page = "Rule Management"
        st.rerun()

    return st.session_state.page


def render_hero():
    render_html(
        '<div class="msa-hero msa-hero-glass">'
        '<div class="msa-hero-overlay">'
        '<div class="msa-hero-eyebrow"><span class="msa-hero-dot"></span>ASTRAZENECA · COMPLIANCE SUITE</div>'
        f'<div class="msa-hero-title">{esc(APP_NAME)}</div>'
        '<div class="msa-hero-subtitle">AI-powered pre-review for promotional and non-promotional materials — helping teams ship compliant content, faster.</div>'
        '<div class="msa-hero-meta"><span><i></i>System ready</span><b>·</b><span>v1.0 · Prototype</span><b>·</b><span>AZ Hackathon 2026</span></div>'
        '</div>'
        '</div>'
    )


def header(icon, title, subtitle=""):
    render_html(
        f'<div class="msa-section-heading"><div class="msa-section-icon">{esc(icon)}</div>'
        f'<div><div class="msa-section-title">{esc(title)}</div>'
        f'<div class="msa-section-subtitle">{esc(subtitle)}</div></div></div>'
    )


# ------------------------------------------------------------
# UI: findings / evidence
# ------------------------------------------------------------
def render_signal_chips(values):
    if not values:
        return '<span class="msa-empty">None detected</span>'
    return "".join(
        f'<span class="msa-chip msa-chip-code">{esc(v)}</span>' for v in values
    )


def render_finding_card(finding, evidence, ai_result, rule_index=1):
    """Render one rule finding in a structured, two-column reviewer layout."""
    slide = finding.get("slide_number", finding.get("slide", "?"))
    rule_id = finding.get("rule_id", "")
    rule_name = finding.get("rule_name", finding.get("category", "Finding"))
    severity = finding.get("severity", "Review")
    action = finding.get("action", "FLAG_FOR_REVIEW")
    signals = detected_values(finding)

    evidence_matches = evidence_for_finding(evidence, finding)
    evidence_item = evidence_matches[0] if evidence_matches else {}
    evidence_values = evidence_item.get(
        "detected_evidence",
        evidence_item.get("matched_keywords", signals),
    )
    if isinstance(evidence_values, str):
        evidence_values = [evidence_values]
    evidence_values = [str(v) for v in (evidence_values or []) if str(v).strip()]

    why = evidence_item.get(
        "why_flagged",
        finding.get(
            "why_flagged",
            finding.get("reason", "The configured review rule identified content that should be reviewed."),
        ),
    )
    recommendation = evidence_item.get(
        "recommendation",
        finding.get(
            "recommendation",
            "Review the detected content and confirm that it is supported by the applicable reference.",
        ),
    )
    rule_description = clean_rule_description(
        evidence_item.get(
            "rule_description",
            finding.get("rule_description", finding.get("description", "")),
        )
    )

    ai_finding = dict(finding)
    ai_finding["recommendation"] = recommendation
    ai_finding["description"] = rule_description or ai_finding.get("description", "")
    ai = build_ai_explanation(ai_finding, ai_result)
    # Read the single global language selection placed below the result navigator.
    language = st.session_state.get("result_language", "English")
    recommendation_display = recommendation
    suggestion_display = ai.get("suggestion", ai.get("recommendation", ""))
    if language == "ไทย":
        recommendation_display, suggestion_display = translate_recommendation_suggestion(
            ai_finding, recommendation, suggestion_display
        )
    ai_available = bool(ai.get("explanation") or ai.get("context") or ai.get("recommendation"))
    ai_panel = ""
    is_thai = language == "ไทย"
    explanation_display = ai.get("explanation", "")
    context_display = ai.get("context", "")
    if is_thai:
        explanation_display, context_display = translate_ai_explanation(ai_finding, ai)
    explanation_label = "คำอธิบายประกอบ" if is_thai else "AI-ASSISTED EXPLANATION"
    suggestion_label = "ข้อเสนอแนะในการดำเนินการ" if is_thai else "Suggested Review Action"
    ai_title = "ดูคำอธิบายประกอบ" if is_thai else "View AI Explanation"
    ai_meta = "คำอธิบายประกอบเพื่อช่วยพิจารณาเท่านั้น" if is_thai else "AI ASSISTED · Contextual support only"
    context_label = "บริบทที่ตรวจพบ" if is_thai else "Context Signal"
    ai_explanation_title = "ข้อมูลเพิ่มเติมเพื่อประกอบการพิจารณา" if is_thai else "Additional context for human review"
    disclaimer = ("ระบบให้ข้อมูลประกอบการพิจารณาเท่านั้น การตัดสินใจด้าน Compliance ขั้นสุดท้ายเป็นหน้าที่ของ Nominated Signatory"
                  if is_thai else "AI provides contextual support only. Final compliance decision remains with the Nominated Signatory.")
    finding_label = "สิ่งที่ตรวจพบ" if is_thai else "Finding"
    evidence_label = "หลักฐานที่ตรวจพบ" if is_thai else "Evidence"
    why_label = "เหตุผลที่ระบบระบุรายการนี้" if is_thai else "Why was this flagged?"
    rule_note_label = "หมายเหตุของกฎ" if is_thai else "Rule note"
    detected_source = f"ตรวจพบในเอกสารที่อัปโหลด · สไลด์ {esc(slide)}" if is_thai else f"Detected in uploaded material · Slide {esc(slide)}"
    finding_content = (f'เนื้อหาในสไลด์ {esc(slide)} ตรงกับกฎตรวจสอบ <strong>{esc(rule_id or rule_name)}</strong>.'
                       if is_thai else f'Content on Slide {esc(slide)} matched the configured review rule <strong>{esc(rule_id or rule_name)}</strong>.')
    why_fallback = "ระบบพบเนื้อหาที่ควรตรวจสอบตามกฎที่ตั้งค่าไว้" if is_thai else "The configured review rule identified content that should be reviewed."
    if ai_available:
        ai_panel = (
            f'<details class="msa-ai-disclosure msa-finding-ai">'
            f'<summary>'
            f'<span class="msa-ai-disclosure-title"><span class="msa-ai-sparkle" aria-hidden="true">✧</span>{esc(ai_title)}</span>'
            f'<span class="msa-ai-disclosure-meta">{esc(ai_meta)} <span class="msa-ai-chevron" aria-hidden="true">⌄</span></span>'
            f'</summary>'
            f'<div class="msa-ai-explanation-card">'
            f'<div class="msa-ai-explanation-top"><div class="msa-ai-field-label">{esc(explanation_label)}</div>'
            f'<span class="msa-status-pill msa-status-info">{("มี AI ช่วยวิเคราะห์" if is_thai else "AI ASSISTED")}</span></div>'
            f'<div class="msa-ai-explanation-title">{esc(ai_explanation_title)}</div>'
            f'<div class="msa-ai-explanation-text">{esc(explanation_display)}</div>'
            f'<div class="msa-ai-explanation-context"><div class="msa-ai-field-label">{esc(context_label)}</div>{esc(context_display)}</div>'
            f'<div class="msa-ai-explanation-context"><div class="msa-ai-field-label">{esc(suggestion_label)}</div>{esc(suggestion_display)}</div>'
            f'<div class="msa-ai-disclaimer">{esc(disclaimer)}</div>'
            f'</div></details>'
        )

    def row(icon, label, content, extra_class=""):
        return (
            f'<div class="msa-finding-row {extra_class}">'
            f'<div class="msa-finding-row-label">'
            f'<span class="msa-finding-row-icon" aria-hidden="true">{icon}</span>'
            f'<span>{label}</span></div>'
            f'<div class="msa-finding-row-content">{content}</div>'
            f'</div>'
        )

    evidence_content = (
        f'<div class="msa-evidence-quote">{render_signal_chips(evidence_values)}'
        f'<div class="msa-evidence-source">{detected_source}</div>'
        f'</div>'
    )
    if is_thai:
        why_text = "ระบบพบข้อความที่ตรงกับเงื่อนไขของกฎตรวจสอบ ควรพิจารณาข้อความและบริบทโดยรอบก่อนสรุปผล"
        if why and any(token in str(why).casefold() for token in ("reference", "citation", "evidence", "claim")):
            why_text = "ข้อความที่ตรวจพบอาจต้องตรวจสอบแหล่งอ้างอิงและหลักฐานสนับสนุนให้ตรงกับสิ่งที่กล่าวอ้าง"
        why_content = esc(why_text)
    else:
        why_content = esc(why or "The configured review rule identified content that should be reviewed.")
    recommendation_content = f'<strong>{esc(recommendation_display)}</strong>'
    recommendation_label = "คำแนะนำ (Recommendation)" if language == "ไทย" else "Recommendation"
    severity_display = {"high": "สูง", "medium": "ปานกลาง", "low": "ต่ำ", "review": "ควรตรวจสอบ", "info": "ข้อมูล"}.get(str(severity).strip().lower(), str(severity)) if is_thai else str(severity)
    action_display = "ส่งให้ตรวจสอบ" if is_thai and str(action).upper() == "FLAG_FOR_REVIEW" else str(action)
    rule_id_label = "รหัสกฎ" if is_thai else "Rule ID"

    rule_note = ""
    if rule_description:
        rule_note = (
            f'<div class="msa-finding-rule-note">'
            f'<span class="msa-finding-rule-note-label">{esc(rule_note_label)}</span>'
            f'<span>{esc(rule_description)}</span></div>'
        )

    return (
        f'<article class="msa-finding-card msa-finding-layout {status_class(severity)}">'
        f'<header class="msa-finding-top {status_class(severity)}">'
        f'<span class="msa-finding-number">{int(rule_index)}</span>'
        f'<div class="msa-finding-heading">'
        f'<div class="msa-finding-title">{esc(rule_name)}</div>'
        f'<div class="msa-finding-meta">{esc(rule_id_label)}: {esc(rule_id)} <span class="msa-meta-dot">·</span> {esc(action_display)}</div>'
        f'</div>'
        f'<span class="msa-status-pill {status_class(severity)}">{esc(severity_display)}</span>'
        f'</header>'
        f'<div class="msa-finding-rows">'
        f'{row("⌕", finding_label, finding_content)}'
        f'{row("▤", evidence_label, evidence_content, "msa-finding-evidence-row")}'
        f'{rule_note}'
        f'{row("▱", why_label, why_content)}'
        f'{row("♧", recommendation_label, recommendation_content, "msa-finding-recommendation-row")}'
        f'</div>'
        f'{ai_panel}'
        f'</article>'
    )


def finding_slide_key(finding):
    """Return a stable slide key for grouping rule findings."""
    value = finding.get("slide_number", finding.get("slide", "?"))
    return str(value if value not in (None, "") else "?")


def slide_group_severity(items):
    """Return the highest severity present on a slide without changing rule results."""
    rank = {"high": 3, "medium": 2, "review": 2, "low": 1, "info": 0}
    def score(item):
        return rank.get(str(item.get("severity", "Review")).strip().lower(), 0)
    highest = max(items, key=score)
    return highest.get("severity", "Review")


def render_slide_findings_group(slide_key, items, evidence, ai_result):
    """Render a collapsed-by-default slide group, keeping individual rule findings intact."""
    severity = slide_group_severity(items)
    rule_count = len(items)
    is_thai = st.session_state.get("result_language", "English") == "ไทย"
    rule_label = ("กฎ" if is_thai else ("rule" if rule_count == 1 else "rules"))
    child_cards = "".join(
        render_finding_card(item, evidence, ai_result, rule_index=index)
        for index, item in enumerate(items, start=1)
    )
    return (
        f'<details class="msa-slide-group {status_class(severity)}">'
        f'<summary class="msa-slide-group-header">'
        f'<span class="msa-slide-group-summary-icon" aria-hidden="true">'
        f'<svg viewBox="0 0 24 24" focusable="false"><path d="M7 9.5l5 5 5-5" /></svg>'
        f'</span>'
        f'<span class="msa-slide-group-heading">'
        f'<span class="msa-slide-group-title">{("สไลด์ " if is_thai else "Slide ")}{esc(slide_key)}</span>'
        f'<span class="msa-slide-group-meta">{(str(rule_count) + " " + rule_label + " ที่ตรวจพบ · คลิกเพื่อดูรายละเอียด") if is_thai else (str(rule_count) + " " + rule_label + " detected · Click to expand findings")}</span>'
        f'</span>'
        f'<span class="msa-slide-group-badges">'
        f'<span class="msa-slide-group-count">{rule_count} {rule_label}</span>'
        f'<span class="msa-status-pill {status_class(severity)}">{("ระดับสูงสุด: " if is_thai else "Highest: ")}{esc({"high":"สูง","medium":"ปานกลาง","low":"ต่ำ","review":"ควรตรวจสอบ","info":"ข้อมูล"}.get(str(severity).strip().lower(), str(severity)) if is_thai else severity)}</span>'
        f'</span>'
        f'</summary>'
        f'<div class="msa-slide-group-body">{child_cards}</div>'
        f'</details>'
    )


def render_evidence_card(item):
    slide = item.get("slide_number", item.get("slide", "?"))
    rule_id = item.get("rule_id", "")
    rule_name = item.get("rule_name", item.get("category", "Rule"))
    severity = item.get("severity", "Review")
    detected = item.get("detected_evidence", item.get("matched_keywords", item.get("keywords", [])))
    if isinstance(detected, str):
        detected = [detected]
    detected = [str(x) for x in (detected or []) if str(x).strip()]

    rule_description = clean_rule_description(item.get("rule_description", item.get("description", "")))
    why = item.get("why_flagged", item.get("reason", ""))
    recommendation = item.get(
        "recommendation",
        "Review the detected content and confirm that it is supported by the applicable reference.",
    )

    render_html(
        f'<div class="msa-evidence-card">'
        f'<div class="msa-evidence-top"><div><div class="msa-evidence-slide">Slide {esc(slide)}</div>'
        f'<div class="msa-finding-meta">{esc(rule_id)} · {esc(rule_name)}</div></div>'
        f'<span class="msa-status-pill {status_class(severity)}">{esc(severity)}</span></div>'
        f'<div class="msa-evidence-section"><div class="msa-evidence-label">Detected Evidence</div>'
        f'<div class="msa-evidence-quote">{render_signal_chips(detected)}</div></div>'
        f'<div class="msa-evidence-section"><div class="msa-evidence-label">Applicable Rule</div>'
        f'<div class="msa-rule-reference"><strong>{esc(rule_id)}</strong> · {esc(rule_name)}</div></div>'
        f'<div class="msa-evidence-section"><div class="msa-evidence-label">Rule Description</div><div>{esc(rule_description or "Configured review rule.")}</div></div>'
        f'<div class="msa-evidence-section"><div class="msa-evidence-label">Why was this flagged?</div><div>{esc(why or "The rule matched content in the material.")}</div></div>'
        f'<div class="msa-evidence-section msa-evidence-recommendation"><div class="msa-evidence-label">Recommended Action</div><strong>{esc(recommendation)}</strong></div>'
        f'</div>'
    )


# ------------------------------------------------------------
# Rule Management Access Control
# ------------------------------------------------------------
def get_rule_management_password():
    """Return the Rule Management administrator password.

    Prototype default: msa1234
    A Streamlit secret or environment variable can override it later.
    """
    try:
        secret_password = st.secrets.get("RULE_MANAGEMENT_PASSWORD")
        if secret_password:
            return str(secret_password)
    except Exception:
        pass

    env_password = os.getenv("RULE_MANAGEMENT_PASSWORD")
    if env_password:
        return str(env_password)

    return "msa1234"


def _password_matches(entered, configured):
    """Compare passwords safely, including non-ASCII input/configuration."""
    entered_bytes = str(entered or "").encode("utf-8")
    configured_bytes = str(configured or "").encode("utf-8")
    return hmac.compare_digest(entered_bytes, configured_bytes)


def set_rule_management_flash(message, level="success"):
    """Persist one Rule Management result message across st.rerun()."""
    st.session_state.rule_management_flash = {
        "message": str(message),
        "level": str(level),
    }


def show_rule_management_flash():
    """Display and consume a one-time Rule Management result message."""
    flash = st.session_state.get("rule_management_flash")
    if not flash:
        return

    message = str(flash.get("message", ""))
    level = str(flash.get("level", "success")).lower()

    if level == "error":
        st.error(message)
    elif level == "warning":
        st.warning(message)
    else:
        st.success(message)

    st.session_state.rule_management_flash = None


def rule_management_auth_gate():
    """Require a password before allowing access to Rule Management."""
    if st.session_state.get("rule_admin_authenticated", False):
        return True

    header(
        "🔐",
        "Rule Management Access",
        "Restricted area. Enter the administrator password to manage screening rules.",
    )

    configured_password = get_rule_management_password()
    login_version = st.session_state.get("rule_admin_login_version", 0)
    password_key = f"rule_admin_password_{login_version}"

    with st.form(f"rule_management_login_{login_version}", clear_on_submit=False):
        password = st.text_input(
            "Administrator Password",
            type="password",
            placeholder="Enter password",
            key=password_key,
        )
        submitted = st.form_submit_button(
            "🔓 Unlock Rule Management",
            type="primary",
            use_container_width=True,
        )

    if submitted:
        if _password_matches(password, configured_password):
            st.session_state.rule_admin_authenticated = True
            st.session_state.rule_admin_login_error = ""
            st.rerun()
        else:
            # Keep the message deliberately simple; do not expose any
            # password/configuration details.
            st.session_state.rule_admin_login_error = "Incorrect password."

    if st.session_state.get("rule_admin_login_error"):
        st.error(st.session_state.rule_admin_login_error)
        if st.button(
            "↻ Re-type Password",
            key=f"rule_admin_retype_{login_version}",
            use_container_width=False,
        ):
            st.session_state.rule_admin_login_error = ""
            st.session_state.rule_admin_login_version = login_version + 1
            st.rerun()

    return False



def read_xlsx_sheet_rows(file_bytes, target_sheet_name="Rule Catalog"):
    """Read rows from a worksheet in an .xlsx workbook using only Python stdlib.

    This keeps Rule Catalog import usable in locked-down environments where
    optional packages such as openpyxl are not installed.
    """
    ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ns_rel_doc = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ns_rel_pkg = "http://schemas.openxmlformats.org/package/2006/relationships"
    ns = {"main": ns_main, "r": ns_rel_doc, "rel": ns_rel_pkg}

    with zipfile.ZipFile(BytesIO(file_bytes)) as archive:
        names = set(archive.namelist())
        if "xl/workbook.xml" not in names:
            raise ValueError("This file is not a valid Excel .xlsx workbook.")

        shared_strings = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall("main:si", ns):
                shared_strings.append("".join(
                    node.text or "" for node in item.findall(".//main:t", ns)
                ))

        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        sheet_target = None
        for sheet in workbook_root.findall("main:sheets/main:sheet", ns):
            if sheet.attrib.get("name") == target_sheet_name:
                rel_id = sheet.attrib.get("{" + ns_rel_doc + "}id")
                sheet_target = rel_id
                break
        if not sheet_target:
            raise ValueError(f"Workbook must contain an '{target_sheet_name}' sheet.")

        rels_path = "xl/_rels/workbook.xml.rels"
        if rels_path not in names:
            raise ValueError("Excel workbook is missing worksheet relationships.")
        rels_root = ET.fromstring(archive.read(rels_path))
        target_path = None
        for rel in rels_root.findall("rel:Relationship", ns):
            if rel.attrib.get("Id") == sheet_target:
                target_path = rel.attrib.get("Target")
                break
        if not target_path:
            raise ValueError(f"Could not locate the '{target_sheet_name}' worksheet.")

        if target_path.startswith("/"):
            worksheet_path = target_path.lstrip("/")
        else:
            worksheet_path = posixpath.normpath(posixpath.join("xl", target_path))
        if worksheet_path not in names:
            raise ValueError(f"Worksheet file not found: {target_sheet_name}")

        worksheet_root = ET.fromstring(archive.read(worksheet_path))
        rows_out = []
        sheet_data = worksheet_root.find("main:sheetData", ns)
        if sheet_data is None:
            return rows_out

        def column_index(cell_ref):
            letters = "".join(ch for ch in cell_ref if ch.isalpha())
            index = 0
            for ch in letters.upper():
                index = index * 26 + ord(ch) - ord("A") + 1
            return max(index - 1, 0)

        for row in sheet_data.findall("main:row", ns):
            row_values = []
            for cell in row.findall("main:c", ns):
                ref = cell.attrib.get("r", "A1")
                col = column_index(ref)
                while len(row_values) <= col:
                    row_values.append(None)

                cell_type = cell.attrib.get("t", "")
                value_node = cell.find("main:v", ns)
                value = value_node.text if value_node is not None else None

                if cell_type == "s" and value is not None:
                    try:
                        value = shared_strings[int(value)]
                    except (ValueError, IndexError):
                        value = ""
                elif cell_type == "inlineStr":
                    value = "".join(
                        node.text or "" for node in cell.findall(".//main:t", ns)
                    )
                elif cell_type == "b" and value is not None:
                    value = value == "1"

                row_values[col] = value

            rows_out.append(tuple(row_values))

        # Normalize row width to the widest row for predictable dictionary mapping.
        width = max((len(row) for row in rows_out), default=0)
        return [tuple(list(row) + [None] * (width - len(row))) for row in rows_out]


def rule_management_page():
    if not rule_management_auth_gate():
        return

    show_rule_management_flash()

    header("⚙️", "Rule Management", "Manage screening rules independently from the review application.")

    lock_col, _ = st.columns([1, 5])
    with lock_col:
        if st.button("🔒 Lock Rule Management", use_container_width=True, key="lock_rule_management"):
            st.session_state.rule_admin_authenticated = False
            st.rerun()
    header("⚙️", "Rule Management", "Manage screening rules independently from the review application.")

    if load_rules is None:
        st.error("Rule Management module could not be loaded.")
        return

    with st.expander("Import Rule Catalog (.xlsx)", expanded=False):
        st.caption(
            "Upload the NS Rule Management workbook. MSA reads the 'Rule Catalog' sheet directly. "
            "Only approved rules explicitly marked Active can become active in the engine."
        )
        catalog_file = st.file_uploader(
            "Upload MSA Rule Management workbook",
            type=["xlsx"],
            key="rm_catalog_upload",
        )
        if st.button(
            "Import Rule Catalog",
            key="rm_import_catalog",
            type="primary",
            disabled=(catalog_file is None or upsert_rules is None),
        ):
            try:
                worksheet_rows = read_xlsx_sheet_rows(catalog_file.getvalue(), "Rule Catalog")
                if not worksheet_rows:
                    raise ValueError("The 'Rule Catalog' sheet is empty.")

                def _header_key(value):
                    return re.sub(r"[^a-z0-9]+", " ", str(value or "").strip().lower()).strip()

                headers = [_header_key(value) for value in worksheet_rows[0]]
                if not headers or "rule id" not in headers:
                    raise ValueError("Rule Catalog is missing the required 'Rule ID' column.")

                def _get(record, *names):
                    for name in names:
                        value = record.get(_header_key(name))
                        if value is not None and str(value).strip():
                            return value
                    return ""

                imported = []
                seen_ids = set()
                for row_number, values in enumerate(worksheet_rows[1:], start=2):
                    record = {headers[i]: values[i] if i < len(values) else None for i in range(len(headers))}
                    raw_id = _get(record, "Rule ID", "rule_id")
                    if not str(raw_id or "").strip():
                        if any(str(v or "").strip() for v in values):
                            raise ValueError(f"Rule Catalog row {row_number}: Rule ID is required.")
                        continue

                    rule_id = str(raw_id).strip()
                    if rule_id in seen_ids:
                        raise ValueError(f"Rule Catalog row {row_number}: duplicate Rule ID '{rule_id}'.")
                    seen_ids.add(rule_id)

                    raw_applies = str(_get(record, "Applies To", "applies_to") or "").strip()
                    applies = [x.strip().upper().replace("-", "_") for x in re.split(r"[;,|]", raw_applies) if x.strip()]
                    if not applies:
                        raise ValueError(f"Rule Catalog row {row_number} ({rule_id}): Applies To is required.")
                    if any(x not in {"PROMOMAT", "NON_PROMOMAT"} for x in applies):
                        raise ValueError(f"Rule Catalog row {row_number} ({rule_id}): Applies To must contain PROMOMAT and/or NON-PROMOMAT.")
                    applies = sorted(set(applies))

                    raw_keywords = _get(record, "Keywords / Phrases (semicolon-separated)", "Keywords / Phrases", "Keywords", "keywords")
                    keywords = [x.strip() for x in re.split(r"[;\r\n]+", str(raw_keywords or "")) if x.strip()]
                    approval = str(_get(record, "Approval Status", "approval_status") or "PENDING_NS_SOP_OWNER_REVIEW").strip().upper()
                    allowed = {"PENDING_NS_SOP_OWNER_REVIEW", "PENDING", "APPROVED", "NS_APPROVED", "APPROVED_BY_NS", "APPROVED_BY_SOP_OWNER"}
                    if approval not in allowed:
                        raise ValueError(f"Rule Catalog row {row_number} ({rule_id}): unsupported Approval Status '{approval}'.")
                    approved = approval in {"APPROVED", "NS_APPROVED", "APPROVED_BY_NS", "APPROVED_BY_SOP_OWNER"}
                    raw_status = str(_get(record, "Operational Status", "Status", "status") or "INACTIVE").strip().upper()
                    requested_active = raw_status in {"ACTIVE", "APPROVED AND ACTIVE"}
                    active = requested_active and approved
                    if active and not keywords:
                        raise ValueError(f"Rule Catalog row {row_number} ({rule_id}): add at least one value to 'Keywords / Phrases (semicolon-separated)' before activating this rule.")

                    raw_type = str(_get(record, "Rule Type", "rule_type") or "KEYWORD").strip().upper()
                    if "WORKFLOW" in raw_type or "CONTEXT" in raw_type or "HUMAN_REVIEW" in raw_type:
                        runtime_type = "context"
                    elif "PHRASE" in raw_type:
                        runtime_type = "phrase"
                    else:
                        runtime_type = "keyword"
                    matching_mode = str(_get(record, "Matching Mode", "matching_mode") or "").strip().upper()
                    if "PHRASE" in matching_mode:
                        runtime_type = "phrase"

                    category = applies[0] if len(applies) == 1 else "BOTH"
                    rule_name = str(_get(record, "Rule Name", "rule_name", "Topic") or rule_id).strip()
                    topic = str(_get(record, "Topic", "topic") or "").strip()
                    message = str(_get(record, "MSA Message", "Message", "message") or _get(record, "Source Requirement (paraphrased)") or rule_name).strip()
                    recommended_action = str(_get(record, "Recommended Action", "recommended_action") or "Route this finding to NS for human review.").strip()
                    imported.append({
                        "rule_id": rule_id,
                        "rule_name": rule_name,
                        "topic": topic,
                        "category": category,
                        "applies_to": applies,
                        "rule_type": runtime_type,
                        "matching_mode": matching_mode or runtime_type.upper(),
                        "keywords": keywords,
                        "context": message,
                        "severity": str(_get(record, "Severity (proposed; NS confirm)", "Severity", "severity") or "Medium").strip().title(),
                        "action": "CONTEXT_ONLY" if "WORKFLOW" in raw_type else "FLAG_FOR_REVIEW",
                        "description": recommended_action,
                        "source_requirement": str(_get(record, "Source Requirement (paraphrased)") or "").strip(),
                        "proposed_detection_logic": str(_get(record, "Proposed Detection Logic") or "").strip(),
                        "status": "Active" if active else "Inactive",
                        "approval_status": approval,
                        "version": str(_get(record, "Rule Version", "Version", "rule_version") or "0.1").strip(),
                        "source_document": str(_get(record, "Source Document", "source_document") or "").strip(),
                        "source_version_status": str(_get(record, "Source Version / Status", "source_version_status") or "").strip(),
                        "source_section_page": str(_get(record, "Source Section / Page (extracted)", "Source Section / Page", "source_section_page") or "").strip(),
                        "effective_date": str(_get(record, "Effective Date", "effective_date") or "").strip(),
                        "rule_owner": str(_get(record, "Rule Owner", "rule_owner") or "NS / SOP Owner").strip(),
                        "approved_by": str(_get(record, "Approved By", "approved_by") or "").strip(),
                        "approved_at": str(_get(record, "Approved At", "approved_at") or "").strip(),
                        "positive_test_case": str(_get(record, "Positive Test Case") or "").strip(),
                        "negative_test_case_limitation": str(_get(record, "Negative Test Case / Limitation") or "").strip(),
                        "notes_dependencies": str(_get(record, "Notes / Dependencies") or "").strip(),
                        "import_row": row_number,
                    })
                if not imported:
                    raise ValueError("No rules with a Rule ID were found in the Rule Catalog sheet.")
                result = upsert_rules(imported)
                active_count = sum(1 for item in imported if item["status"] == "Active")
                st.success(
                    f"Rule Catalog imported successfully — {result['upserted']} rules processed "
                    f"({result['created']} new, {result['updated']} updated). "
                    f"{active_count} rule(s) are Active based on the approval and operational "
                    f"status recorded in Excel."
                )
                # Keep the import panel compact. Detailed keyword diagnostics are omitted
                # from the main UI; imported rules remain visible in the Rule Management list.
            except Exception as exc:
                st.error(f"Import failed: {exc}")

    rules = load_rules()

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Rules", len(rules))
    with col2:
        st.metric("Active", sum(1 for r in rules if r.get("status") == "Active"))
    with col3:
        st.metric("Inactive", sum(1 for r in rules if r.get("status") == "Inactive"))

    st.divider()

    filter_col1, filter_col2 = st.columns(2)
    categories = sorted({r.get("category", "") for r in rules if r.get("category")})
    with filter_col1:
        category_filter = st.selectbox("Category", ["ALL"] + categories, key="rm_category")
    with filter_col2:
        status_filter = st.selectbox("Status", ["ALL", "Active", "Inactive"], key="rm_status")

    filtered = [
        r for r in rules
        if (category_filter == "ALL" or r.get("category") == category_filter)
        and (status_filter == "ALL" or r.get("status") == status_filter)
    ]

    st.subheader(f"Rules ({len(filtered)})")

    for rule in filtered:
        rid = rule.get("rule_id", "")
        status = rule.get("status", "Inactive")
        with st.container(border=True):
            c1, c2, c3, c4 = st.columns([1.2, 3, 1.5, 1.5])
            c1.write(f"**{rid}**")
            c2.write(f"**{rule.get('rule_name', '')}**")
            c3.write(rule.get("category", ""))
            c4.write("🟢 Active" if status == "Active" else "⚪ Inactive")

            with st.expander("View / Edit Rule"):
                name = st.text_input("Rule Name", rule.get("rule_name", ""), key=f"rm_name_{rid}")
                category_options = ["PROMOMAT", "NON_PROMOMAT"]
                current_category = rule.get("category", PROMOMAT)
                if current_category not in category_options:
                    category_options.append(current_category)
                category = st.selectbox("Category", category_options, index=category_options.index(current_category), key=f"rm_cat_{rid}")

                type_options = ["keyword", "phrase", "context"]
                current_type = rule.get("rule_type", "keyword")
                if current_type not in type_options:
                    type_options.append(current_type)
                rule_type = st.selectbox("Rule Type", type_options, index=type_options.index(current_type), key=f"rm_type_{rid}")

                keywords = st.text_area(
                    "Keywords / Phrases — one per line",
                    "\n".join(rule.get("keywords", [])),
                    key=f"rm_keywords_{rid}",
                )
                context = st.text_area("Context", rule.get("context", ""), key=f"rm_context_{rid}")

                ca, cb = st.columns(2)
                severity_options = ["Low", "Medium", "High"]
                current_severity = rule.get("severity", "Medium")
                if current_severity not in severity_options:
                    severity_options.append(current_severity)
                with ca:
                    severity = st.selectbox("Severity", severity_options, index=severity_options.index(current_severity), key=f"rm_severity_{rid}")

                action_options = ["FLAG_FOR_REVIEW", "CONTEXT_ONLY"]
                current_action = rule.get("action", "FLAG_FOR_REVIEW")
                if current_action not in action_options:
                    action_options.append(current_action)
                with cb:
                    action = st.selectbox("Action", action_options, index=action_options.index(current_action), key=f"rm_action_{rid}")

                description = st.text_area("Description", rule.get("description", ""), key=f"rm_desc_{rid}")
                approval_options = ["PENDING_NS_SOP_OWNER_REVIEW", "APPROVED"]
                current_approval = str(rule.get("approval_status", "")).strip().upper() or "APPROVED"
                if current_approval not in approval_options:
                    approval_options.append(current_approval)
                approval_status = st.selectbox(
                    "Approval Status", approval_options,
                    index=approval_options.index(current_approval),
                    key=f"rm_approval_{rid}",
                    help="Only explicitly approved rules may be enabled in the screening engine."
                )
                st.caption(f"Version: {rule.get('version', '1.0')}")

                save_col, status_col, delete_col = st.columns(3)

                with save_col:
                    if st.button("Save Changes", key=f"rm_save_{rid}", type="primary", use_container_width=True):
                        try:
                            update_rule(rid, {
                                "rule_name": name.strip(),
                                "category": category,
                                "rule_type": rule_type,
                                "keywords": [x.strip() for x in keywords.splitlines() if x.strip()],
                                "context": context.strip(),
                                "severity": severity,
                                "action": action,
                                "description": description.strip(),
                                "approval_status": approval_status,
                            })
                            set_rule_management_flash(f"{rid} saved successfully.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Unable to update rule: {e}")

                with status_col:
                    if status == "Active":
                        if st.button("Disable Rule", key=f"rm_disable_{rid}", use_container_width=True):
                            try:
                                disable_rule(rid)
                                set_rule_management_flash(f"{rid} disabled successfully.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Unable to disable rule: {e}")
                    else:
                        approval_value = str(rule.get("approval_status", "")).strip().upper()
                        approval_blocked = bool(approval_value) and approval_value not in {"APPROVED", "NS_APPROVED", "APPROVED_BY_NS", "APPROVED_BY_SOP_OWNER"}
                        if approval_blocked:
                            st.caption("Pending approval — save Approval Status as APPROVED before enabling.")
                        if st.button("Enable Rule", key=f"rm_enable_{rid}", use_container_width=True, disabled=approval_blocked):
                            try:
                                enable_rule(rid)
                                set_rule_management_flash(f"{rid} enabled successfully.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Unable to enable rule: {e}")

                with delete_col:
                    if st.button("Delete Rule", key=f"rm_delete_{rid}", use_container_width=True):
                        try:
                            delete_rule(rid)
                            set_rule_management_flash(f"{rid} deleted successfully.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Unable to delete rule: {e}")

    st.divider()
    with st.expander("＋ Add New Rule"):
        add_col1, add_col2 = st.columns(2)
        with add_col1:
            new_id = st.text_input("Rule ID", key="new_rule_id")
            new_name = st.text_input("Rule Name", key="new_rule_name")
            new_category = st.selectbox("Category", ["PROMOMAT", "NON_PROMOMAT"], key="new_rule_category")
            new_type = st.selectbox("Rule Type", ["keyword", "phrase", "context"], key="new_rule_type")
        with add_col2:
            new_severity = st.selectbox("Severity", ["Low", "Medium", "High"], index=1, key="new_rule_severity")
            new_action = st.selectbox("Action", ["FLAG_FOR_REVIEW", "CONTEXT_ONLY"], key="new_rule_action")
            new_version = st.text_input("Version", value="1.0", key="new_rule_version")

        new_keywords = st.text_area("Keywords / Phrases — one per line", key="new_rule_keywords")
        new_context = st.text_area("Context", key="new_rule_context")
        new_description = st.text_area("Description", value="Illustrative screening rule used for prototype demonstration.", key="new_rule_description")
        new_approval_status = st.selectbox("Approval Status", ["PENDING_NS_SOP_OWNER_REVIEW", "APPROVED"], key="new_rule_approval", help="Use APPROVED only after the appropriate NS/SOP Owner review.")

        if st.button("Add Rule", type="primary", key="add_rule"):
            try:
                if not new_id.strip():
                    raise ValueError("Rule ID is required.")
                if not new_name.strip():
                    raise ValueError("Rule Name is required.")
                if not new_keywords.strip():
                    raise ValueError("At least one keyword or phrase is required.")

                add_rule({
                    "rule_id": new_id.strip(),
                    "rule_name": new_name.strip(),
                    "category": new_category,
                    "rule_type": new_type,
                    "keywords": [x.strip() for x in new_keywords.splitlines() if x.strip()],
                    "context": new_context.strip(),
                    "severity": new_severity,
                    "action": new_action,
                    "status": "Inactive" if new_approval_status != "APPROVED" else "Active",
                    "approval_status": new_approval_status,
                    "applies_to": [new_category],
                    "description": new_description.strip(),
                    "version": new_version.strip() or "1.0",
                })
                set_rule_management_flash(f"{new_id.strip()} added successfully.")
                st.rerun()
            except Exception as e:
                st.error(f"Unable to add rule: {e}")


# ------------------------------------------------------------
# Review UI
# ------------------------------------------------------------
def render_results_nav():
    """Raised-button navigator shown after a review has completed."""
    options = ["Summary", "Findings", "Context Analysis", "AI Analysis", "Report"]
    current = st.session_state.get("result_section", "Summary")
    if current not in options:
        current = "Summary"

    with st.container(key="results-nav-shell"):
        cols = st.columns([1.15, 1.15, 1.55, 1.25, 1.0], gap="small")
        for index, value in enumerate(options):
            with cols[index]:
                is_active = value == current
                if st.button(
                    value,
                    key=f"result_nav_{value.lower().replace(' ', '_')}",
                    type="primary" if is_active else "secondary",
                    use_container_width=True,
                ):
                    if value != st.session_state.get("result_section"):
                        st.session_state.result_section = value
                        st.rerun()

    # One global language selector for all result tabs; do not render it inside cards.
    language_col, _ = st.columns([1.0, 5.0])
    with language_col:
        st.radio(
            "Display language",
            options=["English", "ไทย"],
            horizontal=True,
            key="result_language",
            label_visibility="collapsed",
        )


def review_page():
    # Review Profile is presented as one unified glass card.
    # Native Streamlit buttons are kept for reliable state handling.
    with st.container(key="review-profile-shell"):
        render_html(
            '<div class="msa-review-profile-heading">'
            '<div class="msa-review-profile-icon">◉</div>'
            '<div class="msa-review-profile-heading-copy">'
            '<div class="msa-review-profile-title">Review Profile</div>'
            '<div class="msa-review-profile-subtitle">Select the review profile that best matches your material.</div>'
            '</div>'
            '</div>'
        )

        current_profile = st.session_state.review_profile
        with st.container(key="review-profile-choices"):
            c1, c2, _ = st.columns([0.15, 0.23, 0.62], gap="small")
            with c1:
                promo_clicked = st.button(
                    "●  PROMOMAT",
                    key="review_profile_promomat_btn",
                    type="primary" if current_profile == PROMOMAT else "secondary",
                    use_container_width=False,
                )
            with c2:
                nonpromo_clicked = st.button(
                    "●  NON-PROMOMAT",
                    key="review_profile_nonpromomat_btn",
                    type="primary" if current_profile == NON_PROMOMAT else "secondary",
                    use_container_width=False,
                )

        selected_profile = current_profile
        if promo_clicked:
            selected_profile = PROMOMAT
        elif nonpromo_clicked:
            selected_profile = NON_PROMOMAT

        if selected_profile != current_profile:
            st.session_state.review_profile = selected_profile
            reset_review_state()
            st.rerun()

        if st.session_state.review_profile:
            render_html(
                '<div class="msa-profile-confirmed">'
                '<div class="msa-profile-confirmed-icon">✓</div>'
                '<div>'
                '<div class="msa-profile-confirmed-title">Profile confirmed</div>'
                f'<div class="msa-profile-confirmed-text">{esc(st.session_state.review_profile)} rules will be applied to this material.</div>'
                '</div>'
                '</div>'
            )
        else:
            render_html(
                '<div class="msa-profile-required">'
                'Review Profile is required. Select PROMOMAT or NON-PROMOMAT before starting the review.'
                '</div>'
            )

    with st.container(key="profile-upload-divider"):
        st.divider()

    # Upload Material is presented as a single glass card to match the reference UI.
    with st.container(key="upload-material-card"):
        render_html(
            '<div class="msa-upload-heading">'
            '<div class="msa-upload-icon">↥</div>'
            '<div class="msa-upload-heading-copy">'
            '<div class="msa-upload-title">Upload Material</div>'
            '<div class="msa-upload-subtitle">PowerPoint presentations up to 200MB. Processing begins on Start review.</div>'
            '</div>'
            '</div>'
        )

        render_html('<div class="msa-upload-shell">')
        uploaded = st.file_uploader(
            "Choose a PowerPoint presentation",
            type=["pptx"],
            key=f"material_uploader_{st.session_state.uploader_version}",
        )
        render_html('</div>')

        if uploaded is not None:
            data = uploaded.getvalue()
            digest = hashlib.sha256(data).hexdigest()
            if digest != st.session_state.uploaded_hash:
                st.session_state.uploaded_material = data
                st.session_state.uploaded_filename = uploaded.name
                st.session_state.uploaded_hash = digest
                reset_review_state(clear_material=False)

        has_material = st.session_state.uploaded_material is not None

        if has_material:
            try:
                preview = normalize_slides(extract_material(st.session_state.uploaded_material))
                render_html(
                    '<div class="msa-upload-file-summary">'
                    '<div class="msa-upload-file-icon">□</div>'
                    '<div class="msa-upload-file-copy">'
                    f'<div class="msa-upload-file-name">{esc(st.session_state.uploaded_filename)}</div>'
                    '<div class="msa-upload-file-meta">Uploaded successfully</div>'
                    '</div>'
                    '</div>'
                    f'<div class="msa-upload-status msa-upload-status-info">'
                    f'<div class="msa-upload-status-icon">i</div>'
                    f'<div><div class="msa-upload-status-label">{len(preview)} slides detected</div>'
                    f'<div class="msa-upload-status-text">Material parsed successfully. Ready for review.</div></div>'
                    '</div>'
                )
            except Exception as e:
                st.error(f"Unable to read the PowerPoint file: {e}")
        else:
            render_html(
                '<div class="msa-upload-empty-state">'
                '<div class="msa-upload-empty-icon">↑</div>'
                '<div><div class="msa-upload-empty-title">Upload a PowerPoint presentation</div>'
                '<div class="msa-upload-empty-text">PPTX files up to 200MB.</div></div>'
                '</div>'
            )

    with st.container(key="upload-start-divider"):
        st.divider()

    ready = st.session_state.review_profile is not None and has_material
    review_locked = bool(st.session_state.review_started)
    with st.container(key="start-review-action"):
        start = st.button(
            "▶  START REVIEW",
            use_container_width=True,
            type="primary",
            # Once a review has started, lock this CTA. The user must use
            # "Start New Review" to clear the current result and begin again.
            disabled=(not ready or review_locked),
            key="start_review",
        )

    if start:
        run_current_review()

    if st.session_state.review_started:
        render_results_nav()

    if not st.session_state.review_started:
        if not ready:
            if not st.session_state.review_profile and not has_material:
                st.caption("Select a Review Profile and upload a PowerPoint to enable Start Review.")
            elif not st.session_state.review_profile:
                st.caption("Select a Review Profile to enable Start Review.")
            else:
                st.caption("Upload a PowerPoint presentation to enable Start Review.")
        return

    render_results()


def run_current_review():
    try:
        with st.spinner("Running Material Scanner Agent..."):
            slides = normalize_slides(extract_material(st.session_state.uploaded_material))
            profile = ENGINE_PROFILE.get(st.session_state.review_profile)
            if not profile:
                raise ValueError("A valid Review Profile must be selected.")

            analysis = run_engine(slides, profile)
            evidence = run_evidence(analysis)
            ai = run_ai(slides)
            c = counts(analysis)
            findings = unique_findings(analysis)

            # Review status is determined by configured rule findings, not by
            # AI/context analysis. AI provides supporting interpretation only.
            has_action_required = any(
                str(f.get("action", "")).upper() == "ACTION_REQUIRED"
                or str(f.get("severity", "")).lower() == "high"
                for f in findings
                if isinstance(f, dict)
            )
            if has_action_required:
                status = "ACTION REQUIRED"
            elif findings:
                status = "REVIEW REQUIRED"
            else:
                status = "PASS"

            st.session_state.slides = slides
            st.session_state.analysis_result = analysis
            st.session_state.evidence_result = evidence
            st.session_state.ai_result = ai
            profile_mapping = analysis.get("profile_mapping", get_profile_mapping(st.session_state.review_profile))
            st.session_state.review_result = {
                "status": status,
                "slides": c["slides"] or len(slides),
                "findings": len(findings),
                "high": c["high"],
                "medium": c["medium"],
                "low": c["low"],
                "profile": st.session_state.review_profile,
                "active_rules": int(profile_mapping.get("rule_count", 0)) if isinstance(profile_mapping, dict) else 0,
                "filename": st.session_state.uploaded_filename,
                "review_date": date.today().strftime("%d %b %Y"),
            }
            st.session_state.review_started = True
    except Exception:
        st.session_state.review_started = False
        st.error("Unable to complete the review. Please check the selected Review Profile and PowerPoint file.")
        st.exception(traceback.format_exc())


def _render_summary_section():
    result = st.session_state.review_result or {}
    analysis = st.session_state.analysis_result or {}
    findings = unique_findings(analysis)

    result_status = str(result.get("status", "UNKNOWN"))

    # Keep the summary decision-oriented and visually aligned with the
    # reference design: one outcome hero, five metrics, then the human-review
    # notice. Detailed evidence remains in the Findings/Evidence views.
    if result_status == "PASS":
        result_class = "msa-summary-result-pass"
        result_title = "No Issues Detected"
        result_message = "No configured rule findings were detected in the material."
        notice_title = "Preliminary screening completed"
        notice_message = (
            "This prototype provides preliminary findings and evidence. "
            "Final compliance decision remains with the "
            "<strong>Nominated Signatory</strong>."
        )
    elif result_status == "ACTION REQUIRED":
        result_class = "msa-summary-result-action"
        result_title = "Action Required"
        result_message = "A potential compliance issue was detected and requires further action."
        notice_title = "Further action is required"
        notice_message = (
            "This prototype provides preliminary findings and evidence. "
            "Final compliance decision remains with the "
            "<strong>Nominated Signatory</strong>."
        )
    else:
        result_class = "msa-summary-result-review"
        result_title = "Review Required"
        result_message = "Potential issue(s) were identified and human review is recommended."
        notice_title = "Human review remains required"
        notice_message = (
            "This prototype provides preliminary findings and evidence. "
            "Final compliance decision remains with the "
            "<strong>Nominated Signatory</strong>."
        )

    # Calculate how many slides contain findings for the supporting metric copy.
    finding_slide_numbers = set()
    for finding in findings:
        if not isinstance(finding, dict):
            continue
        slide_value = finding.get("slide_number", finding.get("slide"))
        if slide_value not in (None, "", "?"):
            finding_slide_numbers.add(str(slide_value))

    findings_count = int(result.get("findings", 0) or 0)
    finding_slide_count = len(finding_slide_numbers)

    if findings_count == 0:
        findings_meta = "No findings"
    elif finding_slide_count == 1:
        findings_meta = "Across 1 slide"
    else:
        findings_meta = f"Across {finding_slide_count} slides"

    high_count = int(result.get("high", 0) or 0)
    medium_count = int(result.get("medium", 0) or 0)
    low_count = int(result.get("low", 0) or 0)

    high_meta = "No critical" if high_count == 0 else "Critical attention"
    medium_meta = "Attention" if medium_count > 0 else "None"
    low_meta = "None" if low_count == 0 else "Minor"

    # Severity values use semantic colors only when the count is non-zero.
    # A zero stays neutral/black so the summary remains visually calm.
    high_value_class = " msa-summary-metric-high" if high_count > 0 else ""
    medium_value_class = " msa-summary-metric-medium" if medium_count > 0 else ""
    low_value_class = " msa-summary-metric-low" if low_count > 0 else ""

    render_html(
        f'<div class="msa-summary-result-shell {result_class}">'
        f'  <div class="msa-summary-result-hero">'
        f'    <div class="msa-summary-result-badge">'
        f'      <span class="msa-summary-result-dot"></span>'
        f'      REVIEW COMPLETED'
        f'    </div>'
        f'    <div class="msa-summary-result-title">{esc(result_title)}</div>'
        f'    <div class="msa-summary-result-message">{esc(result_message)}</div>'
        f'    <div class="msa-summary-result-meta">'
        f'      <div>'
        f'        <div class="msa-summary-result-meta-label">PROFILE</div>'
        f'        <div class="msa-summary-result-meta-value">{esc(result.get("profile"))}</div>'
        f'      </div>'
        f'      <div>'
        f'        <div class="msa-summary-result-meta-label">MATERIAL</div>'
        f'        <div class="msa-summary-result-meta-value">{esc(result.get("filename"))}</div>'
        f'      </div>'
        f'    </div>'
        f'  </div>'

        f'  <div class="msa-summary-result-metrics">'
        f'    <div class="msa-summary-metric-card">'
        f'      <div class="msa-summary-metric-label">SLIDES</div>'
        f'      <div class="msa-summary-metric-value">{esc(result.get("slides", 0))}</div>'
        f'      <div class="msa-summary-metric-meta">Processed</div>'
        f'    </div>'
        f'    <div class="msa-summary-metric-card">'
        f'      <div class="msa-summary-metric-label">FINDINGS</div>'
        f'      <div class="msa-summary-metric-value">{esc(findings_count)}</div>'
        f'      <div class="msa-summary-metric-meta">{esc(findings_meta)}</div>'
        f'    </div>'
        f'    <div class="msa-summary-metric-card">'
        f'      <div class="msa-summary-metric-label">HIGH</div>'
        f'      <div class="msa-summary-metric-value{high_value_class}">{esc(high_count)}</div>'
        f'      <div class="msa-summary-metric-meta">{esc(high_meta)}</div>'
        f'    </div>'
        f'    <div class="msa-summary-metric-card">'
        f'      <div class="msa-summary-metric-label">MEDIUM</div>'
        f'      <div class="msa-summary-metric-value{medium_value_class}">{esc(medium_count)}</div>'
        f'      <div class="msa-summary-metric-meta">{esc(medium_meta)}</div>'
        f'    </div>'
        f'    <div class="msa-summary-metric-card">'
        f'      <div class="msa-summary-metric-label">LOW</div>'
        f'      <div class="msa-summary-metric-value{low_value_class}">{esc(low_count)}</div>'
        f'      <div class="msa-summary-metric-meta">{esc(low_meta)}</div>'
        f'    </div>'
        f'  </div>'

        f'  <div class="msa-summary-human-review">'
        f'    <div class="msa-summary-human-icon">!</div>'
        f'    <div>'
        f'      <div class="msa-summary-human-title">{esc(notice_title)}</div>'
        f'      <div class="msa-summary-human-message">{notice_message}</div>'
        f'    </div>'
        f'  </div>'
        f'</div>'
    )


def _render_findings_section():
    result = st.session_state.review_result or {}
    analysis = st.session_state.analysis_result or {}
    evidence = st.session_state.evidence_result
    ai = st.session_state.ai_result or {}
    findings = unique_findings(analysis)
    header("▥", "Material Analysis", "Rule-based preliminary screening and contextual analysis.")
    render_html(
        '<div class="msa-metric-grid">'
        f'<div class="msa-metric-card"><div class="msa-metric-label">Review Profile</div><div class="msa-metric-value">{esc(result.get("profile"))}</div></div>'
        f'<div class="msa-metric-card"><div class="msa-metric-label">Material</div><div class="msa-metric-value">PPTX</div></div>'
        f'<div class="msa-metric-card"><div class="msa-metric-label">Slides Analyzed</div><div class="msa-metric-value">{esc(result.get("slides"))}</div></div>'
        f'<div class="msa-metric-card"><div class="msa-metric-label">Active Rules Applied</div><div class="msa-metric-value">{esc(result.get("active_rules", 0))}</div></div>'
        '</div>'
    )

    render_html('<div id="msa-findings"></div>')
    header("⌕", "Findings", "Potential issues identified by the configured review rules.")
    if findings:
        grouped = {}
        for finding in findings:
            key = finding_slide_key(finding)
            grouped.setdefault(key, []).append(finding)

        def slide_sort_key(value):
            try:
                return (0, int(value))
            except (TypeError, ValueError):
                return (1, str(value))

        for slide_key in sorted(grouped, key=slide_sort_key):
            render_html(render_slide_findings_group(slide_key, grouped[slide_key], evidence, ai))
    else:
        empty_text = "ไม่พบรายการที่ตรงกับกฎตรวจสอบที่ตั้งค่าไว้" if st.session_state.get("result_language", "English") == "ไทย" else "No findings were generated by the configured rules."
        render_html(f'<div class="msa-empty-card">{esc(empty_text)}</div>')

    # Evidence is intentionally presented inside each Finding card.
    # The Evidence Layer remains available in session state/backend, but the
    # separate Evidence list is not rendered here to avoid duplicating the
    # same reviewer-facing information twice.



def _render_context_section():
    is_thai = st.session_state.get("result_language", "English") == "ไทย"
    ai = st.session_state.ai_result or {}
    header(
        "◌",
        "การวิเคราะห์บริบท" if is_thai else "Context Analysis",
        "ระบบระบุบริบททางธุรกิจของแต่ละสไลด์ เพื่อช่วยประกอบการเลือกแนวทางและกฎที่เกี่ยวข้อง"
        if is_thai else
        "The system identifies the business context of each slide to support the applicable review profile and rules.",
    )
    context_items = ai.get("items", []) if isinstance(ai, dict) else []
    if context_items:
        def context_sort_key(item):
            value = item.get("slide_number", item.get("slide", 0))
            try:
                return int(value)
            except (TypeError, ValueError):
                return 0

        category_translations = {
            "BUSINESS_MEETING": "การประชุมทางธุรกิจ",
            "PROMOTIONAL_CONTEXT": "บริบทเชิงส่งเสริมการขาย",
            "NON_PROMOTIONAL_CONTEXT": "บริบทที่ไม่ใช่การส่งเสริมการขาย",
            "NOT_CLASSIFIED": "ยังไม่จัดประเภท",
            "UNCERTAIN": "ไม่แน่ชัด",
            "UNKNOWN": "ไม่ทราบ",
        }
        reason_translations = {
            "The slide contains meeting or discussion language, but there is not enough evidence to establish promotional context.":
                "สไลด์มีถ้อยคำเกี่ยวกับการประชุมหรือการหารือ แต่ยังมีหลักฐานไม่เพียงพอที่จะระบุว่าเป็นบริบทส่งเสริมการขาย",
            "The slide contains promotional/product-related indicators that may require promotional review.":
                "สไลด์มีสัญญาณที่เกี่ยวข้องกับผลิตภัณฑ์หรือการส่งเสริมการขาย ซึ่งอาจต้องได้รับการตรวจสอบตามแนวทางสำหรับสื่อส่งเสริมการขาย",
            "No meaningful contextual indicators were identified in the available slide text.":
                "ไม่พบสัญญาณด้านบริบทที่มีนัยสำคัญจากข้อความสไลด์ที่มีอยู่",
            "The slide contains no meaningful contextual indicators.":
                "สไลด์ไม่มีสัญญาณด้านบริบทที่มีนัยสำคัญ",
            "The slide contains non-promotional business context.":
                "สไลด์มีบริบททางธุรกิจที่ไม่ใช่การส่งเสริมการขาย",
            "The slide may contain product-related information that requires human review.":
                "สไลด์อาจมีข้อมูลเกี่ยวกับผลิตภัณฑ์ที่ควรให้ผู้เชี่ยวชาญตรวจสอบ",
            "No contextual interpretation available.":
                "ไม่มีข้อมูลเพียงพอสำหรับอธิบายบริบทของสไลด์",
        }
        context_cards = []
        for item in sorted(context_items, key=context_sort_key):
            slide_number = item.get("slide_number", item.get("slide", "?"))
            category = item.get("classification", item.get("category", "CONTEXT"))
            confidence = item.get("confidence", 0)
            category_key = str(category).strip().upper()
            # Check NON_PROMOTIONAL before PROMOTIONAL because the former
            # contains the latter as a substring.
            if "NON_PROMOTIONAL" in category_key or "NON-PROMOTIONAL" in category_key or "NONPROMOTIONAL" in category_key:
                category_class = "msa-context-category-nonpromotional"
            elif "BUSINESS_MEETING" in category_key or "BUSINESS MEETING" in category_key:
                category_class = "msa-context-category-business"
            elif "PROMOTIONAL" in category_key:
                category_class = "msa-context-category-promotional"
            elif "UNCERTAIN" in category_key or "UNKNOWN" in category_key:
                category_class = "msa-context-category-uncertain"
            else:
                category_class = "msa-context-category-unclassified"
            reason = item.get("reason", item.get("context", "No contextual interpretation available."))
            keywords = item.get("matched_terms", item.get("keywords", []))
            if isinstance(keywords, str):
                keywords = [keywords]
            display_category = category_translations.get(category_key, category.replace("_", " ").title()) if is_thai else category
            reason_text = str(reason or "").strip()
            if is_thai:
                # Localize known classifier outputs and provide a safe Thai
                # fallback for new English reason strings instead of leaking
                # untranslated text into the Thai UI.
                display_reason = reason_translations.get(reason_text)
                if display_reason is None:
                    reason_lower = reason_text.casefold()
                    if "meeting" in reason_lower or "discussion" in reason_lower:
                        display_reason = "สไลด์มีบริบทเกี่ยวกับการประชุมหรือการหารือ แต่หลักฐานยังไม่เพียงพอที่จะสรุปว่าเป็นสื่อส่งเสริมการขาย"
                    elif "promotional" in reason_lower or "product-related" in reason_lower or "product related" in reason_lower:
                        display_reason = "สไลด์มีสัญญาณที่เกี่ยวข้องกับผลิตภัณฑ์หรือการส่งเสริมการขาย ซึ่งควรตรวจสอบเพิ่มเติม"
                    elif "no meaningful" in reason_lower or "no contextual" in reason_lower or "not identified" in reason_lower:
                        display_reason = "ไม่พบสัญญาณด้านบริบทที่มีนัยสำคัญจากข้อความสไลด์ที่มีอยู่"
                    elif not reason_text:
                        display_reason = "ไม่มีข้อมูลเพียงพอสำหรับอธิบายบริบทของสไลด์"
                    else:
                        display_reason = "ระบบจัดประเภทบริบทจากข้อความในสไลด์ ควรตรวจสอบข้อความต้นฉบับประกอบก่อนสรุปผล"
            else:
                display_reason = reason_text
            confidence_num = 0
            try:
                confidence_num = max(0, min(100, float(confidence)))
            except (TypeError, ValueError):
                pass
            confidence_label = "ความเชื่อมั่น" if is_thai else "Confidence"
            interpretation_label = "คำอธิบาย" if is_thai else "Interpretation"
            slide_label = "สไลด์" if is_thai else "Slide"
            context_cards.append(
                f'<div class="msa-context-card">'
                f'<div class="msa-context-header">'
                f'<strong>{slide_label} {esc(slide_number)}</strong>'
                f'<span class="msa-context-badge msa-context-category {category_class}">{esc(display_category)}</span>'
                f'</div>'
                f'<div class="msa-context-grid">'
                f'<div class="msa-context-confidence">'
                f'<div class="msa-small-label">{confidence_label}</div>'
                f'<div class="msa-confidence-line">'
                f'<div class="msa-confidence-track" role="progressbar" '
                f'aria-label="{slide_label} {esc(slide_number)} {confidence_label}" '
                f'aria-valuemin="0" aria-valuemax="100" aria-valuenow="{confidence_num:.0f}">'
                f'<span class="msa-confidence-fill '
                f'{"msa-confidence-high" if confidence_num >= 90 else "msa-confidence-medium" if confidence_num >= 85 else "msa-confidence-low"}" '
                f'style="width:{confidence_num}%"></span>'
                f'</div><strong class="msa-confidence-value">{esc(confidence)}%</strong>'
                f'</div></div>'
                f'<div><div class="msa-small-label">{interpretation_label}</div>{esc(display_reason)}</div>'
                f'</div><div class="msa-chip-wrap">{render_signal_chips(keywords)}</div></div>'
            )
        render_html('<div class="msa-context-grid-wrap">' + ''.join(context_cards) + '</div>')
    else:
        empty_text = "ยังไม่มีข้อมูลการวิเคราะห์บริบท" if is_thai else "No context analysis is available."
        render_html(f'<div class="msa-empty-card">{esc(empty_text)}</div>')


def _render_ai_section():
    """Reviewer-oriented cross-slide synthesis; findings remain the source of truth."""
    is_thai = st.session_state.get("result_language", "English") == "ไทย"
    ai = st.session_state.ai_result or {}
    analysis = st.session_state.analysis_result or {}
    findings = unique_findings(analysis)

    header(
        "✦",
        "สรุปประเด็นและแผนดำเนินการด้วย AI" if is_thai else "AI Review Priorities & Action Plan",
        "สรุปประเด็นจากหลายสไลด์เพื่อช่วยจัดลำดับการตรวจสอบ ไม่ใช่คำตัดสิน Compliance ขั้นสุดท้าย"
        if is_thai else
        "A cross-slide summary to help prioritize human review. This is not a final compliance decision.",
    )

    # Aggregate findings by rule so the tab adds cross-slide prioritization,
    # rather than repeating the per-finding explanation already available in Findings.
    rule_groups = {}
    slide_numbers = set()
    for finding in findings:
        rule_id = str(finding.get("rule_id", finding.get("rule_code", "")) or "")
        rule_name = str(finding.get("rule_name", finding.get("rule_title", "Configured Rule")) or "Configured Rule")
        key = rule_id or rule_name
        group = rule_groups.setdefault(key, {
            "rule_id": rule_id, "rule_name": rule_name, "findings": [],
            "slides": set(), "severity": "low", "terms": set()
        })
        group["findings"].append(finding)
        slide = finding.get("slide_number", finding.get("slide"))
        if slide is not None and str(slide).strip():
            group["slides"].add(str(slide))
            slide_numbers.add(str(slide))
        severity = str(finding.get("severity", finding.get("risk_level", "low"))).casefold()
        rank = {"critical": 4, "high": 3, "medium": 2, "low": 1}
        if rank.get(severity, 1) > rank.get(group["severity"], 1):
            group["severity"] = severity
        terms = finding.get("matched_keywords", finding.get("matched_terms", finding.get("evidence_keywords", [])))
        if isinstance(terms, str):
            terms = [terms]
        if isinstance(terms, (list, tuple, set)):
            group["terms"].update(str(t) for t in terms if str(t).strip())

    priority = {"critical": 0, "high": 1, "medium": 2, "low": 3}
    groups = sorted(rule_groups.values(), key=lambda g: (priority.get(g["severity"], 4), -len(g["findings"]), g["rule_name"].casefold()))
    high_count = sum(1 for f in findings if str(f.get("severity", f.get("risk_level", "low"))).casefold() in {"high", "critical"})
    context_items = ai.get("items", []) if isinstance(ai, dict) else []
    context_count = len(context_items) if isinstance(context_items, list) else 0

    metric_labels = ("รายการ Findings", "สไลด์ที่พบประเด็น", "กฎที่เกี่ยวข้อง") if is_thai else ("Findings", "Slides flagged", "Rules involved")
    metrics = [len(findings), len(slide_numbers), len(groups)]
    render_html('<div class="msa-ai-metric-grid">' + ''.join(
        f'<div class="msa-ai-metric"><div>{esc(label)}</div><strong>{esc(value)}</strong></div>'
        for label, value in zip(metric_labels, metrics)
    ) + '</div>')

    if not groups:
        render_html(f'<div class="msa-empty-card">{"ยังไม่มี Finding ให้จัดลำดับความสำคัญ" if is_thai else "No findings are available to prioritize."}</div>')
    else:
        render_html(f'<div class="msa-ai-section-title">{"ประเด็นที่ควรตรวจสอบก่อน" if is_thai else "Review priorities"}</div>')
        cards = []
        for index, group in enumerate(groups, 1):
            slides = sorted(group["slides"], key=lambda x: (0, int(x)) if x.isdigit() else (1, x))
            slide_text = ", ".join(slides)
            count = len(group["findings"])
            severity = group["severity"].title()
            severity_th = {"Critical": "วิกฤต", "High": "สูง", "Medium": "ปานกลาง", "Low": "ต่ำ"}.get(severity, severity)
            terms = sorted(group["terms"], key=str.casefold)[:10]
            chips = ''.join(f'<span class="msa-ai-term">{esc(term)}</span>' for term in terms)
            cards.append(
                f'<article class="msa-ai-priority-card">'
                f'<div class="msa-ai-priority-top"><span class="msa-ai-rank">{("ลำดับ " if is_thai else "Priority ")}{index}</span>'
                f'<span class="msa-ai-severity msa-ai-severity-{esc(group["severity"])}">{esc(severity_th if is_thai else severity)}</span></div>'
                f'<h3>{esc(group["rule_name"])}</h3>'
                f'<div class="msa-ai-priority-meta">{esc(group["rule_id"])} · '
                f'{("สไลด์ " if is_thai else "Slides ")}{esc(slide_text or "—")} · '
                f'{count} {("รายการ" if is_thai else ("finding" if count == 1 else "findings"))}</div>'
                f'<div class="msa-ai-term-list">{chips}</div>'
                f'</article>'
            )
        render_html('<div class="msa-ai-priority-grid">' + ''.join(cards) + '</div>')

        render_html(f'<div class="msa-ai-section-title">{"แผนดำเนินการที่แนะนำ" if is_thai else "Recommended action plan"}</div>')
        actions = []
        for index, group in enumerate(groups, 1):
            representative = group["findings"][0]
            explanation = build_ai_explanation(representative, ai) or {}
            if is_thai:
                _, _ = translate_ai_explanation(representative, explanation)
                rec, _suggestion = translate_recommendation_suggestion(
                    representative,
                    explanation.get("recommendation", ""),
                    explanation.get("suggestion", explanation.get("recommendation", "")),
                )
                fallback = "ตรวจสอบหลักฐานและแหล่งอ้างอิงที่เกี่ยวข้อง พร้อมยืนยันกับผู้รับผิดชอบและ NS ก่อนดำเนินการต่อ"
            else:
                rec = explanation.get("recommendation", explanation.get("suggestion", ""))
                fallback = "Review the supporting evidence and applicable references with the material owner and NS before proceeding."
            action_text = str(rec or fallback).strip()
            slides = sorted(group["slides"], key=lambda x: (0, int(x)) if x.isdigit() else (1, x))
            actions.append(
                f'<article class="msa-ai-action-row"><div class="msa-ai-action-number">{index}</div>'
                f'<div class="msa-ai-action-content"><h3>{esc(group["rule_name"])}</h3>'
                f'<p>{esc(action_text)}</p><div class="msa-ai-priority-meta">{esc(group["rule_id"])} · '
                f'{("สไลด์ " if is_thai else "Slides ")}{esc(", ".join(slides) or "—")}</div></div></article>'
            )
        render_html('<div class="msa-ai-action-list">' + ''.join(actions) + '</div>')

    # Context to confirm: compact, color-coded rows for quick human review.
    if context_items:
        context_title = "ข้อมูลที่ควรยืนยันเพิ่มเติม" if is_thai else "Context to confirm"
        context_subtitle = (
            "ใช้ข้อมูลนี้ประกอบการพิจารณาและตรวจสอบกับเนื้อหาต้นฉบับ"
            if is_thai else
            "Use these signals to guide review against the original slide content."
        )
        render_html(
            '<div class="msa-ai-context-heading">'
            '<div class="msa-ai-context-heading-icon">▣</div>'
            '<div><div class="msa-ai-section-title">' + esc(context_title) + '</div>'
            '<div class="msa-ai-context-subtitle">' + esc(context_subtitle) + '</div></div></div>'
        )
        context_rows = []
        for item in sorted(
            context_items,
            key=lambda x: int(x.get("slide_number", x.get("slide", 0)) or 0)
            if str(x.get("slide_number", x.get("slide", 0)) or 0).isdigit() else 0
        ):
            slide = item.get("slide_number", item.get("slide", "?"))
            raw_category = str(item.get("classification", item.get("category", "NOT_CLASSIFIED"))).upper()
            category_map = {
                "PROMOTIONAL_CONTEXT": ("บริบทเชิงส่งเสริมการขาย", "ส่งเสริมการขาย") if is_thai else ("Promotional context", "Promotional"),
                "NON_PROMOTIONAL_CONTEXT": ("บริบทที่ไม่ใช่การส่งเสริมการขาย", "ไม่ใช่การส่งเสริมการขาย") if is_thai else ("Non-promotional context", "Non-promotional"),
                "BUSINESS_MEETING": ("บริบทการประชุมทางธุรกิจ", "การประชุม") if is_thai else ("Business meeting context", "Business meeting"),
                "NOT_CLASSIFIED": ("ยังไม่จัดประเภท", "ยังไม่จัดประเภท") if is_thai else ("Not classified", "Not classified"),
                "UNCERTAIN": ("บริบทไม่ชัดเจน", "ควรตรวจสอบ") if is_thai else ("Uncertain context", "Review needed"),
            }
            category_title, category_tag = category_map.get(
                raw_category,
                (("ควรตรวจสอบบริบทเพิ่มเติม", "ควรตรวจสอบ") if is_thai else ("Context requires review", "Review needed"))
            )
            reason = str(item.get("reason", ""))
            low = reason.casefold()
            if is_thai:
                if "meeting" in low or "discussion" in low:
                    reason = "พบถ้อยคำเกี่ยวกับการประชุมหรือการหารือ ควรตรวจสอบบริบทจากเนื้อหาต้นฉบับ"
                elif "promotional" in low or "product-related" in low:
                    reason = "พบสัญญาณที่เกี่ยวข้องกับผลิตภัณฑ์หรือการส่งเสริมการขาย ควรตรวจสอบเพิ่มเติม"
                else:
                    reason = "ใช้เป็นข้อมูลประกอบการพิจารณาและตรวจสอบกับเนื้อหาต้นฉบับ"
            elif not reason.strip():
                reason = "Confirm the context against the original slide content."
            row_class = "msa-ai-context-row msa-ai-context-promotional" if raw_category == "PROMOTIONAL_CONTEXT" else "msa-ai-context-row msa-ai-context-neutral"
            slide_label = "สไลด์ " if is_thai else "Slide "
            context_rows.append(
                f'<div class="{row_class}"><div class="msa-ai-context-row-main">'
                f'<div class="msa-ai-context-row-copy"><strong>{esc(slide_label + str(slide))} · {esc(category_title)}</strong>'
                f'<span>{esc(reason)}</span></div>'
                f'<div class="msa-ai-context-tag">{esc(category_tag)}</div></div></div>'
            )
        render_html('<div class="msa-ai-context-list">' + ''.join(context_rows) + '</div>')

    if ai.get("error"):
        render_html(f'<div class="msa-technical-note">{("รายละเอียดทางเทคนิค" if is_thai else "Technical detail")}: {esc(ai.get("error"))}</div>')
    render_html('<div class="msa-technical-note">' + (
        'ผลลัพธ์นี้ช่วยจัดลำดับการตรวจสอบเท่านั้น ต้องตรวจสอบหลักฐานและข้อมูลการอนุมัติที่เกี่ยวข้อง โดยการตัดสินใจ Compliance ขั้นสุดท้ายเป็นหน้าที่ของ NS'
        if is_thai else
        'This output only helps prioritize review. Supporting evidence and approval records must be checked; the final Compliance decision remains with the Nominated Signatory (NS).'
    ) + '</div>')


def _render_report_section():
    result = st.session_state.review_result or {}
    result_status = str(result.get("status", "UNKNOWN"))
    # --------------------------------------------------------
    # Pre-Review Report
    # --------------------------------------------------------
    render_html('<div id="msa-pre-review-report"></div>')
    header("□", "Pre-Review Report", "Generate a print-ready preliminary cover page for the material review.")
    report_status = result_status
    report_severity = "high" if report_status == "ACTION REQUIRED" else ("medium" if report_status == "REVIEW REQUIRED" else "low")
    badge = report_status
    message = {
        "PASS": "No findings were detected. The cover page will indicate PASS.",
        "REVIEW REQUIRED": "Potential findings were detected. The cover page will indicate REVIEW REQUIRED.",
        "ACTION REQUIRED": "A potential high-severity issue was detected. The cover page will indicate ACTION REQUIRED.",
    }.get(report_status, "Preliminary review completed.")

    render_html(
        f'<div class="msa-report-status"><span class="msa-status-pill {status_class(report_severity)}">{esc(badge)}</span>'
        f'<span>{esc(message)}</span><div class="msa-report-note">This is a preliminary pre-review document. It does not represent final approval or rejection.</div></div>'
    )

    if st.button("📄 Generate Pre-Review Cover Page", type="primary", use_container_width=True, key="generate_cover"):
        st.session_state.cover_generated = True

    if st.session_state.get("cover_generated", False):
        cover_html = build_cover_html(result)
        finding_count = int(result.get("findings", 0) or 0)
        high_count = int(result.get("high", 0) or 0)
        medium_count = int(result.get("medium", 0) or 0)
        low_count = int(result.get("low", 0) or 0)
        active_rules = int(result.get("active_rules", 0) or 0)

        if report_status == "PASS":
            cover_message = "No configured rule findings were detected in the material."
            cover_next_step = "Material may proceed to the next step, subject to the Nominated Signatory's final decision."
        elif report_status == "ACTION REQUIRED":
            cover_message = "A potential high-severity issue was detected and requires further review."
            cover_next_step = "Review the high-severity finding, evidence, and applicable reference before proceeding."
        else:
            cover_message = "Potential issue(s) were identified and human review is recommended."
            cover_next_step = "Review the flagged findings, supporting evidence, and applicable reference before proceeding."

        render_html(
            f'<div class="msa-cover-preview">'
            f'<div class="msa-cover-preview-header">'
            f'<div><div class="msa-cover-kicker">ASTRAZENECA · MATERIAL SCANNER AGENT</div>'
            f'<div class="msa-cover-title">Preliminary Material Pre-Review</div>'
            f'<div class="msa-cover-subtitle">AI-assisted screening summary for human review</div></div>'
            f'<div class="msa-cover-preview-badge {status_class(report_severity)}">{esc(badge)}</div>'
            f'</div>'
            f'<div class="msa-cover-status msa-cover-status-{report_severity}">'
            f'<strong>{esc(cover_message)}</strong><span>{esc(cover_next_step)}</span></div>'
            f'<div class="msa-cover-meta-grid">'
            f'<div><div class="msa-small-label">Material</div><strong>{esc(result.get("filename"))}</strong></div>'
            f'<div><div class="msa-small-label">Review Profile</div><strong>{esc(result.get("profile"))}</strong></div>'
            f'<div><div class="msa-small-label">Review Date</div><strong>{esc(result.get("review_date"))}</strong></div>'
            f'<div><div class="msa-small-label">Slides Reviewed</div><strong>{esc(result.get("slides"))}</strong></div>'
            f'</div>'
            f'<div class="msa-cover-summary-label">SCREENING SUMMARY</div>'
            f'<div class="msa-cover-summary-grid">'
            f'<div><strong>{finding_count}</strong><span>Findings</span></div>'
            f'<div class="msa-cover-count-high"><strong>{high_count}</strong><span>High</span></div>'
            f'<div class="msa-cover-count-medium"><strong>{medium_count}</strong><span>Medium</span></div>'
            f'<div class="msa-cover-count-low"><strong>{low_count}</strong><span>Low</span></div>'
            f'<div><strong>{active_rules}</strong><span>Active Rules</span></div>'
            f'</div>'
            f'<div class="msa-cover-next-step"><div class="msa-small-label">NEXT STEP</div>{esc(cover_next_step)}</div>'
            f'<div class="msa-human-review-alert">'
            f'<div class="msa-human-review-icon">!</div>'
            f'<div class="msa-human-review-copy">'
            f'<div class="msa-human-review-title">Human review remains required</div>'
            f'<div class="msa-human-review-text">This prototype provides preliminary findings and evidence. Final compliance decision remains with the <strong>Nominated Signatory</strong>.</div>'
            f'</div>'
            f'</div>'
            f'</div>'
        )
        st.download_button(
            "⬇️ Download Pre-Review Cover Page",
            data=cover_html,
            file_name="Material_Scanner_Agent_Pre_Review.html",
            mime="text/html",
            use_container_width=True,
            key="download_cover",
        )

    # --------------------------------------------------------
    # New review controls
    # --------------------------------------------------------
    st.divider()
    header("↻", "Review Actions", "Start a new review with a new material.")

    # Start immediately: clear the current result and return to the
    # New Review state without showing a second confirmation panel.
    if st.button("＋ Start New Review", type="primary", use_container_width=True, key="new_review"):
        reset_review_state(clear_material=True)
        st.session_state.show_new_review_confirm = False
        st.session_state.cover_generated = False
        st.session_state.result_section = "Summary"
        st.rerun()

def render_results():
    """Render only the selected result view so users do not need to scroll through the full analysis."""
    section = st.session_state.get("result_section", "Summary")
    if section == "Summary":
        _render_summary_section()
    elif section == "Findings":
        _render_findings_section()
    elif section == "Context Analysis":
        _render_context_section()
    elif section == "AI Analysis":
        _render_ai_section()
    elif section == "Report":
        _render_report_section()
    else:
        _render_summary_section()


def build_cover_html(result):
    """Build a standalone, print-friendly Pre-Review Cover Page."""
    status = str(result.get("status", "UNKNOWN"))
    is_action = status == "ACTION REQUIRED"
    is_review = status == "REVIEW REQUIRED"
    is_pass = status == "PASS"
    badge = status if status in {"PASS", "REVIEW REQUIRED", "ACTION REQUIRED"} else "UNKNOWN"

    if is_action:
        accent = "#b42318"
        accent_bg = "#fff0f0"
        status_text = "A potential high-severity issue was detected and requires further review."
        next_step = "Review the high-severity finding, supporting evidence, and applicable reference before proceeding."
    elif is_review:
        accent = "#a76500"
        accent_bg = "#fff7df"
        status_text = "Potential issue(s) were identified and human review is recommended."
        next_step = "Review the flagged findings, supporting evidence, and applicable reference before proceeding."
    elif is_pass:
        accent = "#237a45"
        accent_bg = "#eaf8ef"
        status_text = "No configured rule findings were detected in the material."
        next_step = "Material may proceed to the next step, subject to the Nominated Signatory's final decision."
    else:
        accent = "#64748b"
        accent_bg = "#f1f5f9"
        status_text = "Preliminary screening completed."
        next_step = "Review the screening summary before proceeding."

    filename = esc(result.get("filename", "-"))
    profile = esc(result.get("profile", "-"))
    review_date = esc(result.get("review_date", date.today().strftime("%d %b %Y")))
    slides = esc(result.get("slides", 0))
    findings = esc(result.get("findings", 0))
    high = esc(result.get("high", 0))
    medium = esc(result.get("medium", 0))
    low = esc(result.get("low", 0))
    active_rules = esc(result.get("active_rules", 0))

    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Material Scanner Agent - Pre-Review Cover Page</title>
<style>
:root{{--navy:#003865;--mulberry:#830051;--text:#24313d;--muted:#6b7785;--border:#dfe5ea;--soft:#f7f8fa;--accent:{accent};--accent-bg:{accent_bg};}}
*{{box-sizing:border-box;}}
body{{margin:0;background:#f4f5f8;color:var(--text);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;line-height:1.5;}}
.page{{width:min(920px,calc(100% - 40px));margin:40px auto;background:#fff;border:1px solid var(--border);border-radius:18px;box-shadow:0 14px 40px rgba(30,35,50,.08);overflow:hidden;}}
.header{{padding:30px 34px 26px;background:linear-gradient(135deg,#4d0030 0%,#830051 72%,#a4005f 100%);color:#fff;}}
.eyebrow{{font-size:10px;font-weight:800;letter-spacing:.12em;text-transform:uppercase;opacity:.78;}}
h1{{margin:10px 0 4px;font-size:28px;line-height:1.15;font-weight:750;letter-spacing:-.025em;}}
.subtitle{{font-size:13px;color:rgba(255,255,255,.82);}}
.status{{margin:26px 34px 0;padding:18px 20px;border:1px solid var(--accent);border-left:5px solid var(--accent);border-radius:12px;background:var(--accent-bg);}}
.status-badge{{display:inline-flex;padding:6px 10px;border-radius:999px;background:#fff;border:1px solid var(--accent);color:var(--accent);font-size:11px;font-weight:800;letter-spacing:.04em;}}
.status-text{{margin-top:9px;color:#28343e;font-size:14px;font-weight:650;}}
.content{{padding:0 34px 34px;}}
.section-label{{margin:24px 0 10px;color:var(--muted);font-size:10px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;}}
.grid{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px;}}
.card{{border:1px solid var(--border);border-radius:11px;padding:13px 14px;background:#fff;min-width:0;}}
.label{{color:var(--muted);font-size:10px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;}}
.value{{margin-top:5px;color:var(--navy);font-size:14px;font-weight:650;overflow-wrap:anywhere;}}
.metrics{{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:9px;}}
.metric{{border:1px solid var(--border);border-radius:11px;padding:12px 8px;text-align:center;background:var(--soft);}}
.metric strong{{display:block;color:var(--navy);font-size:20px;line-height:1.1;}}
.metric span{{display:block;margin-top:4px;color:var(--muted);font-size:9px;font-weight:800;letter-spacing:.05em;text-transform:uppercase;}}
.metric.high{{background:#fff5f5;border-color:#efc1c1;}} .metric.high strong{{color:#b42318;}}
.metric.medium{{background:#fff9eb;border-color:#f0d49a;}} .metric.medium strong{{color:#a76500;}}
.metric.low{{background:#f0faf3;border-color:#c4e3cd;}} .metric.low strong{{color:#237a45;}}
.next{{margin-top:18px;padding:15px 16px;border-radius:11px;background:#f7f8fa;border:1px solid var(--border);}}
.next .label{{margin-bottom:5px;}}
.human-review-alert{{display:flex;align-items:flex-start;gap:12px;margin-top:18px;padding:14px 16px;background:#fff7df;border:1px solid #efc34f;border-left:4px solid #f2ad00;border-radius:11px;box-sizing:border-box;}}
.human-review-icon{{flex:0 0 auto;width:28px;height:28px;display:flex;align-items:center;justify-content:center;border-radius:50%;background:#f2ad00;color:#fff;font-size:16px;font-weight:800;line-height:1;}}
.human-review-copy{{min-width:0;}}
.human-review-title{{color:#755400;font-size:12px;font-weight:800;line-height:1.35;margin-bottom:4px;}}
.human-review-text{{color:#5f4b16;font-size:11px;line-height:1.55;}}
.human-review-text strong{{color:#4d3b0f;}}
.disclaimer{{margin-top:18px;padding-top:16px;border-top:1px solid var(--border);color:var(--muted);font-size:11px;line-height:1.55;}}
.footer{{padding:15px 34px;border-top:1px solid var(--border);color:#7b8793;font-size:10px;background:#fbfbfc;}}
@media print{{body{{background:#fff;}}.page{{width:100%;margin:0;border:0;box-shadow:none;border-radius:0;}}}}
@media(max-width:680px){{.human-review-alert{{gap:10px;padding:12px 13px;}}.human-review-icon{{width:26px;height:26px;font-size:15px;}}}}
@media(max-width:680px){{.page{{width:calc(100% - 20px);margin:10px auto;}}.header,.content{{padding-left:20px;padding-right:20px;}}.status{{margin-left:20px;margin-right:20px;}}.grid,.metrics{{grid-template-columns:1fr 1fr;}}.metric:last-child{{grid-column:1 / -1;}}}}
@media(max-width:430px){{.grid,.metrics{{grid-template-columns:1fr;}}.metric:last-child{{grid-column:auto;}}h1{{font-size:24px;}}}}
</style></head>
<body><main class="page">
<header class="header"><div class="eyebrow">ASTRAZENECA · MATERIAL SCANNER AGENT</div><h1>Preliminary Material Pre-Review</h1><div class="subtitle">AI-assisted screening summary for human review</div></header>
<section class="status"><span class="status-badge">{badge}</span><div class="status-text">{status_text}</div></section>
<section class="content">
<div class="section-label">Review Details</div>
<div class="grid">
<div class="card"><div class="label">Material</div><div class="value">{filename}</div></div>
<div class="card"><div class="label">Review Profile</div><div class="value">{profile}</div></div>
<div class="card"><div class="label">Review Date</div><div class="value">{review_date}</div></div>
<div class="card"><div class="label">Slides Reviewed</div><div class="value">{slides}</div></div>
</div>
<div class="section-label">Screening Summary</div>
<div class="metrics">
<div class="metric"><strong>{findings}</strong><span>Findings</span></div>
<div class="metric high"><strong>{high}</strong><span>High</span></div>
<div class="metric medium"><strong>{medium}</strong><span>Medium</span></div>
<div class="metric low"><strong>{low}</strong><span>Low</span></div>
<div class="metric"><strong>{active_rules}</strong><span>Active Rules</span></div>
</div>
<div class="next"><div class="label">Next Step</div>{next_step}</div>
<div class="human-review-alert">
<div class="human-review-icon">!</div>
<div class="human-review-copy">
<div class="human-review-title">Human review remains required</div>
<div class="human-review-text">This prototype provides preliminary findings and evidence. Final compliance decision remains with the <strong>Nominated Signatory</strong>.</div>
</div>
</div>
</section>
<footer class="footer">Generated by Material Scanner Agent · AI-powered Material Pre-Review Prototype · AstraZeneca Hackathon 2026</footer>
</main></body></html>"""


# ------------------------------------------------------------
# Navigation
# ------------------------------------------------------------
render_app_chrome()
render_hero()
nav = render_application_nav()

if nav != st.session_state.page:
    st.session_state.page = nav
    st.rerun()

if st.session_state.page == "Rule Management":
    rule_management_page()
else:
    review_page()
