
import streamlit as st
import os
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
    )
except Exception:
    load_rules = get_rule = add_rule = update_rule = delete_rule = None
    enable_rule = disable_rule = None
    count_rules = count_active_rules = None


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
css_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "styles.css")
if os.path.exists(css_path):
    with open(css_path, "r", encoding="utf-8") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)


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
        category = str(rule.get("category", "")).strip().upper()
        status = str(rule.get("status", "")).strip().lower()
        if category == str(canonical).upper() and status == "active":
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
        if str(item_slide) == str(slide) and (not rule_id or str(item_rule) == str(rule_id)):
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
    item = ai_item_for_finding(ai_result, finding)
    slide = finding.get("slide_number", finding.get("slide", "?"))
    rule_name = finding.get("rule_name", finding.get("category", "configured rule"))
    signals = detected_values(finding)

    if item:
        context = item.get("context", "Relevant material context was identified.")
        reason = item.get("reason", "")
        recommendation = item.get(
            "recommendation",
            finding.get("recommendation", "Review the detected content against the applicable reference."),
        )
        risk = item.get("risk_level", finding.get("severity", "Review"))
        explanation = reason or (
            f"Slide {slide} was identified in a context relevant to {rule_name}. "
            "The finding should be reviewed against the applicable reference."
        )
        return {
            "explanation": explanation,
            "context": context,
            "recommendation": recommendation,
            "risk": risk,
        }

    signal_text = ", ".join(signals) if signals else "the detected content"
    return {
        "explanation": (
            f"Slide {slide} triggered {rule_name}. "
            f"The finding is supported by {signal_text} and should be reviewed "
            "against the applicable reference."
        ),
        "context": "No slide-level AI context item was available for this finding.",
        "recommendation": finding.get(
            "recommendation",
            "Review the detected content and confirm that it is supported by the applicable reference.",
        ),
        "risk": finding.get("severity", "Review"),
    }


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
        st.session_state.page = "Review"
        st.rerun()

    if rules_clicked and current != "Rule Management":
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


def render_finding_card(finding, evidence, ai_result):
    """Render one decision-ready finding in the Phase 2 review sequence.

    Reviewer flow is intentionally linear:
        Finding -> Evidence -> Why flagged -> Recommendation

    The Evidence Layer is the preferred source for evidence, rationale, and
    recommendation. Rule-engine values remain the fallback so older outputs
    continue to render safely. AI remains supporting interpretation only.
    """
    slide = finding.get("slide_number", finding.get("slide", "?"))
    rule_id = finding.get("rule_id", "")
    rule_name = finding.get("rule_name", finding.get("category", "Finding"))
    severity = finding.get("severity", "Review")
    action = finding.get("action", "FLAG_FOR_REVIEW")
    signals = detected_values(finding)

    # Evidence Layer is the source of truth for the reviewer-facing
    # explanation when it is available.
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
            finding.get(
                "reason",
                "The configured review rule identified content that should be reviewed.",
            ),
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
    evidence_source = evidence_item.get("source", "Evidence Layer / Rule Engine")

    ai = build_ai_explanation(finding, ai_result)
    ai_available = bool(ai.get("explanation") or ai.get("context") or ai.get("recommendation"))

    rule_note = (
        f'<div class="msa-inline-rule-note">{esc(rule_description)}</div>'
        if rule_description else ""
    )

    ai_panel = ""
    if ai_available:
        ai_panel = (
            f'<details class="msa-ai-disclosure">'
            f'<summary>'
            f'<span class="msa-ai-disclosure-title">View AI Explanation</span>'
            f'<span class="msa-ai-disclosure-meta">AI ASSISTED · Contextual support only</span>'
            f'</summary>'
            f'<div class="msa-ai-explanation-card">'
            f'<div class="msa-ai-explanation-top"><div class="msa-ai-field-label">AI-ASSISTED EXPLANATION</div>'
            f'<span class="msa-status-pill msa-status-info">AI ASSISTED</span></div>'
            f'<div class="msa-ai-explanation-title">Additional context for human review</div>'
            f'<div class="msa-ai-explanation-text">{esc(ai["explanation"])}</div>'
            f'<div class="msa-ai-explanation-context"><div class="msa-ai-field-label">Context Signal</div>{esc(ai["context"])}</div>'
            f'<div class="msa-ai-explanation-context"><div class="msa-ai-field-label">Suggested Review Action</div>{esc(ai["recommendation"])}</div>'
            f'<div class="msa-ai-disclaimer">AI provides contextual support only. Final compliance decision remains with the Nominated Signatory.</div>'
            f'</div>'
            f'</details>'
        )

    # Phase 2 WOW sequence:
    # 1) Finding — what was detected
    # 2) Evidence — what in the material supports the finding
    # 3) Why flagged — why the configured rule matters
    # 4) Recommendation — what the reviewer should do next
    render_html(
        f'<div class="msa-finding-card {status_class(severity)}">'
        f'<div class="msa-finding-top"><div>'
        f'<div class="msa-finding-title">{esc(rule_name)}</div>'
        f'<div class="msa-finding-meta">Slide {esc(slide)} · Rule {esc(rule_id)} · {esc(action)}</div>'
        f'</div><span class="msa-status-pill {status_class(severity)}">{esc(severity)}</span></div>'

        f'<div class="msa-finding-field"><div class="msa-finding-label">Finding</div>'
        f'<div class="msa-finding-text">Content on Slide {esc(slide)} matched the configured review rule <strong>{esc(rule_id or rule_name)}</strong>.</div></div>'

        f'<div class="msa-evidence-section"><div class="msa-evidence-label">Evidence</div>'
        f'<div class="msa-evidence-quote">{render_signal_chips(evidence_values)}'
        f'<div class="msa-evidence-source">Evidence source: {esc(evidence_source)} · Slide {esc(slide)}</div>'
        f'</div></div>'

        f'{rule_note}'

        f'<div class="msa-finding-field"><div class="msa-finding-label">Why was this flagged?</div>'
        f'<div class="msa-finding-text">{esc(why)}</div></div>'

        f'<div class="msa-finding-field msa-finding-action"><div class="msa-finding-label">Recommendation</div>'
        f'<div class="msa-finding-action-text"><strong>{esc(recommendation)}</strong></div></div>'

        f'{ai_panel}'
        f'</div>'
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
    source = item.get("source", "Rule Engine")

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
        f'<div class="msa-evidence-source">Evidence source: {esc(source)}</div>'
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
                        if st.button("Enable Rule", key=f"rm_enable_{rid}", use_container_width=True):
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
                    "status": "Active",
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
        for finding in findings:
            render_finding_card(finding, evidence, ai)
    else:
        render_html('<div class="msa-empty-card">No findings were generated by the configured rules.</div>')

    # Evidence is intentionally presented inside each Finding card.
    # The Evidence Layer remains available in session state/backend, but the
    # separate Evidence list is not rendered here to avoid duplicating the
    # same reviewer-facing information twice.



def _render_context_section():
    ai = st.session_state.ai_result or {}
    header("◌", "Context Analysis", "The system identifies the business context of each slide to support the applicable review profile and rules.")
    context_items = ai.get("items", []) if isinstance(ai, dict) else []
    if context_items:
        def context_sort_key(item):
            value = item.get("slide_number", item.get("slide", 0))
            try:
                return int(value)
            except (TypeError, ValueError):
                return 0

        context_cards = []
        for item in sorted(context_items, key=context_sort_key):
            slide_number = item.get("slide_number", item.get("slide", "?"))
            category = item.get(
                "classification",
                item.get("category", "CONTEXT"),
            )
            confidence = item.get("confidence", 0)
            reason = item.get(
                "reason",
                item.get("context", "No contextual interpretation available."),
            )
            keywords = item.get(
                "matched_terms",
                item.get("keywords", []),
            )
            if isinstance(keywords, str):
                keywords = [keywords]

            context_cards.append(
                f'<div class="msa-context-card">'
                f'<div class="msa-context-header">'
                f'<strong>Slide {esc(slide_number)}</strong>'
                f'<span class="msa-status-pill msa-status-info">{esc(category)}</span>'
                f'</div>'
                f'<div class="msa-context-grid">'
                f'<div><div class="msa-small-label">Confidence</div>'
                f'<strong>{esc(confidence)}%</strong></div>'
                f'<div><div class="msa-small-label">Interpretation</div>'
                f'{esc(reason)}</div>'
                f'</div>'
                f'<div class="msa-chip-wrap">{render_signal_chips(keywords)}</div>'
                f'</div>'
            )
        render_html('<div class="msa-context-grid-wrap">' + ''.join(context_cards) + '</div>')
    else:
        render_html('<div class="msa-empty-card">No context analysis is available.</div>')



def _render_ai_section():
    ai = st.session_state.ai_result or {}
    analysis = st.session_state.analysis_result or {}
    findings = unique_findings(analysis)
    header("✦", "AI Context Analysis", "AI-supported contextual interpretation. AI does not make the final compliance decision.")

    ai_status = str(ai.get("status", "COMPLETED"))
    ai_summary = ai.get("overall_summary", "Context analysis completed.")

    # Surface one concrete AI explanation as the main intelligence moment.
    # The detailed per-finding explanation remains available inside Findings.
    spotlight_finding = findings[0] if findings else None
    spotlight = build_ai_explanation(spotlight_finding, ai) if spotlight_finding else None

    spotlight_html = ""
    if spotlight_finding and spotlight:
        spotlight_html = (
            f'<div class="msa-ai-spotlight">'
            f'<div class="msa-ai-spotlight-kicker">WHY THIS NEEDS HUMAN REVIEW</div>'
            f'<div class="msa-ai-spotlight-title">Slide {esc(spotlight_finding.get("slide_number", spotlight_finding.get("slide", "?")))} · {esc(spotlight_finding.get("rule_name", "Configured Rule"))}</div>'
            f'<div class="msa-ai-spotlight-text">{esc(spotlight.get("explanation", "Contextual explanation is available in the finding."))}</div>'
            f'<div class="msa-ai-spotlight-grid">'
            f'<div><div class="msa-ai-field-label">CONTEXT SIGNAL</div>{esc(spotlight.get("context", "Relevant material context was identified."))}</div>'
            f'<div><div class="msa-ai-field-label">RECOMMENDED NEXT STEP</div>{esc(spotlight.get("recommendation", "Review the detected content against the applicable reference."))}</div>'
            f'</div>'
            f'</div>'
        )

    render_html(
        f'<div class="msa-ai-card">'
        f'<div class="msa-ai-top">'
        f'<div><div class="msa-ai-field-label">ANALYSIS STATUS</div>'
        f'<span class="msa-status-pill {status_class(ai_status)}">{esc(ai_status.title())}</span></div>'
        f'<span class="msa-status-pill msa-status-info">AI ASSISTED</span>'
        f'</div>'
        f'<div class="msa-ai-field">'
        f'<div class="msa-ai-field-label">OVERALL INTERPRETATION</div>'
        f'{esc(ai_summary)}'
        f'</div>'
        f'{spotlight_html}'
        f'</div>'
    )

    if ai.get("error"):
        render_html(
            f'<div class="msa-technical-note">'
            f'Technical detail: {esc(ai.get("error"))}'
            f'</div>'
        )

    render_html(
        '<div class="msa-technical-note">'
        'Finding explanation: Each finding uses the available slide-level context '
        'to explain why human review may be needed. This is supporting analysis, '
        'not a final compliance decision.'
        '</div>'
    )

    header("▤", "Extracted Content", "PowerPoint text extracted slide-by-slide.")
    slides_for_extraction = st.session_state.slides or []
    with st.expander(f"View extracted content · {len(slides_for_extraction)} slides", expanded=False):
        for slide in slides_for_extraction:
            texts = slide.get("texts", [])
            safe = [esc(x) for x in texts if str(x).strip()]
            bullet_items = "".join(f'<li class="msa-extracted-item">{x}</li>' for x in safe)
            render_html(
                f'<div class="msa-slide-card msa-extracted-card"><div class="msa-slide-topline">'
                f'<span class="msa-slide-pill">Slide {esc(slide.get("slide"))}</span><span class="msa-slide-count">Extracted content</span></div>'
                f'<div class="msa-extracted-bubble"><div class="msa-small-label">Slide Content</div><ul>{bullet_items}</ul></div></div>'
            )



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
    # Leaving Rule Management automatically locks the restricted area.
    # This ensures that returning to Rule Management requires the
    # administrator password again.
    if (
        st.session_state.page == "Rule Management"
        and nav == "Review"
    ):
        st.session_state.rule_admin_authenticated = False
        st.session_state.rule_admin_login_error = ""
        st.session_state.rule_admin_login_version = (
            st.session_state.get("rule_admin_login_version", 0) + 1
        )
        st.session_state.rule_management_flash = None

    st.session_state.page = nav
    st.rerun()

if st.session_state.page == "Rule Management":
    rule_management_page()
else:
    review_page()
