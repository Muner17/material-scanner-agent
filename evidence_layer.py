from typing import List, Dict, Any
import re


# ============================================================
# Evidence Layer
# Material Scanner Agent
#
# Purpose:
# Convert rule findings into human-readable evidence.
#
# This layer does NOT make a final compliance decision.
# It only identifies:
#   - What was detected
#   - Where it was detected
#   - Which rule triggered
#   - Why it may require review
#   - What the reviewer should verify
# ============================================================


def normalize_text(text: str) -> str:
    """
    Normalize text for matching while preserving the original
    text for evidence display.
    """
    if not text:
        return ""

    return " ".join(text.lower().split())


def clean_evidence_text(text: str) -> str:
    """
    Clean extracted text for display in the UI.
    """
    if not text:
        return ""

    text = " ".join(text.split())

    return text.strip()


# ------------------------------------------------------------
# Rule-specific semantic evidence patterns
#
# These patterns supplement keyword-based rules.
#
# Example:
# "more effective than"
#
# This may represent comparative language even if the rule
# keyword list only contains "better than".
# ------------------------------------------------------------

SEMANTIC_PATTERNS = {
    "comparative_statement": [
        r"\bmore\s+\w+\s+than\b",
        r"\bless\s+\w+\s+than\b",
        r"\bbetter\s+than\b",
        r"\bworse\s+than\b",
        r"\bsuperior\s+to\b",
        r"\binferior\s+to\b",
        r"\bcompared\s+with\b",
        r"\bcompared\s+to\b",
        r"\bcompared\s+against\b",
    ],
    "product_claim": [
        r"\bmore\s+effective\b",
        r"\beffective\b",
        r"\befficacy\b",
        r"\bsafety\b",
        r"\bsafe\b",
    ],
}


def find_pattern_evidence(text: str, pattern_list: List[str]) -> List[str]:
    """
    Find semantic phrases in slide text.
    Returns the actual matching phrases from the original text.
    """
    if not text:
        return []

    matches = []

    for pattern in pattern_list:
        for match in re.finditer(
            pattern,
            text,
            flags=re.IGNORECASE
        ):
            value = match.group(0).strip()

            if value and value not in matches:
                matches.append(value)

    return matches


def find_keyword_evidence(
    text: str,
    keywords: List[str]
) -> List[str]:
    """
    Find exact configured rule keywords in the original text.
    """
    if not text:
        return []

    normalized = normalize_text(text)
    matches = []

    for keyword in keywords or []:
        keyword_clean = " ".join(str(keyword).lower().split())

        if not keyword_clean:
            continue

        if keyword_clean in normalized:
            if keyword_clean not in matches:
                matches.append(keyword_clean)

    return matches


def infer_rule_family(
    rule_name: str = "",
    rule_type: str = ""
) -> str:
    """
    Map a rule to an evidence pattern family.
    """

    name = normalize_text(rule_name)
    rtype = normalize_text(rule_type)

    if (
        "compar" in name
        or "comparative" in name
        or rtype == "phrase"
    ):
        return "comparative_statement"

    if (
        "claim" in name
        or "efficacy" in name
        or "safety" in name
    ):
        return "product_claim"

    return ""


def build_evidence(
    slide_number: Any,
    slide_text: str,
    finding: Dict[str, Any]
) -> Dict[str, Any]:
    """
    Convert one rule finding into an evidence record.
    """

    rule_id = finding.get("rule_id", "")
    rule_name = finding.get("rule_name", "")
    rule_type = finding.get("rule_type", "")
    severity = finding.get("severity", "")
    action = finding.get("action", "")
    context = finding.get("context", "")
    description = finding.get("description", "")
    configured_keywords = finding.get("matched_keywords", [])

    # Exact keyword evidence
    keyword_evidence = find_keyword_evidence(
        slide_text,
        configured_keywords
    )

    # Semantic evidence
    family = infer_rule_family(
        rule_name,
        rule_type
    )

    semantic_evidence = []

    if family in SEMANTIC_PATTERNS:
        semantic_evidence = find_pattern_evidence(
            slide_text,
            SEMANTIC_PATTERNS[family]
        )

    # Combine evidence without duplicates
    detected = []

    for item in keyword_evidence + semantic_evidence:
        if item.lower() not in [
            x.lower() for x in detected
        ]:
            detected.append(item)

    # --------------------------------------------------------
    # Rule-specific reviewer guidance (screening support only)
    # --------------------------------------------------------
    normalized_name = normalize_text(rule_name)
    normalized_id = normalize_text(rule_id)
    normalized_terms = [normalize_text(x) for x in detected]
    combined = " ".join([normalized_name, normalized_id] + normalized_terms)

    if any(token in combined for token in ("comparative", "more effective", "better than", "superior", "inferior", "compared")):
        why = (
            f"The detected wording ({', '.join(detected[:4]) or rule_name}) may communicate a comparison or superiority claim. "
            "The reviewer should check what is being compared, whether the comparison is fair and balanced, and whether the cited evidence directly supports it."
        )
        recommendation = (
            "Identify the comparator and the exact outcome being compared, then check the cited study/reference for a direct, like-for-like comparison. "
            "Confirm that relevant limitations and context are presented before the material proceeds."
        )
    elif any(token in combined for token in ("claim", "efficacy", "effectiveness", "safety", "benefit", "substantiat", "rule-0001")):
        terms_text = ", ".join(detected[:5]) if detected else "the flagged claim wording"
        why = (
            f"The rule matched {terms_text}. This wording may make a product-related efficacy, safety, or other factual claim; "
            "the match alone does not establish whether the claim is accurate or acceptable."
        )
        recommendation = (
            "Locate the exact claim in the slide and trace it to the cited approved source. Check that the source supports the same product, population, "
            "outcome and wording, and that any material qualifications are visible. Escalate unresolved evidence questions to the responsible reviewer/NS."
        )
    elif any(token in combined for token in ("off-label", "indication", "dosage", "dose", "population", "route of administration", "prescribing information")):
        why = (
            f"The rule matched wording associated with approved use or prescribing information ({', '.join(detected[:4]) or rule_name}). "
            "The wording needs to be checked against the applicable locally approved product information."
        )
        recommendation = (
            "Compare the statement with the current locally approved product information, including indication, population, dose and route where relevant. "
            "Do not infer approval from the keyword match; refer any discrepancy to the responsible medical/regulatory reviewer."
        )
    elif any(token in combined for token in ("approval", "approved material", "expiry", "expiration", "expired", "renewal", "re-approval", "reapproval", "version")):
        why = (
            f"The rule matched an approval, validity or version-control indicator ({', '.join(detected[:4]) or rule_name}). "
            "This may require a check of the material's approval record and current version."
        )
        recommendation = (
            "Check the material ID and version against the authoritative approval record, including approval status and expiry/renewal date where applicable. "
            "If the material changed after approval or the record cannot be confirmed, route it to the material owner/NS before use."
        )
    elif any(token in combined for token in ("privacy", "personal data", "cybersecurity", "it security", "copyright", "consent", "permission", "image rights")):
        why = (
            f"The rule matched a privacy, security, consent or rights-related indicator ({', '.join(detected[:4]) or rule_name}). "
            "The text match does not by itself confirm a breach, but it identifies an item that may need specialist review."
        )
        recommendation = (
            "Check the specific asset or data use, the applicable permission/consent and any required privacy or security notice. "
            "Refer uncertain cases to the relevant privacy, IT security or rights owner rather than assuming the item is cleared."
        )
    elif any(token in combined for token in ("reference", "citation", "source", "evidence", "bibliography")):
        why = (
            f"The rule matched a reference or substantiation indicator ({', '.join(detected[:4]) or rule_name}). "
            "The reviewer should verify that the source is present, traceable and supports the specific statement it accompanies."
        )
        recommendation = (
            "Open the cited source and confirm that it supports the exact claim and context shown on this slide. Check that the citation is complete and that "
            "the referenced source is the appropriate approved version; flag any unsupported wording for review."
        )
    else:
        why = (
            f"The configured rule '{rule_name or rule_id}' matched {', '.join(detected[:4]) if detected else 'content on this slide'}. "
            "This is a screening signal, not a conclusion that the material is non-compliant."
        )
        recommendation = (
            f"Review the matched wording in the context of this slide and the rule's purpose ({description or rule_name or rule_id}). "
            "Confirm the relevant source, approval or process evidence, and refer any unresolved issue to the responsible reviewer/NS."
        )

    return {
        "slide_number": slide_number,
        "rule_id": rule_id,
        "rule_name": rule_name,
        "severity": severity,
        "action": action,
        "context": context,
        "detected_evidence": detected,
        "why_flagged": why,
        "recommendation": recommendation,
        "rule_description": description,
    }


def build_evidence_report(
    slides: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    """
    Build evidence records for all rule findings across slides.

    Expected slide structure:

    {
        "slide_number": 2,
        "text": "...",
        "rule_findings": [
            {...}
        ]
    }
    """

    evidence_report = []

    for slide in slides:

        slide_number = slide.get("slide_number")
        slide_text = slide.get("text", "")
        findings = slide.get("rule_findings", [])

        for finding in findings:

            evidence = build_evidence(
                slide_number=slide_number,
                slide_text=slide_text,
                finding=finding
            )

            evidence_report.append(evidence)

    return evidence_report


def attach_evidence(
    analysis_result: Dict[str, Any]
) -> Dict[str, Any]:
    """
    Attach Evidence Layer output to an existing MSA result.

    The original analysis result is preserved.
    Evidence is added under:

        result["evidence"]

    This makes the layer backward-compatible with the
    existing MSA Engine.
    """

    if not analysis_result:
        return analysis_result

    slides = analysis_result.get("slides", [])

    evidence = build_evidence_report(slides)

    analysis_result["evidence"] = evidence

    return analysis_result