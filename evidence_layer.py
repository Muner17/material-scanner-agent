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
    # Why was this flagged?
    # --------------------------------------------------------

    if "comparative" in normalize_text(rule_name):
        why = (
            "Comparative language was detected in the material. "
            "The statement may require verification against the "
            "applicable supporting evidence or reference."
        )

    elif "claim" in normalize_text(rule_name):
        why = (
            "Product-related claim language was detected. "
            "The statement may require verification against the "
            "applicable approved evidence or reference."
        )

    else:
        why = (
            "Content matching the configured review rule was "
            "detected and should be reviewed against the "
            "applicable reference."
        )

    # --------------------------------------------------------
    # Reviewer recommendation
    # --------------------------------------------------------

    if severity == "High":
        recommendation = (
            "Verify the statement against the applicable "
            "approved reference/evidence before proceeding."
        )

    elif severity == "Medium":
        recommendation = (
            "Review the detected content and confirm that it "
            "is supported by the applicable reference."
        )

    else:
        recommendation = (
            "Review the detected content in context."
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