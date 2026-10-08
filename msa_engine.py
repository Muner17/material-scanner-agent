from typing import Any, Dict, List

from ai_context_analysis import analyze_slide
from rule_engine import evaluate_rules

try:
    from rule_manager import load_rules
except Exception:
    load_rules = None


def build_profile_mapping(category: str) -> Dict[str, Any]:
    """Return the active rules mapped to the selected review profile.

    This is a configuration view for the prototype. It does not create
    compliance decisions; it reports which active rules are configured
    for the selected profile.
    """
    normalized = str(category or "").strip().upper()
    if normalized == "NON-PROMOMAT":
        normalized = "NON_PROMOMAT"

    if normalized not in {"PROMOMAT", "NON_PROMOMAT"}:
        return {
            "profile": normalized,
            "rule_count": 0,
            "rules": [],
        }

    rules = []
    if load_rules is not None:
        try:
            all_rules = load_rules()
        except Exception:
            all_rules = []
        for rule in all_rules if isinstance(all_rules, list) else []:
            if not isinstance(rule, dict):
                continue
            rule_category = str(rule.get("category", "")).strip().upper()
            status = str(rule.get("status", "")).strip().lower()
            if rule_category == normalized and status == "active":
                rules.append({
                    "rule_id": rule.get("rule_id", ""),
                    "rule_name": rule.get("rule_name", ""),
                    "rule_type": rule.get("rule_type", ""),
                    "severity": rule.get("severity", ""),
                    "action": rule.get("action", ""),
                    "version": rule.get("version", "1.0"),
                    "description": rule.get("description", ""),
                })

    return {
        "profile": normalized,
        "rule_count": len(rules),
        "rules": rules,
    }


def analyze_material(slides: List[dict], category: str) -> Dict[str, Any]:
    """
    Run the MSA prototype analysis pipeline.

    Pipeline:
        Slide text
            -> Context Analysis
            -> Rule Engine
            -> Combined finding

    Expected slide format:
        [
            {"slide_number": 1, "text": "..."},
            {"slide_number": 2, "text": "..."}
        ]

    category:
        "PROMOMAT" or "NON_PROMOMAT"

    This is a prototype pre-review engine.
    It does NOT make a final compliance decision.
    """

    if not isinstance(slides, list):
        raise TypeError("slides must be a list of dictionaries")

    category = str(category).strip().upper()

    if category not in {"PROMOMAT", "NON_PROMOMAT"}:
        raise ValueError(
            "category must be 'PROMOMAT' or 'NON_PROMOMAT'"
        )

    profile_mapping = build_profile_mapping(category)
    results: List[Dict[str, Any]] = []

    for index, slide in enumerate(slides, start=1):
        if not isinstance(slide, dict):
            continue

        slide_number = slide.get("slide_number", index)
        text = str(slide.get("text", "") or "")

        # ---------------------------------------------------------------
        # 1. Context Analysis
        # ---------------------------------------------------------------
        context_result = analyze_slide(text)

        # ---------------------------------------------------------------
        # 2. Rule Engine
        # ---------------------------------------------------------------
        rule_results = evaluate_rules(text, category)

        # Make sure a single rule-engine result never breaks the pipeline.
        if rule_results is None:
            rule_results = []

        # ---------------------------------------------------------------
        # 3. Combine both layers
        # ---------------------------------------------------------------
        results.append(
            {
                "slide_number": slide_number,
                "text": text,
                "context": {
                    "classification": context_result.classification,
                    "confidence": context_result.confidence,
                    "reason": context_result.reason,
                    "matched_terms": context_result.matched_terms,
                },
                "rule_findings": rule_results,
            }
        )

    return {
        "category": category,
        "slide_count": len(results),
        "slides": results,
        "summary": build_summary(results),
        "profile_mapping": profile_mapping,
    }


def build_summary(results: List[Dict[str, Any]]) -> Dict[str, int]:
    """
    Build a simple summary for the Streamlit UI.

    The summary is intentionally derived from the rule engine output
    rather than making new compliance decisions.
    """

    total_rule_findings = 0
    high = 0
    medium = 0
    low = 0

    for slide in results:
        findings = slide.get("rule_findings", [])

        if not isinstance(findings, list):
            continue

        total_rule_findings += len(findings)

        for finding in findings:
            if not isinstance(finding, dict):
                continue

            severity = str(
                finding.get("severity", "")
            ).strip().upper()

            if severity == "HIGH":
                high += 1
            elif severity == "MEDIUM":
                medium += 1
            elif severity == "LOW":
                low += 1

    return {
        "slides_analyzed": len(results),
        "rule_findings": total_rule_findings,
        "high": high,
        "medium": medium,
        "low": low,
    }


def print_analysis(result: Dict[str, Any]) -> None:
    """Pretty-print the pipeline result for local testing."""

    print("")
    print("========================================")
    print("MSA MATERIAL ANALYSIS")
    print("========================================")
    print(f"Category       : {result['category']}")
    print(f"Slides analyzed: {result['summary']['slides_analyzed']}")
    print(f"Rule findings  : {result['summary']['rule_findings']}")
    print(f"High           : {result['summary']['high']}")
    print(f"Medium         : {result['summary']['medium']}")
    print(f"Low            : {result['summary']['low']}")
    print("")

    for slide in result["slides"]:
        print("----------------------------------------")
        print(f"Slide {slide['slide_number']}")

        context = slide["context"]

        print(
            f"Context        : {context['classification']}"
        )
        print(
            f"Context score  : {context['confidence']}"
        )

        if context["matched_terms"]:
            print(
                "Context terms  : "
                + ", ".join(context["matched_terms"])
            )

        findings = slide["rule_findings"]

        if not findings:
            print("Rule findings  : None")
            continue

        print("Rule findings  :")

        for finding in findings:
            print(
                f"  - {finding.get('rule_id', 'UNKNOWN')} | "
                f"{finding.get('rule_name', 'Unknown Rule')} | "
                f"severity={finding.get('severity', 'Unknown')} | "
                f"action={finding.get('action', 'Unknown')}"
            )

            matched = finding.get("matched_keywords", [])

            if matched:
                print(
                    "    matched: "
                    + ", ".join(str(item) for item in matched)
                )


if __name__ == "__main__":
    # ---------------------------------------------------------------
    # Demo scenario representing the type of Sales Rep material
    # we want to test.
    #
    # This is mock/demo content only.
    # ---------------------------------------------------------------
    demo_slides = [
        {
            "slide_number": 1,
            "text": (
                "Sales representative meeting with customer. "
                "Discussion of product information and next steps."
            ),
        },
        {
            "slide_number": 2,
            "text": (
                "Product X is better than Product Y and has "
                "strong safety and efficacy."
            ),
        },
        {
            "slide_number": 3,
            "text": (
                "Internal meeting agenda and training discussion. "
                "No product claim is presented."
            ),
        },
    ]

    result = analyze_material(
        demo_slides,
        "PROMOMAT",
    )

    print_analysis(result)
