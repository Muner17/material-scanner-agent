from typing import Dict, Any

from pptx_extractor import extract_pptx
from msa_engine import analyze_material
from evidence_layer import attach_evidence


def analyze_pptx(
    file_path: str,
    category: str
) -> Dict[str, Any]:
    """
    Complete MSA analysis pipeline.

    Flow:
        PPTX
        -> Extract slides
        -> Context Analysis
        -> Rule Engine
        -> MSA Engine
        -> Evidence Layer
    """

    # --------------------------------------------------------
    # 1. Extract PPTX
    # --------------------------------------------------------

    slides = extract_pptx(file_path)

    if not slides:
        return {
            "category": category,
            "slide_count": 0,
            "slides": [],
            "evidence": [],
            "summary": {
                "slides_analyzed": 0,
                "rule_findings": 0,
                "high": 0,
                "medium": 0,
                "low": 0,
            },
            "error": "No readable slides were found."
        }

    # --------------------------------------------------------
    # 2. Run existing MSA Engine
    # --------------------------------------------------------

    result = analyze_material(
        slides,
        category
    )

    # --------------------------------------------------------
    # 3. Add Evidence Layer
    # --------------------------------------------------------

    result = attach_evidence(result)

    # --------------------------------------------------------
    # 4. Return complete result
    # --------------------------------------------------------

    return result


if __name__ == "__main__":

    print("=" * 50)
    print("MSA COMPLETE PIPELINE TEST")
    print("=" * 50)

    file_path = "mock_sales_rep_material.pptx"
    category = "PROMOMAT"

    result = analyze_pptx(
        file_path,
        category
    )

    print()
    print("Category:", result.get("category"))
    print("Slides:", result.get("slide_count"))

    print()
    print("Summary:")
    print(result.get("summary"))

    print()
    print("Evidence Findings:")
    print("-" * 50)

    for evidence in result.get("evidence", []):

        print(
            f"Slide {evidence.get('slide_number')} | "
            f"{evidence.get('rule_id')} | "
            f"{evidence.get('rule_name')} | "
            f"Severity={evidence.get('severity')}"
        )

        print(
            "Detected:",
            ", ".join(
                evidence.get("detected_evidence", [])
            )
        )

        print(
            "Why:",
            evidence.get("why_flagged")
        )

        print(
            "Recommendation:",
            evidence.get("recommendation")
        )

        print("-" * 50)