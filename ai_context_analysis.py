from dataclasses import dataclass
from typing import List
import re

CONTEXT_ANALYZER_VERSION = "V3"

@dataclass
class AIAnalysisResult:
    classification: str
    confidence: float
    reason: str
    matched_terms: List[str]

# Prototype context indicators only. NOT AZ compliance rules.
CONTEXT_GROUPS = {
    "business_meeting": [
        "meeting", "discussion", "follow-up", "sales rep",
        "sales representative", "representative", "customer",
        "client", "sales", "commercial"
    ],
    "promotional": [
        "better than", "compared with", "compared to", "more effective",
        "most effective", "effective", "efficacy", "safety", "product claim",
        "clinical benefit", "proven", "superior", "improves", "improvement"
    ],
    "non_promotional": [
        "meeting agenda", "internal meeting", "internal discussion",
        "training", "internal training", "administrative", "logistics",
        "process discussion", "agenda"
    ],
}

NEGATION_PATTERNS = [
    r"\bno\s+(?:\w+\s+){{0,3}}{phrase}\b",
    r"\bnot\s+(?:a\s+|an\s+|the\s+)?{phrase}\b",
    r"\bwithout\s+(?:a\s+|an\s+|the\s+)?{phrase}\b",
    r"\babsence\s+of\s+(?:a\s+|an\s+|the\s+)?{phrase}\b",
]


def normalize_text(text: str) -> str:
    if not text:
        return ""
    text = str(text).lower().replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", text).strip()


def phrase_is_negated(text: str, phrase: str) -> bool:
    text = normalize_text(text)
    phrase = normalize_text(phrase)
    escaped = re.escape(phrase)
    for template in NEGATION_PATTERNS:
        if re.search(template.format(phrase=escaped), text):
            return True
    # Simple local-window fallback for phrases such as "no product claim".
    start = text.find(phrase)
    while start >= 0:
        preceding = text[max(0, start - 45):start].split()
        if any(w in {"no", "not", "without", "none", "never"} for w in preceding[-6:]):
            return True
        start = text.find(phrase, start + 1)
    return False


def find_group_matches(text: str, group: str) -> List[str]:
    normalized = normalize_text(text)
    result = []
    for term in CONTEXT_GROUPS.get(group, []):
        if term in normalized and not phrase_is_negated(normalized, term):
            result.append(term)
    return result


def find_matched_terms(text: str) -> List[str]:
    result = []
    for group in CONTEXT_GROUPS:
        for term in find_group_matches(text, group):
            if term not in result:
                result.append(term)
    return result


def has_negated_promotional_language(text: str) -> List[str]:
    normalized = normalize_text(text)
    return [
        term for term in CONTEXT_GROUPS["promotional"]
        if term in normalized and phrase_is_negated(normalized, term)
    ]


def analyze_slide(text: str) -> AIAnalysisResult:
    normalized = normalize_text(text)
    if not normalized:
        return AIAnalysisResult(
            "UNCERTAIN", 0.0,
            "No readable text was available for contextual analysis.", []
        )

    promo = find_group_matches(normalized, "promotional")
    business = find_group_matches(normalized, "business_meeting")
    nonpromo = find_group_matches(normalized, "non_promotional")
    negated_promo = has_negated_promotional_language(normalized)

    strong_promo = [
        x for x in promo if x in {
            "better than", "compared with", "compared to", "more effective",
            "most effective", "efficacy", "safety", "product claim",
            "clinical benefit", "proven", "superior"
        }
    ]

    # Positive promotional statements take priority.
    if strong_promo:
        confidence = min(0.98, 0.72 + 0.06 * len(strong_promo) + (0.03 if business else 0))
        return AIAnalysisResult(
            "PROMOTIONAL_CONTEXT", round(confidence, 2),
            "The slide contains promotional/product-related indicators that may require promotional review.",
            promo
        )

    # Sales/customer context = business meeting when no strong promotional signal exists.
    business_core = [x for x in business if x in {
        "sales rep", "sales representative", "representative",
        "customer", "client", "sales", "commercial"
    }]
    if business_core:
        confidence = min(0.96, 0.78 + 0.05 * len(business_core) + (0.03 if "meeting" in business else 0))
        return AIAnalysisResult(
            "BUSINESS_MEETING", round(confidence, 2),
            "The slide contains business meeting context with sales or commercial indicators.",
            business
        )

    # Internal/training/agenda context explicitly wins over weak product mentions.
    internal_core = [x for x in nonpromo if x in {
        "meeting agenda", "internal meeting", "internal discussion",
        "training", "internal training", "administrative", "logistics",
        "process discussion"
    }]
    if internal_core:
        confidence = min(0.97, 0.82 + 0.04 * len(internal_core) + (0.03 if negated_promo else 0))
        reason = "The slide appears to contain internal, training, agenda, or non-promotional meeting content."
        if negated_promo:
            reason += " Promotional language was detected only in an explicitly negated context."
        return AIAnalysisResult(
            "NON_PROMOTIONAL_CONTEXT", round(confidence, 2), reason, nonpromo
        )

    if business:
        return AIAnalysisResult(
            "BUSINESS_MEETING", 0.72,
            "The slide contains meeting or discussion language, but there is not enough evidence to establish promotional context.",
            business
        )

    matches = find_matched_terms(normalized)
    if matches:
        return AIAnalysisResult(
            "UNCERTAIN", 0.55,
            "Some contextual indicators were detected, but there is not enough evidence to determine the primary context.",
            matches
        )

    return AIAnalysisResult(
        "NOT_CLASSIFIED", 0.90,
        "No meaningful contextual indicators were identified in the available slide text.", []
    )


def analyze_slides(slides: List[dict]) -> List[dict]:
    results = []
    for slide in slides:
        analysis = analyze_slide(slide.get("text", ""))
        results.append({
            "slide_number": slide.get("slide_number"),
            "classification": analysis.classification,
            "confidence": analysis.confidence,
            "reason": analysis.reason,
            "matched_terms": analysis.matched_terms,
        })
    return results


if __name__ == "__main__":
    tests = [
        ("CASE 1 - BUSINESS MEETING",
         "Sales rep meeting with customer to discuss product information and follow-up actions."),
        ("CASE 2 - PROMOTIONAL",
         "The product is better than the comparator and has strong efficacy and safety."),
        ("CASE 3 - NON-PROMOTIONAL",
         "Internal meeting agenda and training discussion. No product claim is presented."),
    ]

    print("=" * 70)
    print("MSA CONTEXT ANALYSIS V2 TEST")
    print("=" * 70)
    for title, text in tests:
        r = analyze_slide(text)
        print("\n" + title)
        print("-" * 70)
        print("Classification:", r.classification)
        print("Confidence   :", r.confidence)
        print("Reason       :", r.reason)
        print("Matched terms:", r.matched_terms)
