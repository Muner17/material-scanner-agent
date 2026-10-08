from rule_manager import get_active_rules


def evaluate_rules(text, category=None):
    """
    Evaluate active MSA rules against extracted slide text.

    Args:
        text (str): Extracted text from slide/document.
        category (str, optional): PROMOMAT or NON_PROMOMAT.

    Returns:
        list: Matched rules.
    """

    if not text:
        return []

    text_lower = text.lower()

    active_rules = get_active_rules()
    matched_rules = []

    for rule in active_rules:

        # Category filter
        rule_category = rule.get("category", "").upper()

        if category:
            if rule_category != category.upper():
                continue

        keywords = rule.get("keywords", [])

        matched_keywords = []

        for keyword in keywords:
            if keyword.lower() in text_lower:
                matched_keywords.append(keyword)

        if matched_keywords:

            matched_rules.append(
                {
                    "rule_id": rule.get("rule_id"),
                    "rule_name": rule.get("rule_name"),
                    "category": rule.get("category"),
                    "rule_type": rule.get("rule_type"),
                    "matched_keywords": matched_keywords,
                    "severity": rule.get("severity"),
                    "action": rule.get("action"),
                    "context": rule.get("context"),
                    "description": rule.get("description"),
                    "version": rule.get("version"),
                }
            )

    return matched_rules