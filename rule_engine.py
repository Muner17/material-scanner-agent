import re
from rule_manager import get_active_rules


# Deterministic text cues for rules whose catalog entries intentionally have no
# keyword list. These are screening indicators only; they do not establish
# approval, off-label status, or compliance.
_CONTEXT_CUES = {
    "Rule-0002": [
        ("indication", r"\bindication(s)?\b|\bindicated for\b"),
        ("dosage", r"\bdos(e|age|ing)\b|\bmg\s*/\s*(day|daily)\b|\bonce daily\b|\btwice daily\b"),
        ("population", r"\b(patient population|population|paediatric|pediatric|adult patients|children)\b"),
        ("route", r"\b(route of administration|administered orally|oral administration|intravenous|subcutaneous)\b"),
        ("use", r"\b(used for|use in|treatment of|treating|indicated to)\b"),
        ("approved label", r"\b(approved label|prescribing information|product label)\b"),
    ],
    "Rule-0003": [
        ("promotional", r"\b(promotional material|promotional|promote|promotion)\b"),
        ("HCP", r"\b(HCPs?|healthcare professionals|health care professionals)\b"),
        ("branded product", r"\b(brand(ed)? product|product claim|brand name)\b"),
    ],
    "Rule-0004": [
        ("patient education", r"\b(patient education|patient information|patient leaflet|educational material for patients)\b"),
        ("brand/logo", r"\b(brand(ed)?|logo|company logo|brand imagery|brand quote)\b"),
        ("generic name", r"\b(generic name|generic medicine|active ingredient)\b"),
    ],
    "Rule-0005": [
        ("speaker slides", r"\b(speaker slides?|speaker presentation|external speaker|speaker deck)\b"),
        ("brand/product", r"\b(brand(ed)?|brand name|product logo|product name|brand quote)\b"),
    ],
    "Rule-0006": [
        ("non-promotional external meeting", r"\b(non[- ]promotional|nonpromotional)\b.{0,50}\b(external meeting|speaker|event)\b|\b(external meeting|speaker event)\b"),
        ("brand/promotion", r"\b(brand name|branded product|promotional message|off[- ]label|promotional claim)\b"),
    ],
}

def _normalize_category(value):
    return str(value or "").strip().upper().replace("-", "_").replace(" ", "_")


def _applies_to(rule):
    raw = rule.get("applies_to", rule.get("category", ""))
    if isinstance(raw, str):
        values = re.split(r"[;,|]", raw)
    elif isinstance(raw, (list, tuple, set)):
        values = list(raw)
    else:
        values = []
    values.extend([rule.get("category", "")])
    return {_normalize_category(v) for v in values if str(v or "").strip()}


def _matches(text, term, rule_type):
    term = str(term or "").strip()
    if not term:
        return False
    if re.search(r"\s", term):
        phrase_pattern = r"\s+".join(re.escape(part) for part in term.split())
        pattern = r"(?<!\w)" + phrase_pattern + r"(?!\w)"
    else:
        pattern = r"(?<!\w)" + re.escape(term) + r"(?!\w)"
    return re.search(pattern, text, flags=re.IGNORECASE) is not None


def _context_matches(rule, text):
    rule_id = str(rule.get("rule_id", "")).strip()
    cues = _CONTEXT_CUES.get(rule_id, [])
    found = []
    for label, pattern in cues:
        if re.search(pattern, text, flags=re.IGNORECASE):
            found.append(label)
    # A context rule is triggered only when enough distinct contextual signals
    # are present; generic single-word matches are intentionally insufficient.
    if rule_id == "Rule-0002":
        # Off-label screening needs more than one broad contextual signal.
        # A generic mention of "indication" alone is not sufficient.
        required = 2
    elif rule_id == "Rule-0003":
        required = 2
    elif rule_id == "Rule-0004":
        required = 2
    elif rule_id == "Rule-0005":
        required = 2
    elif rule_id == "Rule-0006":
        required = 2
    else:
        required = 99
    return found if len(found) >= required else []


def evaluate_rules(text, category=None):
    """Return active/approved rule matches for one slide.

    Keyword/phrase rules use their configured terms. Context rules without
    keywords use conservative, rule-specific textual cues. Metadata-only rules
    are not inferred from slide text and therefore are not emitted as findings.
    All outputs remain screening indicators for human review, not decisions.
    """
    if not text or not str(text).strip():
        return []
    text = str(text)
    category_key = _normalize_category(category) if category else None
    matched_rules = []
    for rule in get_active_rules():
        if not isinstance(rule, dict):
            continue
        applies = _applies_to(rule)
        if category_key and category_key not in applies and "BOTH" not in applies:
            continue

        # Rule-0003 is a profile-mismatch check, not a general promotional-content
        # detector. Promotional cues are not a mismatch when PROMOMAT is selected.
        if rule.get("rule_id") == "Rule-0003" and category_key == "PROMOMAT":
            continue

        rule_type = str(rule.get("rule_type", "keyword")).strip().lower()
        raw_keywords = rule.get("keywords", [])
        if isinstance(raw_keywords, str):
            raw_keywords = re.split(r"[;\n]", raw_keywords)
        if not isinstance(raw_keywords, (list, tuple)):
            raw_keywords = []

        matched_keywords = []

        # Rule-0003 is a profile-mismatch check. Do not trigger it from a
        # single configured keyword such as "promotional". Use the rule-specific
        # context cues and require multiple distinct signals instead.
        if rule.get("rule_id") == "Rule-0003":
            matched_keywords = _context_matches(rule, text)
        else:
            for keyword in raw_keywords:
                if _matches(text, keyword, rule_type):
                    label = str(keyword).strip()
                    if label and label.casefold() not in {x.casefold() for x in matched_keywords}:
                        matched_keywords.append(label)

        # Rule-0002 contains broad generic terms as well as stronger indicators.
        # Generic terms alone must be corroborated by multiple distinct context cues.
        if rule.get("rule_id") == "Rule-0002" and matched_keywords:
            generic_terms = {
                "indication", "dosage", "dose", "population",
                "route of administration", "treatment", "use",
            }
            if all(str(item).strip().casefold() in generic_terms for item in matched_keywords):
                matched_keywords = _context_matches(rule, text)

        # Only context rules with no configured terms use the conservative cue map.
        if not matched_keywords and rule_type == "context" and not raw_keywords:
            if rule.get("rule_id") != "Rule-0003" or category_key == "NON_PROMOMAT":
                matched_keywords = _context_matches(rule, text)

        if not matched_keywords:
            continue

        result = {
            "rule_id": rule.get("rule_id"),
            "rule_name": rule.get("rule_name"),
            "category": rule.get("category"),
            "applies_to": sorted(applies),
            "rule_type": rule.get("rule_type"),
            "matched_keywords": matched_keywords,
            "severity": rule.get("severity", "Medium"),
            "action": rule.get("action", "FLAG_FOR_REVIEW"),
            "context": rule.get("context", ""),
            "description": rule.get("description", ""),
            "version": rule.get("version", rule.get("rule_version", "1.0")),
        }
        for field in ("source_document", "source_version_status", "source_section_page",
                      "effective_date", "approval_status", "topic",
                      "recommended_action", "message", "proposed_detection_logic"):
            if field in rule:
                result[field] = rule[field]
        matched_rules.append(result)
    return matched_rules
