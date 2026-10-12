import json
from pathlib import Path
from typing import Iterable

RULES_FILE = Path(__file__).parent / "rules.json"
APPROVED_STATES = {"APPROVED", "NS_APPROVED", "APPROVED_BY_NS", "APPROVED_BY_SOP_OWNER"}


def _read_data():
    if not RULES_FILE.exists():
        raise FileNotFoundError(f"Rules file not found: {RULES_FILE}")
    with RULES_FILE.open("r", encoding="utf-8") as file:
        data = json.load(file)
    if not isinstance(data, dict) or not isinstance(data.get("rules", []), list):
        raise ValueError("rules.json must contain a top-level 'rules' list.")
    return data


def _write_data(data):
    temporary = RULES_FILE.with_suffix(".json.tmp")
    with temporary.open("w", encoding="utf-8") as file:
        json.dump(data, file, indent=2, ensure_ascii=False)
    temporary.replace(RULES_FILE)


def load_rules():
    return _read_data().get("rules", [])


def get_rule(rule_id):
    return next((r for r in load_rules() if r.get("rule_id") == rule_id), None)


def _approval_is_allowed(rule):
    state = str(rule.get("approval_status", "")).strip().upper()
    # Legacy prototype rules without an approval_status retain their current behaviour.
    return not state or state in APPROVED_STATES


def get_active_rules():
    return [r for r in load_rules() if str(r.get("status", "")).strip().lower() == "active" and _approval_is_allowed(r)]


def add_rule(rule):
    if not isinstance(rule, dict):
        raise TypeError("rule must be a dictionary")
    rule = dict(rule)
    rule_id = str(rule.get("rule_id") or "").strip()
    if not rule_id:
        raise ValueError("rule_id is required")
    if get_rule(rule_id) is not None:
        raise ValueError(f"Rule already exists: {rule_id}")
    if not _approval_is_allowed(rule):
        rule["status"] = "Inactive"
    data = _read_data()
    data.setdefault("rules", []).append(rule)
    _write_data(data)
    return rule


def update_rule(rule_id, updates):
    if not isinstance(updates, dict):
        raise TypeError("updates must be a dictionary")
    data = _read_data()
    for rule in data.get("rules", []):
        if rule.get("rule_id") == rule_id:
            update = dict(updates)
            update.pop("rule_id", None)
            rule.update(update)
            if not _approval_is_allowed(rule):
                rule["status"] = "Inactive"
            _write_data(data)
            return rule
    raise ValueError(f"Rule not found: {rule_id}")


def delete_rule(rule_id):
    data = _read_data()
    rules = data.get("rules", [])
    filtered = [r for r in rules if r.get("rule_id") != rule_id]
    if len(filtered) == len(rules):
        raise ValueError(f"Rule not found: {rule_id}")
    data["rules"] = filtered
    _write_data(data)
    return True


def set_rule_status(rule_id, status):
    if status not in ("Active", "Inactive"):
        raise ValueError("status must be 'Active' or 'Inactive'")
    rule = get_rule(rule_id)
    if rule is None:
        raise ValueError(f"Rule not found: {rule_id}")
    if status == "Active" and not _approval_is_allowed(rule):
        raise ValueError("Rule cannot be activated until approval_status is APPROVED.")
    return update_rule(rule_id, {"status": status})


def enable_rule(rule_id):
    return set_rule_status(rule_id, "Active")


def disable_rule(rule_id):
    return set_rule_status(rule_id, "Inactive")


def upsert_rules(imported_rules):
    """Merge validated catalog records by rule_id, preserving unrelated rules."""
    if not isinstance(imported_rules, list) or not imported_rules:
        raise ValueError("No rules were supplied for import.")
    data = _read_data()
    rules = data.setdefault("rules", [])
    index = {str(r.get("rule_id")): i for i, r in enumerate(rules) if isinstance(r, dict) and r.get("rule_id")}
    created = updated = 0
    seen = set()
    for incoming in imported_rules:
        if not isinstance(incoming, dict) or not str(incoming.get("rule_id") or "").strip():
            raise ValueError("Every imported rule must have a rule_id.")
        rule = dict(incoming)
        rid = str(rule["rule_id"]).strip()
        if rid in seen:
            raise ValueError(f"Duplicate rule_id in uploaded workbook: {rid}")
        seen.add(rid)
        rule["rule_id"] = rid
        state = str(rule.get("approval_status", "PENDING_NS_SOP_OWNER_REVIEW")).strip().upper()
        rule["approval_status"] = state
        if state not in APPROVED_STATES:
            rule["status"] = "Inactive"
        else:
            rule["status"] = "Active" if str(rule.get("status", "Inactive")).strip().upper() == "ACTIVE" else "Inactive"
        if rid in index:
            # Catalog import updates matching catalog-managed IDs but does not delete unrelated rules.
            rules[index[rid]].update(rule)
            updated += 1
        else:
            index[rid] = len(rules)
            rules.append(rule)
            created += 1
    _write_data(data)
    return {"created": created, "updated": updated, "upserted": created + updated}


def count_rules():
    return len(load_rules())


def count_active_rules():
    return len(get_active_rules())
