import json
from pathlib import Path


RULES_FILE = Path(__file__).parent / "rules.json"


# =========================================================
# Internal helpers
# =========================================================

def _read_data():
    """Read rules.json and return the full data structure."""
    if not RULES_FILE.exists():
        raise FileNotFoundError(
            f"Rules file not found: {RULES_FILE}"
        )

    with RULES_FILE.open("r", encoding="utf-8") as file:
        return json.load(file)


def _write_data(data):
    """Write the full data structure back to rules.json."""
    with RULES_FILE.open("w", encoding="utf-8") as file:
        json.dump(
            data,
            file,
            indent=2,
            ensure_ascii=False
        )


# =========================================================
# Read operations
# =========================================================

def load_rules():
    """
    Load all rules from rules.json.

    Returns:
        list: A list of rule dictionaries.
    """
    data = _read_data()
    return data.get("rules", [])


def get_rule(rule_id):
    """
    Get a single rule by rule_id.

    Returns:
        dict or None
    """
    rules = load_rules()

    for rule in rules:
        if rule.get("rule_id") == rule_id:
            return rule

    return None


def get_active_rules():
    """
    Return only rules with Active status.

    Returns:
        list: Active rules.
    """
    return [
        rule
        for rule in load_rules()
        if rule.get("status") == "Active"
    ]


# =========================================================
# Create
# =========================================================

def add_rule(rule):
    """
    Add a new rule.

    Raises:
        ValueError: If rule_id already exists.
    """
    if not isinstance(rule, dict):
        raise TypeError("rule must be a dictionary")

    rule_id = rule.get("rule_id")

    if not rule_id:
        raise ValueError("rule_id is required")

    existing_rule = get_rule(rule_id)

    if existing_rule is not None:
        raise ValueError(
            f"Rule already exists: {rule_id}"
        )

    data = _read_data()
    data.setdefault("rules", [])
    data["rules"].append(rule)

    _write_data(data)

    return rule


# =========================================================
# Update
# =========================================================

def update_rule(rule_id, updates):
    """
    Update an existing rule.

    rule_id cannot be changed.

    Returns:
        dict: Updated rule.

    Raises:
        ValueError: If rule does not exist.
    """
    if not isinstance(updates, dict):
        raise TypeError("updates must be a dictionary")

    data = _read_data()

    for rule in data.get("rules", []):
        if rule.get("rule_id") == rule_id:

            # Prevent accidental rule_id changes
            updates = dict(updates)
            updates.pop("rule_id", None)

            rule.update(updates)

            _write_data(data)

            return rule

    raise ValueError(
        f"Rule not found: {rule_id}"
    )


# =========================================================
# Delete
# =========================================================

def delete_rule(rule_id):
    """
    Delete a rule by rule_id.

    Returns:
        bool: True if deleted.

    Raises:
        ValueError: If rule does not exist.
    """
    data = _read_data()
    rules = data.get("rules", [])

    original_count = len(rules)

    data["rules"] = [
        rule
        for rule in rules
        if rule.get("rule_id") != rule_id
    ]

    if len(data["rules"]) == original_count:
        raise ValueError(
            f"Rule not found: {rule_id}"
        )

    _write_data(data)

    return True


# =========================================================
# Status management
# =========================================================

def set_rule_status(rule_id, status):
    """
    Change rule status.

    Allowed values:
        Active
        Inactive

    Returns:
        dict: Updated rule.
    """
    if status not in ("Active", "Inactive"):
        raise ValueError(
            "status must be 'Active' or 'Inactive'"
        )

    return update_rule(
        rule_id,
        {"status": status}
    )


def enable_rule(rule_id):
    """Enable a rule."""
    return set_rule_status(
        rule_id,
        "Active"
    )


def disable_rule(rule_id):
    """Disable a rule."""
    return set_rule_status(
        rule_id,
        "Inactive"
    )


# =========================================================
# Utility
# =========================================================

def count_rules():
    """Return the total number of rules."""
    return len(load_rules())


def count_active_rules():
    """Return the number of active rules."""
    return len(get_active_rules())