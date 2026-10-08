import streamlit as st

from rule_manager import (
    load_rules,
    get_rule,
    update_rule,
    enable_rule,
    disable_rule,
)


def render_rule_management():
    """
    Render the MSA Rule Management page.

    This page allows users to:
    - View rules
    - Filter rules
    - Edit existing rules
    - Enable / Disable rules
    """

    st.title("⚙️ Rule Management")
    st.caption("Manage MSA screening rules and rule configuration.")

    # =========================================================
    # Load Rules
    # =========================================================

    rules = load_rules()

    if not rules:
        st.info("No rules are currently available.")
        return

    # =========================================================
    # Summary
    # =========================================================

    active_count = sum(
        1 for rule in rules
        if rule.get("status") == "Active"
    )

    inactive_count = sum(
        1 for rule in rules
        if rule.get("status") == "Inactive"
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric(
            "Total Rules",
            len(rules),
        )

    with col2:
        st.metric(
            "Active",
            active_count,
        )

    with col3:
        st.metric(
            "Inactive",
            inactive_count,
        )

    st.divider()

    # =========================================================
    # Filters
    # =========================================================

    st.subheader("Rule Filters")

    col1, col2 = st.columns(2)

    categories = sorted(
        {
            rule.get("category", "")
            for rule in rules
            if rule.get("category")
        }
    )

    with col1:
        category_filter = st.selectbox(
            "Category",
            ["ALL"] + categories,
        )

    with col2:
        status_filter = st.selectbox(
            "Status",
            ["ALL", "Active", "Inactive"],
        )

    # =========================================================
    # Apply Filters
    # =========================================================

    filtered_rules = rules

    if category_filter != "ALL":
        filtered_rules = [
            rule
            for rule in filtered_rules
            if rule.get("category") == category_filter
        ]

    if status_filter != "ALL":
        filtered_rules = [
            rule
            for rule in filtered_rules
            if rule.get("status") == status_filter
        ]

    st.divider()

    # =========================================================
    # Rule List
    # =========================================================

    st.subheader(
        f"Rules ({len(filtered_rules)})"
    )

    if not filtered_rules:
        st.info(
            "No rules match the selected filters."
        )
        return

    # =========================================================
    # Display Rules
    # =========================================================

    for rule in filtered_rules:

        rule_id = rule.get(
            "rule_id",
            "",
        )

        rule_name = rule.get(
            "rule_name",
            "",
        )

        category = rule.get(
            "category",
            "",
        )

        severity = rule.get(
            "severity",
            "",
        )

        status = rule.get(
            "status",
            "",
        )

        if status == "Active":
            status_display = "🟢 Active"
        else:
            status_display = "⚪ Inactive"

        # -----------------------------------------------------
        # Rule Header
        # -----------------------------------------------------

        with st.container(border=True):

            col1, col2, col3, col4, col5 = st.columns(
                [1.2, 3, 1.6, 1.2, 1.5]
            )

            with col1:
                st.write(
                    f"**{rule_id}**"
                )

            with col2:
                st.write(
                    f"**{rule_name}**"
                )

            with col3:
                st.write(category)

            with col4:
                st.write(severity)

            with col5:
                st.write(status_display)

            # -------------------------------------------------
            # Edit Area
            # -------------------------------------------------

            with st.expander(
                "View / Edit Rule"
            ):

                current_rule = get_rule(
                    rule_id
                )

                if not current_rule:
                    st.error(
                        f"Unable to load rule {rule_id}."
                    )
                    continue

                # =================================================
                # Basic Information
                # =================================================

                st.markdown(
                    "### Rule Information"
                )

                edited_name = st.text_input(
                    "Rule Name",
                    value=current_rule.get(
                        "rule_name",
                        "",
                    ),
                    key=f"name_{rule_id}",
                )

                # -------------------------------------------------
                # Category
                # -------------------------------------------------

                category_options = [
                    "PROMOMAT",
                    "NON_PROMOMAT",
                ]

                current_category = current_rule.get(
                    "category",
                    "PROMOMAT",
                )

                if current_category not in category_options:
                    category_options.append(
                        current_category
                    )

                edited_category = st.selectbox(
                    "Category",
                    category_options,
                    index=category_options.index(
                        current_category
                    ),
                    key=f"category_{rule_id}",
                )

                # -------------------------------------------------
                # Rule Type
                # -------------------------------------------------

                rule_type_options = [
                    "keyword",
                    "phrase",
                    "context",
                ]

                current_rule_type = current_rule.get(
                    "rule_type",
                    "keyword",
                )

                if current_rule_type not in rule_type_options:
                    rule_type_options.append(
                        current_rule_type
                    )

                edited_rule_type = st.selectbox(
                    "Rule Type",
                    rule_type_options,
                    index=rule_type_options.index(
                        current_rule_type
                    ),
                    key=f"type_{rule_id}",
                )

                # =================================================
                # Keywords
                # =================================================

                keywords = current_rule.get(
                    "keywords",
                    [],
                )

                keywords_text = st.text_area(
                    "Keywords / Phrases",
                    value="\n".join(keywords),
                    help=(
                        "Enter one keyword or phrase per line."
                    ),
                    key=f"keywords_{rule_id}",
                )

                # =================================================
                # Context
                # =================================================

                edited_context = st.text_area(
                    "Context",
                    value=current_rule.get(
                        "context",
                        "",
                    ),
                    key=f"context_{rule_id}",
                )

                # =================================================
                # Severity / Action
                # =================================================

                col_a, col_b = st.columns(2)

                severity_options = [
                    "Low",
                    "Medium",
                    "High",
                ]

                current_severity = current_rule.get(
                    "severity",
                    "Medium",
                )

                if current_severity not in severity_options:
                    severity_options.append(
                        current_severity
                    )

                action_options = [
                    "FLAG_FOR_REVIEW",
                    "CONTEXT_ONLY",
                ]

                current_action = current_rule.get(
                    "action",
                    "FLAG_FOR_REVIEW",
                )

                if current_action not in action_options:
                    action_options.append(
                        current_action
                    )

                with col_a:

                    edited_severity = st.selectbox(
                        "Severity",
                        severity_options,
                        index=severity_options.index(
                            current_severity
                        ),
                        key=f"severity_{rule_id}",
                    )

                with col_b:

                    edited_action = st.selectbox(
                        "Action",
                        action_options,
                        index=action_options.index(
                            current_action
                        ),
                        key=f"action_{rule_id}",
                    )

                # =================================================
                # Description
                # =================================================

                edited_description = st.text_area(
                    "Description",
                    value=current_rule.get(
                        "description",
                        "",
                    ),
                    key=f"description_{rule_id}",
                )

                # =================================================
                # Version
                # =================================================

                st.caption(
                    f"Current version: "
                    f"{current_rule.get('version', '1.0')}"
                )

                st.divider()

                # =================================================
                # Actions
                # =================================================

                col_save, col_status = st.columns(2)

                # -------------------------------------------------
                # Save Changes
                # -------------------------------------------------

                with col_save:

                    if st.button(
                        "💾 Save Changes",
                        key=f"save_{rule_id}",
                        type="primary",
                        use_container_width=True,
                    ):

                        updated_rule = {
                            "rule_name": edited_name,
                            "category": edited_category,
                            "rule_type": edited_rule_type,
                            "keywords": [
                                item.strip()
                                for item in keywords_text.splitlines()
                                if item.strip()
                            ],
                            "context": edited_context,
                            "severity": edited_severity,
                            "action": edited_action,
                            "description": edited_description,
                        }

                        try:

                            update_rule(
                                rule_id,
                                updated_rule,
                            )

                            st.success(
                                f"{rule_id} updated successfully."
                            )

                            st.rerun()

                        except Exception as e:

                            st.error(
                                f"Unable to update rule: {e}"
                            )

                # -------------------------------------------------
                # Enable / Disable
                # -------------------------------------------------

                with col_status:

                    if current_rule.get(
                        "status"
                    ) == "Active":

                        if st.button(
                            "⏸ Disable Rule",
                            key=f"disable_{rule_id}",
                            use_container_width=True,
                        ):

                            try:

                                disable_rule(
                                    rule_id
                                )

                                st.success(
                                    f"{rule_id} disabled."
                                )

                                st.rerun()

                            except Exception as e:

                                st.error(
                                    f"Unable to disable rule: {e}"
                                )

                    else:

                        if st.button(
                            "▶ Enable Rule",
                            key=f"enable_{rule_id}",
                            use_container_width=True,
                        ):

                            try:

                                enable_rule(
                                    rule_id
                                )

                                st.success(
                                    f"{rule_id} enabled."
                                )

                                st.rerun()

                            except Exception as e:

                                st.error(
                                    f"Unable to enable rule: {e}"
                                )