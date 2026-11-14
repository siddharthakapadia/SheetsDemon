"""sheetdoc — an Excel tutor, auditor and builder for intro spreadsheet courses."""

from .workbook import load  # noqa: F401

__version__ = "1.0.0"

__all__ = [
    "load",
    # 1. auditing
    "detect_hardcoded_values", "check_reference_types",
    "validate_range_bounds", "verify_lookup_syntax",
    # 2. formatting
    "audit_number_formats", "apply_conditional_formatting_rule",
    "check_header_and_grid_structure",
    # 3. tutoring
    "hint_generator_socratic", "explain_error_code", "generate_practice_variant",
    # 4. data and analysis
    "build_mock_dataset", "explain_pivot_table_setup", "chart_type_recommender",
    # 5. table rows
    "add_worksheet_tablerow", "update_worksheet_tablerow",
]


def __getattr__(name: str):
    """Lazily expose the fourteen public functions at package level."""
    for module in ("audit", "formatting", "tutor", "datagen", "builder"):
        mod = __import__(f"sheetdoc.{module}", fromlist=[name])
        if hasattr(mod, name):
            return getattr(mod, name)
    raise AttributeError(f"module 'sheetdoc' has no attribute {name!r}")
