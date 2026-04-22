import json

# §1 Constants & Exceptions

# §2 Classes and Sub Classes

# §3 Private Helper Functions

# §4 Public Functions


def test_roundtrip_schema_identical(fixture_path, managed_tmp_dir):
    """extract -> build -> re-extract -> schemas must be deeply equal."""
    from mindoff_data_export import extract_template
    from mindoff_data_export.builder import build_template

    schema_1 = extract_template(fixture_path)
    built_path = str(managed_tmp_dir / "rebuilt.xlsx")
    build_template(schema_1, built_path)
    schema_2 = extract_template(built_path)

    for sheet in schema_1["sheets"]:
        sheet["merged_regions"].sort()
    for sheet in schema_2["sheets"]:
        sheet["merged_regions"].sort()

    assert json.dumps(schema_1, sort_keys=True) == json.dumps(schema_2, sort_keys=True)


# §5 Entrypoints
