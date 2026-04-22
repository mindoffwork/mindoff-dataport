import json


def test_roundtrip_schema_identical(fixture_path, tmp_path):
    """
    extract → build → re-extract → schemas must be deeply equal.
    This is the gold-standard fidelity test.
    """
    from mindoff_data_export import build_template, extract_template

    schema_1 = extract_template(fixture_path)
    built_path = str(tmp_path / "rebuilt.xlsx")
    build_template(schema_1, built_path)
    schema_2 = extract_template(built_path)

    for sheet in schema_1["sheets"]:
        sheet["merged_regions"].sort()
    for sheet in schema_2["sheets"]:
        sheet["merged_regions"].sort()

    assert json.dumps(schema_1, sort_keys=True) == json.dumps(schema_2, sort_keys=True)
