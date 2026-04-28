import json

# §1 Constants & Exceptions

# §2 Classes and Sub Classes

# §3 Private Helper Functions

# §4 Public Functions


def test_roundtrip_schema_identical(fixture_path, managed_tmp_dir):
    """extract -> bundle export -> re-extract -> schemas must be deeply equal."""
    from mindoff_dataport import mo_dataport

    schema_1 = mo_dataport.extract(fixture_path)
    built_path = str(managed_tmp_dir / "rebuilt.xlsx")
    data = {sheet["name"]: {} for sheet in schema_1["sheets"]}
    bundle = mo_dataport.compile(schema_1, data)
    mo_dataport.export(bundle, built_path)
    schema_2 = mo_dataport.extract(built_path)

    for sheet in schema_1["sheets"]:
        sheet["merged_regions"].sort()
    for sheet in schema_2["sheets"]:
        sheet["merged_regions"].sort()

    assert json.dumps(schema_1, sort_keys=True) == json.dumps(schema_2, sort_keys=True)


# §5 Entrypoints
