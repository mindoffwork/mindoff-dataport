# Roadmap

Tracks where `mindoff-dataport` stands today and what is likely to come next.

It is not a release promise. It is the clearest current picture of direction.

## Current Status

- [x] Extract `.xlsx` templates into a reusable schema
- [x] Inspect template inputs with a sheet-scoped contract
- [x] Compile runtime data into a portable `ReportBundle`
- [x] Export compiled bundles to `.xlsx`
- [x] Export compiled bundles to `.pdf`
- [x] Reuse one compiled bundle across both `.xlsx` and `.pdf` outputs
- [x] Support dataframe placeholders in templates
- [x] Support repeat sections and dynamic sheets
- [x] Support streaming XLSX export for large datasets
- [x] Preserve template-driven layout, styles, merges, and page-break-aware rendering
- [x] Support PDF pagination and optional custom fonts

## Planned

### Near Term

- [ ] CSV export
- [ ] Excel-friendly CSV export presets or helpers
- [ ] Google Sheets export
- [ ] Image export

### PDF Improvements

- [ ] Flatten PDF option
- [ ] Compressed PDF option

### Codebase Improvements

- [ ] Leaner internal architecture
- [ ] Better-organized modules and file boundaries
- [ ] More minimalist contributor-facing structure

## Notes

- CSV and Excel-friendly CSV export are related, but they may not land as the same API shape.
- Google Sheets export likely needs a different write path than `.xlsx` and `.pdf`.
- Image export needs a tighter definition before implementation starts. It could mean sheet snapshots, page rendering, or selected-region export.
- PDF flattening and PDF compression may involve fidelity and file-size trade-offs.
