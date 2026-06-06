# Architecture

The inner workings of `mindoff-dataport`: how a styled workbook becomes a schema, how a schema and data become a portable bundle, and how that bundle is rendered to Excel and PDF without ever holding the whole dataset in memory. This section is for understanding the machinery: making design decisions, debugging the unexpected, or extending the library with confidence.

Use this section to understand how the package works end to end internally. For task-focused walkthroughs on building reports with these internals, use the [Developer Guide](../developer_guide/index.md).

If you are new, read from top to bottom. If you are looking for a specific subsystem, jump straight to it from the menu below.

{{ ARCHITECTURE_MENU }}
