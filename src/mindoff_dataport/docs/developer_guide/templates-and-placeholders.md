# Templates & Placeholders

A template is just a normal Excel workbook with a few cells marked up to say "data goes here." You design the layout, fonts, borders, and colors exactly as you want them to appear, then drop in placeholders where the live values belong. The library reads those markers and builds the input contract from them.

The marker syntax is always `{{key:type}}`. The `key` is the name you'll use in your payload; the `type` tells the library what kind of value to expect.

```text
{{report_title:string}}
{{invoice_number:number}}
{{generated_on:date}}
{{line_items:dataframe}}
{{line_items:dataframe-header}}
{{line_items:dataframe-content}}
{{reports:repeat-start}}
  ...
{{reports:repeat-end}}
```

## Prerequisites

- `mindoff-dataport` installed (see [Installation](installation.md))
- An `.xlsx` workbook open in Excel or any spreadsheet editor where you can type cell values

## Implementation

### 1. Scalar Placeholders

Scalars are single values: a title, a number, a date. The placeholder cell is replaced **in place** with your value and keeps every style from the template: font, fill, border, alignment. What you styled is what you get.

| Type      | Accepted Python values                |
|-----------|---------------------------------------|
| `string`  | `str`                                 |
| `number`  | `int`, `float`                        |
| `int`     | `int`                                 |
| `float`   | `float`                               |
| `date`    | `datetime.date`, `datetime.datetime`  |
| `boolean` | `bool`                                |

### 2. Dataframe Placeholders

Dataframe placeholders are where a table lands. You anchor one cell, and the library writes the headers and rows outward from there, applying the anchor's style to every generated cell. Column names become the header text.

There are three flavors, depending on how much control you want over header styling:

| Type                 | What it writes                                            | Typical use                                       |
|----------------------|----------------------------------------------------------|---------------------------------------------------|
| `dataframe`          | Headers on the anchor row, content starting the next row | The all-in-one drop-in for a complete table       |
| `dataframe-header`   | Column headers only, on the anchor row                   | A header row styled separately from the content   |
| `dataframe-content`  | Data rows only, starting at the anchor row               | A content area sitting below a custom header row  |

Splitting `dataframe-header` and `dataframe-content` lets you style the header band one way (say, bold white-on-green) and the body another (zebra striping, right-aligned numbers). See [Dataframe Layout & Shifting](dataframe-layout.md) for the full layout options.

<div class="admonition tip">
<p class="admonition-title">Streaming note</p>
<p><code>dataframe-content</code> placeholders can stream straight from Parquet, which is what keeps memory flat on huge exports. In streaming mode, a non-repeat sheet may hold only <strong>one</strong> <code>dataframe-content</code> placeholder. More on that in <a href="exporting-xlsx.md">Exporting to Excel</a>.</p>
</div>

### 3. Repeat Placeholders

A repeat block is a section of your template that gets rendered once per record in a list. Picture an invoice layout (customer header, a line-items table, a footer) that you want to stamp out once per customer, stacked down the sheet. Mark the first and last rows of that section:

| Type           | Description                                                            |
|----------------|------------------------------------------------------------------------|
| `repeat-start` | Marks the first row of the repeating block (a control row, not rendered) |
| `repeat-end`   | Marks the last row of the repeating block (a control row, not rendered)  |

The control rows themselves never appear in the output; they only bracket the region. Everything between them repeats. See the [Repeat Sections recipe](recipes.md#2-repeat-sections-per-customer-blocks) for a full example and the rules that apply.

### 4. Dynamic Sheets

Sometimes you don't want a repeating block; you want a whole **sheet** per group (one tab per region, per customer, per month). Name the template sheet exactly `{{key}}` and it becomes a stencil: pass a dict of `output_sheet_name -> payload` and the library produces one sheet per entry, in insertion order. Details and payload shape live in [The Data Contract](data-contract.md#3-dynamic-sheet-group).

### 5. Manual Page Breaks

Templates can also carry Excel's own manual print breaks, the ones you set with **Page Layout → Breaks**. These are not placeholder syntax; they're real Excel metadata, and the library reads them too.

- `row_page_breaks`: 1-based template rows after which a new printed page begins.
- `column_page_breaks`: 1-based template columns after which a new printed page begins.

During `compile()` these breaks are re-resolved against the final layout (after tables expand and content shifts). XLSX preserves both row and column breaks; PDF honors row breaks as hard page boundaries and ignores column breaks. The full behavior is covered in [Dataframe Layout & Shifting](dataframe-layout.md#3-manual-page-breaks).

## Core Concepts

### 1. How the Library Discovers Placeholders

Extraction reads every cell's value looking for `{{key:type}}` markers. The resulting schema is a static snapshot of the template at that moment. If you edit the template file afterward — adding a placeholder, renaming a key, changing a type — re-run `extract()` before compiling; the schema won't update itself. An unrecognized type string (for example `{{x:strring}}`) is treated as plain text and won't appear in `inputs()`.

### 2. Key, Type, and Column Name

These three things are distinct and easy to conflate. The `key` is the name you use in your payload dict. The `type` constrains what value the library accepts and validates at compile time. For dataframe placeholders, the `key` identifies which dataframe to use; the column **names** of that `DataFrame` or `LazyFrame` become the header text written to the output. Renaming a dataframe column changes the header in the report, not the payload key.

## Troubleshooting

- **A placeholder came through as literal text in the output.**
  Check the syntax; it must be exactly `{{key:type}}` with a supported type. A typo in the type (`{{x:strring}}`) is treated as plain text.
- **`inputs()` doesn't list a placeholder I added.**
  Re-run `extract()` after editing the template. The schema is a snapshot taken at extraction time.
