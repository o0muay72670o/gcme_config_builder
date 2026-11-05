# GCME Config Builder

A Python and Codex chat workflow for building the **Config worksheet** used to map
engineering Excel form fields to data-entry targets. No separate AI API key is required.

This is a separate project from [GCME Formatter](https://github.com/o0muay72670o/gcme_formatter).
It creates field-binding metadata; formatting and rebuilding engineering forms belong
to the formatter project.

```mermaid
flowchart LR
    A[Excel form] --> B[Python prepare]
    B --> C[Local request + schema]
    C --> D[Codex chat: review field mappings]
    D --> E[config-plan.json]
    E --> F[Python validate and apply]
    F --> G[New workbook copy with Config]
    G --> H[Human review]
```

Python extracts cells and merged ranges, validates mappings and edits only Config
worksheet XML plus the package metadata needed to add a new sheet. Other worksheet
XML and embedded assets are retained. Codex interprets unfamiliar field layouts;
the workflow does not automatically start a Codex conversation.

## Config schema

| Column | Meaning |
|---|---|
| Label | Field name, composite sign-off label, or ordered table column names |
| Cells | Label cell address; blank for composite/table labels |
| Group Name | Header, DataTable, completion role or other section |
| Sequence | Unique positive integer controlling row order |
| Field Require | Y, N, or blank when unconfirmed |
| DB | Entry/value cell, or first data cell for a table |
| DataType | Text, Date, Dropdown, tabular, Image |
| Sheet | Exact DB target form or Condition sheet name; Cells may reference the selected form label when DB is on Condition |
| OptionValue | Confirmed dropdown choices |
| Remark | Mapping notes |
| DefautlValue | Confirmed default; spelling retained for compatibility |
| ValueExpression | Confirmed downstream expression such as `{ProjectName}` |
| ValueSource | Blank, 1 for constant, or 2 for expression, using the reference convention |
| ColumnList | Ordered Excel column letters for tabular fields |

For example, an approved-welders table can use one `tabular` mapping with labels
`IDENTITY NO., NAME, WPS_1, QUAL_1, WPS_2, QUAL_2` and ColumnList
`A, F, O, S, V, Z`. Actual addresses depend on the workbook. Blank metadata is
preferable to invented required flags, defaults or business rules.

## Setup and Visual Studio 2022

Use Python 3.11+ and install the Visual Studio **Python development** workload.

```powershell
git clone https://github.com/o0muay72670o/gcme_config_builder.git
cd gcme_config_builder
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Open `GCMEConfigBuilder.sln` in Visual Studio 2022. Select `.venv` in Python
Environments. Startup arguments default to `--help`; change **Project Properties →
Debug → Script Arguments** to the desired command. All project paths are relative.

## Use

Put your workbook in `input/`. Inspect its sheet names, then select the actual form
and any Condition sheet needed for dropdown targets.

```powershell
.\.venv\Scripts\python.exe main.py prepare "input/form.xlsx" --sheets "Form" "Condition"
```

Ask Codex:

> Read AGENTS.md and the prepared config-request.json. Create config-plan.json
> mapping actual labels, entry cells, tables, dropdowns and sign-off roles. Preserve
> the 14-column schema and source_hash. Leave unconfirmed business metadata blank
> and flag ambiguities. Validate and apply locally without an API.

Then apply the reviewed plan:

```powershell
.\.venv\Scripts\python.exe main.py apply "jobs/<job>/config-request.json" --output output
```

The output copy has a generated/replaced Config sheet and a JSON report. Original
inputs and existing output files are never overwritten. Prepare does not overwrite
an existing hand-authored plan. No formatter `config.json` or Master file is required.

## Validation and limits

- Rejects changed-workbook fingerprints, invalid/duplicate targets, merged children,
  label/entry overlap, inconsistent table columns, missing dropdown options and
  formula-like metadata. It does not execute downstream expressions.
- Refuses Config sheets with formula dependencies, since replacing them could affect
  engineering calculations. Review those workbooks separately.
- Mapping completeness and business meaning still require review; address validation
  does not prove a field was understood correctly. Defaults and required flags are
  not inferred from sample record values.
- Existing source types may be mistaken: a company field marked Image or signature
  marked Text should be reviewed, rather than blindly reproduced.
- Native Excel UI/print behavior is not verified by the CLI. Config is metadata and
  does not change the form's print layout. Workbook formula/answer cells are not edited.
- XLSX input only. No engineering files, credentials or generated jobs are published.

## Tests

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Tests use synthetic workbooks and do not require network access or an AI API.
