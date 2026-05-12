# VBA Excel Macros

A library of Excel VBA macros for common data cleaning, formatting, and spreadsheet automation tasks.

## Macros Included

| Macro | Description |
|---|---|
| `UpperCase` | Converts selected cell values to uppercase |
| `CONVERT_2_TEXT` | Converts cell contents to text format |
| `CONVERT_REDACT` | Redacts sensitive values in a selection |
| `CONVERT_REDACT_HAI_UNITS` | Redaction variant for HAI unit data |
| `ODD_ROW_HIGHLIGHT` | Applies alternating row highlighting for readability |
| `EIN_DLT_EMPTY_IN_RNG` | Deletes empty cells within a range and shifts up |
| `EIN_DLT_EMPTY_IN_RNG2` | Variant with alternate shift behavior |
| `DARKLINEBORDER` | Applies dark border formatting to a selection |
| `unhide_me` | Unhides all hidden rows and columns in the active sheet |
| `DERESTRICT_FORMAT` | Removes formatting restrictions from locked cells |
| `PasswordBreaker` | Removes worksheet protection passwords |
| `DELETE_ROW_IF_EMPTY` | Deletes entire rows where a target column is blank |
| `MOVE_ALL_2_SINGLE_COL` | Consolidates multi-column data into a single column |

## Usage

1. Open Excel and press `Alt + F11` to open the VBA editor
2. Import the macro file via **File → Import File**
3. Run macros via `Alt + F8` or assign to buttons/shortcuts

## Tech
- Microsoft Excel VBA (compatible with Excel 2016+)
