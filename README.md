# excel2json

A minimal, zero-configuration Excel → Python dict / JSON converter.

Author configurations in a spreadsheet, load them in code.  
No schema files. No YAML. No boilerplate.

```python
from excel2json import excel2dict

data = excel2dict("config.xlsx")
print(data["Game"]["rtp"])   # 0.965
```

---

## Installation

```bash
pip install openpyxl      # only runtime dependency
```

Copy `excel2json.py` into your project (single-file, no package needed).

---

## Concepts

### Two sheet formats

| Sheet name | Format | Returns |
|---|---|---|
| Starts with `T_` | Table | `list` or `dict` (when an `ID` column exists) |
| Anything else | Key-value | `dict` |
| Starts with `#` | — | Skipped |

### Key-value sheet

Each row is `KEY  TYPE  VALUE(S)`:

```
A              B        C          D    E    F
machine_name   STR      AwesomeSlot
reel_count     INT      5
rtp            REAL     0.965
enabled        BOOL     1
tags           [STR]    wild       scatter  bonus
```

Result:

```json
{
  "machine_name": "AwesomeSlot",
  "reel_count": 5,
  "rtp": 0.965,
  "enabled": true,
  "tags": ["wild", "scatter", "bonus"]
}
```

### Table sheet (`T_` prefix)

Sheet name: `T_Items`

```
Name     Multiple   IsWild
STR      INT        BOOL
Wild     0          1
Scatter  0          0
Cherry   3          0
```

Result key: `"Items"` (prefix stripped)

```json
[
  {"Name": "Wild",    "Multiple": 0, "IsWild": true},
  {"Name": "Scatter", "Multiple": 0, "IsWild": false},
  {"Name": "Cherry",  "Multiple": 3, "IsWild": false}
]
```

Add an `ID` column to get a dict instead of a list:

```
ID       Name     Multiple
STR      STR      INT
WILD     Wild     0
SCAT     Scatter  0
```

Result:

```json
{
  "WILD": {"Name": "Wild",    "Multiple": 0},
  "SCAT": {"Name": "Scatter", "Multiple": 0}
}
```

---

## Type reference

| TYPE | Python type | Notes |
|---|---|---|
| `INT` | `int` | `None` / empty → `0` |
| `REAL` | `float` | `None` / empty → `0.0` |
| `STR` | `str` | `None` / empty → `""` |
| `BOOL` | `bool` | `0`/`FALSE` → `False`, `1`/`TRUE` → `True` |
| `DICT` | `dict` | Value cell must contain valid JSON |
| `[INT]` `[REAL]` `[STR]` `[BOOL]` | `list` | Values span cols C, D, E … on the **same row** |
| `[]` | `list` | Sub-region table (rows below, cols C+) |
| `{}` | `list` or `dict` | Sub-region key-value table (rows below, cols C+) |
| `{T}` | `dict` | Sub-region transposed key-value |
| `{DICT}` | `dict` | Sub-region flat dict (see below) |

---

## Sub-region types

For `[]`, `{}`, `{T}`, and `{DICT}` the data lives in the rows **immediately below** the KEY row.

**Rule:** continuation rows must have **empty** cells in columns A and B.  
The parser treats "non-empty col A" as the start of the next top-level key.  
Sub-region data starts at **column C**.

### `[]` — sub-region table

```
A          B    C       D         E
inventory  []   Name    Qty       Price
                STR     INT       REAL
                Sword   10        9.99
                Shield  5         14.50
next_key   INT  …
```

```json
{"inventory": [{"Name": "Sword", "Qty": 10, "Price": 9.99}, …]}
```

### `{DICT}` — flat dict

First entry goes on the **same row** as the key; subsequent entries go on continuation rows with col C = sub-key, col D = sub-type, col E+ = sub-values.

```
A       B        C       D      E    F    G
lines   {DICT}   Line1   [INT]  2    2    2
                 Line2   [INT]  3    3    3
                 Line3   [INT]  1    2    1
reels   INT      5
```

```json
{
  "lines": {
    "Line1": [2, 2, 2],
    "Line2": [3, 3, 3],
    "Line3": [1, 2, 1]
  },
  "reels": 5
}
```

### `{T}` — transposed key-value

Each **column** (C, D, E …) represents one key-value pair.  
Row 1 (same as KEY row) = field names, row 2 = types, row 3+ = values.

```
A         B    C      D
settings  {T}  width  height
               INT    INT
               1920   1080
```

```json
{"settings": {"width": 1920, "height": 1080}}
```

---

## Nameless key — inline merge

If a key is named exactly `Nameless`, its parsed value (must be a `dict`) is merged directly into the parent dict. The key `Nameless` never appears in the output.

Useful for grouping rows visually in the spreadsheet without adding nesting in JSON:

```
A          B        C         D      E
before     INT      10
Nameless   {DICT}   child_a   INT    100
                    child_b   STR    hello
after      INT      20
```

Result:

```json
{"before": 10, "child_a": 100, "child_b": "hello", "after": 20}
```

An **entire sheet** can consist of `Nameless` blocks — the sheet result is a flat dict:

```
A          B        C         D      E
Nameless   {DICT}   name      STR    example
                    version   INT    3
Nameless   {DICT}   min_val   INT    0
                    max_val   INT    100
```

Result:

```json
{"name": "example", "version": 3, "min_val": 0, "max_val": 100}
```

> **Note:** `Nameless` only works with dict-producing types (`{}`, `{T}`, `{DICT}`).  
> Using it with a list type (`[]`, `[INT]`, …) raises `TypeError`.

---

## API

### `excel2dict(path)`

```python
def excel2dict(path: str | Path) -> dict
```

Parse an entire workbook.  Returns `{sheet_name: parsed_data}`.

```python
data = excel2dict("config.xlsx")
data["Game"]["rtp"]           # float
data["Items"][0]["Name"]      # str  (list sheet)
data["Items"]["WILD"]["Name"] # str  (dict sheet, ID column)
```

### `excel2json(path, output=None, indent=2)`

```python
def excel2json(path: str | Path,
               output: str | Path | None = None,
               indent: int = 2) -> str
```

Parse and serialise to JSON.

```python
# Print to stdout
print(excel2json("config.xlsx"))

# Write to file
excel2json("config.xlsx", output="config.json")

# Compact (minified) output
excel2json("config.xlsx", output="config.min.json", indent=None)
```

---

## CLI

```bash
# Print JSON to stdout
python excel2json.py config.xlsx

# Write JSON to file
python excel2json.py config.xlsx config.json
```

---

## Error messages

All parse errors include a breadcrumb showing the sheet name, key, and row/column, making it easy to locate the problem:

```
ValueError: [Sheet('Game')[key='rtp',R5]] Cannot convert to REAL: 'N/A'
```

---

## Limitations

- Reads `.xlsx` only (uses `openpyxl` with `data_only=True`; formulas must be pre-evaluated or saved with cached values).
- `datetime` / `time` cell values are coerced to strings.
- Merged cells are not supported; unmerge before parsing.

---

## License

MIT
