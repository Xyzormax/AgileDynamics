# excel_coda — Design Notes

## Goal

A set of VBA macros that give Excel authors a familiar, Coda-like API for
manipulating tabular data: filtering, adding, modifying, and deleting rows,
and resolving relationships between tables — without VLOOKUP.

## Approach

Pure procedural VBA. Macros have side effects and should never be used as
cell formulas (doing so defeats Excel's dependency graph). Logic is written
as VBA Sub/Function calls, not worksheet formulas.

---

## Data Model

Each table is an **Excel ListObject** (Insert → Table) on its own sheet.
ListObjects give named column references, auto-expansion on row add, and a
stable VBA API.

---

## Object Model

### Entry point

```vba
GetTable("Tasks")   ' returns a DBTable
```

### DBTable

| Method | Signature | Returns | Notes |
|---|---|---|---|
| `Filter` | `(col, value)` | `DBTable` | Eager: scans immediately, stores matching row indices. Chainable — multiple Filter calls AND together. |
| `AddRow` | `(col1, val1, col2, val2, ...)` | — | Paired args. Always appends; not affected by filter context. |
| `ModifyRows` | `(col1, val1, col2, val2, ...)` | — | Paired args. Operates on filtered row set. |
| `DeleteRows` | — | — | Operates on filtered row set. |
| `Values` | `(col)` | `Variant()` | Returns column values for filtered rows. |
| `Rows` | — | `Collection` of `DBRow` | Iterate filtered rows. |
| `First` | — | `DBRow` | First row in filtered set. |

### DBRow

| Method | Signature | Returns | Notes |
|---|---|---|---|
| `Column` | `(name)` | `Variant` | Read a cell value. |
| `SetColumn` | `(name, value)` | — | Write a cell value. |
| `Lookup` | `(targetTable, joinCol, returnCol)` | `Variant` | Explicit join — no metadata sheet. |

---

## Key Design Decisions

### Filter is eager (Option A)
`Filter()` scans the underlying ListObject immediately and stores the
matching row indices. Subsequent chained `Filter()` calls narrow that set.
Terminal methods (`DeleteRows`, `ModifyRows`, `Values`, `Rows`, `First`)
operate on the stored index set.

### All mutations require a prior Filter
`ModifyRows`, `DeleteRows`, and `SetColumn` always act on the filtered
context. To affect all rows, call no Filter before them.

### Relationships are explicit per call
No metadata sheet or schema declaration for relationships. The caller
passes join parameters directly to `Lookup`:
```vba
row.Lookup("Projects", "ProjectID", "Name")
' Find row in Projects where Projects.ID = this row's ProjectID, return Name
```

### Column names are strings
VBA cannot use bare identifiers as dynamic property names. Column names are
always passed as string arguments.

---

## Usage Examples

```vba
' Multi-condition filter → delete
GetTable("Tasks").Filter("Status", "Cancelled").Filter("Priority", 0).DeleteRows

' Filter → modify multiple columns
GetTable("Tasks").Filter("Status", "Pending").ModifyRows "Status", "Done", "Resolved", Now()

' Add a row
GetTable("Tasks").AddRow "Name", "New Task", "Status", "Pending", "Priority", 1

' Iterate and resolve a relationship
Dim row As DBRow
For Each row In GetTable("Tasks").Filter("Status", "Done").Rows
    Debug.Print row.Lookup("Projects", "ProjectID", "Name")
Next
```

---

## Open Question

Currently every statement manipulates the sheet directly and immediately:
- No atomicity — partial failure leaves the sheet in an inconsistent state
- Each write triggers Excel recalculation/redraw (mitigated by disabling
  ScreenUpdating and Calculation at macro start)
- VBA operations complicate Excel's undo stack

**Under consideration:** a staged/in-memory execution model where changes
are accumulated and flushed to the sheet in a single pass at the end.
Decision deferred.
