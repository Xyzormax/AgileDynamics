# excel_cfl — Design Notes

## Goal

A VBA framework that gives Excel authors the closest possible experience to
writing Coda Formula Language (CFL) actions. The author writes procedural VBA
against an in-memory object graph; the Excel sheet is the view and persistence
layer, not the database.

The target author is a Coda maker who understands CFL but has limited Excel
knowledge. Differences from CFL should be minimal and understandable.

---

## Architecture: MVC

| Layer | Role |
|---|---|
| **Model** | In-memory object graph — `DBTable` and `DBRow` instances |
| **View** | Excel sheets — ListObjects rendered from the model |
| **Controller** | Button macros — VBA subs written by the author |

The sheet is only read on load and written on save. All logic runs against RAM.

---

## Data Model

Each table is an **Excel ListObject** (Insert → Table) on its own sheet.
Tables are loaded into a graph of `DBTable` / `DBRow` objects at the start of
each button press and flushed back at the end.

### Table variables

Each table gets a `Public` module-level variable in a shared `DBTables`
module, declared once by the author when a new table is created:

```vba
' DBTables module — author adds one line per table
Public Tasks As DBTable
Public Projects As DBTable
Public Log As DBTable
```

The framework populates these variables during `DB.Load`. This gives the
author bare-name table access (`Tasks.Filter(...)`) matching the CFL feel.

---

## Object Model

### DBTable

| Method | Signature | Returns | Notes |
|---|---|---|---|
| `Filter` | `(col, value)` | `DBTable` | Eager: scans row set immediately, stores matching indices. Chainable — multiple calls AND together. |
| `AddRow` | `(col1, val1, col2, val2, ...)` | — | Paired args. Always appends; ignores filter context. |
| `ModifyRows` | `(col1, val1, col2, val2, ...)` | — | Paired args. Operates on filtered row set. |
| `DeleteRows` | — | — | Operates on filtered row set. |
| `Values` | `(col)` | `Variant()` | Column values for filtered rows. |
| `Rows` | — | `Collection` of `DBRow` | Iterate filtered rows. |
| `First` | — | `DBRow` | First row in filtered set. |

### DBRow

| Method | Signature | Returns | Notes |
|---|---|---|---|
| `Column` | `(name)` | `Variant` | Read a cell value. |
| `SetColumn` | `(name, value)` | — | Write a cell value. |
| `Lookup` | `(targetTable, joinCol, returnCol)` | `Variant` | Explicit join — no schema declaration needed. |

### thisRow

When a button macro fires from within a table row, the framework auto-injects
`thisRow` as a module-level `DBRow` pointing to the row containing the button.
This replicates CFL's implicit row context.

```vba
Sub OnApproveClick()
    thisRow.SetColumn "Status", "Approved"
    thisRow.SetColumn "ResolvedAt", Now()

    Dim projectName As String
    projectName = thisRow.Lookup("Projects", "ProjectID", "Name")

    Log.AddRow "Event", "Approved", "Project", projectName
End Sub
```

---

## Load / Save Cycle

Every button press runs through a dispatcher that owns the full cycle:

```vba
Sub RunAction(macroName As String)
    DB.Load
    On Error GoTo Rollback
    Application.Run macroName
    DB.Save
    Exit Sub
Rollback:
    MsgBox "Action failed: " & Err.Description & " — no changes saved."
End Sub
```

### Load

- Triggered **per button press** (not on workbook open)
- Reads all ListObjects into the in-memory object graph
- Populates all `Public DBTable` variables in `DBTables`
- Detects the calling button's sheet position and sets `thisRow`
- Guarantees a fresh model even if cells were edited directly between presses

### Save

- **Full flush** — all tables are written back to their ListObjects on every
  save, regardless of which tables were touched
- Simple to implement, no risk of missed changes, negligible cost at
  hundreds-of-rows scale

### On error

- Save is **skipped** — the sheet is left unchanged
- Error message shown to the user
- No partial state is committed

---

## Button Mechanism

Buttons are **Form Controls** (not ActiveX). Each button has a macro assigned
that calls `RunAction` with the author's sub name:

```vba
' Assigned to the Form Control button:
Sub ApproveButton_Click()
    RunAction "OnApproveClick"
End Sub
```

`thisRow` is resolved via `Application.Caller`, which returns the shape name
of the Form Control. The framework reads the shape's position on the sheet to
determine which ListObject row it occupies.

---

## Relationships

No schema or metadata sheet. Relationships are expressed explicitly per call
via `Lookup` on a `DBRow`:

```vba
row.Lookup("Projects", "ProjectID", "Name")
' Find the row in Projects where Projects.ID = this row's ProjectID
' Return that row's Name column
```

---

## CFL Comparison

| Coda CFL | excel_cfl VBA |
|---|---|
| `Tasks.Filter(Status = "Done")` | `Tasks.Filter("Status", "Done")` |
| `thisRow.Status = "Done"` | `thisRow.SetColumn "Status", "Done"` |
| `thisRow.Status` | `thisRow.Column("Status")` |
| `Tasks.Filter(...).Name` | `Tasks.Filter(...).Values("Name")` |
| Implicit row context | `thisRow` auto-injected by dispatcher |
| Table as bare name | `Public Tasks As DBTable` declared once |

The main differences are string column names (VBA has no dynamic bare
properties) and explicit `SetColumn` / `Column` accessors.
