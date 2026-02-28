# Cyclic Dependency Finder

A simple tool that reads an Excel file and checks if there are any **circular (cyclic) dependencies**.

**What is a cyclic dependency?**
If Task A depends on Task B, Task B depends on Task C, and Task C depends on Task A — that is a cycle. No task can start because each one is waiting for another.

---

## Prerequisites

- **Python 3.8 or newer** must be installed on your computer.
  - To check, open a terminal and type: `python --version`
  - If not installed, download from https://www.python.org/downloads/

---

## Setup (one time only)

Open a terminal (Command Prompt or PowerShell on Windows) and run:

```
pip install -r requirements.txt
```

This installs the library needed to read Excel files.

---

## How Your Excel File Should Look

The tool works with `.xlsx` files. Your Excel file needs **at least two columns**:

1. **Source column** — the name of the item (e.g., Task, Module, Component).
2. **Target column** — what it depends on.

### Example

| Task    | Depends On |
|---------|------------|
| Task A  | Task B     |
| Task B  | Task C     |
| Task C  | Task A     |

In this example, the tool will report a cycle: `Task A -> Task B -> Task C -> Task A`.

### Multiple dependencies in one cell

If one item depends on multiple things, separate them with commas:

| Task    | Depends On       |
|---------|------------------|
| Task A  | Task B, Task C   |
| Task B  | Task D           |

### Multiple sheets

The tool reads **all sheets** in the Excel file automatically. Just make sure the column names are the same across sheets.

---

## How to Run

Open a terminal, navigate to this folder, and run:

```
python find_cycles.py <your-file.xlsx> --source "<source column name>" --target "<target column name>"
```

### Real examples

**Example 1** — columns named "Task" and "Depends On":

```
python find_cycles.py my_tasks.xlsx --source "Task" --target "Depends On"
```

**Example 2** — columns named "Module" and "Requires":

```
python find_cycles.py modules.xlsx --source "Module" --target "Requires"
```

**Example 3** — dependencies separated by semicolon instead of comma:

```
python find_cycles.py data.xlsx --source "Component" --target "Needs" --separator ";"
```

---

## Understanding the Output

### No cycles found (good)

```
=== RESULT: No cyclic dependencies found! ===
```

### Cycles found (problem)

```
=== RESULT: Found 1 cyclic dependency(ies)! ===

  Cycle 1:  Task A -> Task B -> Task C -> Task A
            Task A -> Task B  (found in sheet: Sheet1)
            Task B -> Task C  (found in sheet: Sheet1)
            Task C -> Task A  (found in sheet: Sheet2)
```

The output tells you:
- The full cycle chain.
- Which sheet each dependency was found in, so you can go fix it.

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| `python` command not found | Try `python3` instead of `python`, or install Python. |
| `No data found` error | Double-check that the column names you typed match the Excel headers exactly. |
| File not found error | Make sure the Excel file path is correct. Use quotes if the path has spaces. |
| `.xls` file not supported | Save the file as `.xlsx` in Excel (File > Save As > .xlsx). |
