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

1. Open a terminal (Command Prompt or PowerShell on Windows).
2. Navigate to this folder:
   ```
   cd path\to\cyclic-dependency-finder
   ```
3. Install the required library:
   ```
   pip install -r requirements.txt
   ```

---

## How to Run

Open a terminal, navigate to this folder, and run:

```
python find_cycles.py myfile.xlsx
```

Or run without any arguments — it will ask you for the file path:

```
python find_cycles.py
```

**That's it. The script will guide you through everything step by step:**

```
============================================================
  Cyclic Dependency Finder
============================================================
Enter the path to your Excel file (.xlsx): myfile.xlsx

  Loaded: myfile.xlsx
  Sheets: ProjectTasks, Modules, NoCycles

============================================================
  Columns found in your Excel file:
============================================================
  1. Task
  2. Depends On

Which column contains the ITEM NAME (e.g. Task, Module)?
  1. Task
  2. Depends On
Enter number (1-2): 1

Which column contains the DEPENDENCY (e.g. Depends On, Requires)?
  1. Depends On
Enter number (1-1): 1

If one cell can list multiple dependencies, how are they separated?
  1. Comma        (e.g. Task A, Task B)
  2. Semicolon    (e.g. Task A; Task B)
  3. Pipe         (e.g. Task A | Task B)
  4. New line     (e.g. each on its own line inside the cell)
  5. Only one dependency per cell (no separator needed)
Enter number (1-5): 1
```

You just type a number and press Enter at each step. No flags or technical commands needed.

---

## How Your Excel File Should Look

The tool works with `.xlsx` files. Your Excel file needs at least two columns:

1. One column for the **item name** (e.g. Task, Module, Component).
2. One column for **what it depends on**.

### Example

| Task    | Depends On |
|---------|------------|
| Task A  | Task B     |
| Task B  | Task C     |
| Task C  | Task A     |

The tool will report: `Task A -> Task B -> Task C -> Task A` (cycle!).

### Multiple dependencies in one cell

If one item depends on multiple things, you can list them in a single cell separated by commas, semicolons, etc. The script will ask you which separator you use.

| Task    | Depends On       |
|---------|------------------|
| Task A  | Task B, Task C   |

### Multiple sheets

The tool reads **all sheets** automatically. Just make sure the column names are the same across sheets.

---

## Understanding the Output

### No cycles (good)

```
=== RESULT: No cyclic dependencies found! ===
```

### Cycles found (needs fixing)

```
=== RESULT: Found 2 cyclic dependency(ies)! ===

  Cycle 1:  Task A -> Task B -> Task C -> Task A
            Task A -> Task B  (found in sheet: ProjectTasks)
            Task B -> Task C  (found in sheet: ProjectTasks)
            Task C -> Task A  (found in sheet: ProjectTasks)

  Cycle 2:  Module X -> Module Y -> Module Z -> Module X
            Module X -> Module Y  (found in sheet: Modules)
            Module Y -> Module Z  (found in sheet: Modules)
            Module Z -> Module X  (found in sheet: Modules)
```

It tells you:
- The full cycle chain.
- **Which sheet** each link was found in, so you know exactly where to fix it.

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| `python` command not found | Try `python3` instead of `python`, or install Python. |
| `No data found` error | The column names you selected may not exist in every sheet — that's OK, it only needs to be in at least one. |
| File not found error | Make sure the Excel file path is correct. Use quotes around the path if it has spaces. |
| `.xls` file not supported | Save the file as `.xlsx` in Excel (File > Save As > .xlsx). |
