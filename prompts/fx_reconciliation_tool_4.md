Now we've implemented fx_reconciliation_core.py and fx_consolidation_postprocess.py, and here the workflow:
- execute fx_reconciliation_core.py against a folder, which import the related files into the sheet, and find the missing data.
- user manually fill these missing data by looking up in the production system.
- execute fx_consolidation_postprocess to generate the final report. 

Context
- fx_summary_workflow_app.py and run_fx_summary_workflow_windows.cmd were created initially before finalize_fx_summary_report.py for non-tech people to use the tool from Windows.

New Requirement
- I need better file names to replace fx_reconciliation_core.py and fx_consolidation_postprocess.py. Please give me options for me to choose.
  - The target file name is `各通道需换汇情况汇总` in Chinese. You can choose the proper translation as a reference for the naming.
- fx_reconciliation_core.py generates the output file in result folder now. Now `-wip` is appended in the file name. There aare two options as follows. Please let me know which one is better and why.
  - remove the `-wip`
  - keep the `-wip` in the output file, and remove it after executing fx_consolidation_postprocess.py
- Update fx_summary_workflow_app.py to support the workflow mentioned above
  1. The existing function: choose a folder to run fx_reconciliation_core.py. After the execution, provide info for user to know what sheets need to fill data. 
     If no data needs to fill, skip the step 2. 
  2. Provide checkboxes for user to confirm that she or he has filled the necessary data. The checkbox label should be the sheet name. 
  3. Only execute fx_consolidation_postprocess.py if no manually data filling is needed or user has checked all the checkboxes. 
- Change the Windows title to Chinese `生成各通道需换汇情况汇总`



## FX Reconciliation Workflow Enhancement

### Current Workflow

The system currently consists of two main scripts:

- `fx_reconciliation_core.py`
  - Imports source files into a workbook
  - Identifies missing data that requires manual completion

- `fx_consolidation_postprocess.py`
  - Processes the completed workbook
  - Generates the final FX summary report

### Execution Flow

1. Run `fx_reconciliation_core.py` on a selected folder
2. User manually fills missing data by referencing the production system
3. Run `fx_consolidation_postprocess.py` to generate the final report

---

### Context

- `fx_summary_workflow_app.py` and `run_fx_summary_workflow_windows.cmd` were created earlier
- Purpose: enable **non-technical users** to run the workflow on Windows

---

## New Requirements

### 1. Improve Script Naming

Provide better file name options to replace:
- `fx_reconciliation_core.py`
- `fx_consolidation_postprocess.py`

#### Requirements:
- Names should be:
  - Clear and user-friendly (especially for non-technical users)
  - Consistent with the final report meaning

- Reference (Chinese target name):
  - `各通道需换汇情况汇总`

- Provide **multiple naming options** for selection

---

### 2. Output File Naming Strategy

Currently:
- `fx_reconciliation_core.py` generates output files in the `result` folder
- File name includes suffix: `-wip`

#### Evaluate the following options:

- **Option A:** Remove `-wip` entirely
- **Option B:** Keep `-wip` initially, then remove it after running `fx_consolidation_postprocess.py`

#### Requirement:
- Recommend the better option
- Provide reasoning (consider usability, clarity, and workflow safety)

---

### 3. Update Windows UI Workflow (`fx_summary_workflow_app.py`)

Enhance the UI to support the full workflow:

#### Step 1 — Initial Execution
- Allow user to select a folder
- Run `fx_reconciliation_core.py`
- After execution:
  - Display which sheets require manual data input

- If **no manual input is required**:
  - Skip Step 2

---

#### Step 2 — Manual Confirmation
- Display a list of sheets requiring input
- Provide **checkboxes** for user confirmation:
  - Each checkbox label = sheet name

---

#### Step 3 — Final Processing
- Only allow execution of `fx_consolidation_postprocess.py` when:
  - No manual input is required, OR
  - User has checked all required checkboxes

---

### 4. Windows Application Title

- Update the window title to: `生成各通道需换汇情况汇总`



---

## Expected Outcome

- Improved usability for non-technical users
- Clearer script naming aligned with business meaning
- Safer and more guided workflow execution
- Reduced risk of incomplete data processing
