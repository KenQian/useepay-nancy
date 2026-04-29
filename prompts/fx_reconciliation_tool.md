# Task: Foreign Exchange (FX) Channel Settlement Automation

## 1. Objective
Develop a Python script to automate the daily Foreign Exchange (FX) reconciliation process. The script must process multiple Excel source files, merge data based on lookup keys, apply **live financial formulas**, and handle "In-Progress" (WIP) file states.

## 2. File Environment & Initialization
* **Root Directory:** Accept a CLI argument for the folder path containing the source files.
* **The Baseline File:** Identify `各通道需换汇情况汇总-<date>.xlsx` (where `<date>` is the most recent `YYYY-MM-DD` found in the filename).
* **Workspace Creation:**
    * Copy the Baseline File.
    * Name the copy: `各通道需换汇情况汇总-<new_date>-wip.xlsx`.
    * **Final Step:** Once all processing is finished, rename by removing the `-wip` suffix.
* **Required Source Files (Validation):**
    * The script must confirm the presence of these files in the folder before starting:
        1. `各通道需换汇情况汇总-<date>.xlsx`
        2. `渠道订单.xls`
        3. `消费-商户本金(清算后).xls`
        4. `退款-商户本金(清算后).xls`

## 3. Module A: `账户流水` (Account Statement) Processing
**Target:** `账户流水` sheet in the WIP file.
* **Clear Old Data:** Delete all content and only keep the header (Row 1).
* **Data Ingestion (Stacked Import):**
    1. Load `消费-商户本金(清算后).xls` (Consumption File) -> sheet `交易详情`. Copy this data into the target sheet starting at **Row 2 (Columns A-Q)**.
    2. Load `退款-商户本金(清算后).xls` (Refund File) -> sheet `交易详情`. **Append** this data to columns **A-Q** immediately below the last row of data from the Consumption File. (Ensure headers from the Refund File are not duplicated).
* **Calculated Columns (Live Formulas):** Write the following **live Excel formulas** into the cells:
    * **Col M (余额):** `=K[row] - L[row]`
    * **Col R (Key):** `=B[row] & E[row]`
    * **Col S:** Copy value from **Col G**.
* **Requirement:** Do NOT convert formulas to static values; they must remain live for verification.

## 4. Module B: `渠道订单` (Channel Orders) Processing
**Target:** `渠道订单` sheet in the WIP file.
* **Source 1 (Daily Orders):** Load `渠道订单.xls` -> sheet `交易详情`.
    * **Filter:** Skip rows where Column J (`交易类型`) is "预授权申请" or "预授权撤销".
    * **Mapping:** Source A-D $\rightarrow$ Target A-D; Source F-AH $\rightarrow$ Target E-AG; Source E $\rightarrow$ Target AH.
* **Source 2 (Pending/Special Orders - Re-validation Logic):** 
    1. Load the **Baseline File** -> sheet `特殊的渠道订单`.
    2. **Internal Re-calculation (Critical):** Because Column AJ in this sheet contains a formula (`=XLOOKUP(AI,账户流水!R:R,账户流水!R:R)`) that depends on the data just updated in Module A, the script must **simulate this lookup in Python memory**. 
    3. Compare `特殊的渠道订单` Column AI against the **newly stacked** `账户流水` Column R. 
    4. **Selection:** If a match is found (meaning the row is now "valid" based on the updated statement), **append** that row (Columns A-AH) to the Target `渠道订单` sheet.

## 5. Module C: Financial Logic & Lookups
**Execution Order Requirement:** Due to nested dependencies (AL depends on AR, AR depends on AQ, etc.), the script must resolve lookups internally via Python in the following order before writing formula strings:
1. **Resolve AP & AQ:** Identify channel and map merchant IDs.
2. **Construct AR:** Build lookup key from AP, AQ, and AJ.
3. **Resolve AI:** Build account lookup key.
4. **Perform Lookups:** Find internal values for AL, AM, and AN to drive Module D logic.

 
## 6. Module D: Post-Processing & Exception Handling
For rows in `渠道订单` where the **internal Python lookup** for Column AM results in no match (which would manifest as `#N/A` in Excel):
* **Logic Source:** Identify these exceptions by checking if the internal "Shadow Calculation" merge for Column AM yielded a null/empty result.
* **Check Exception List:** If **Column AP** value is **NOT** "PayPal" or "Afterpay直连" (case-insensitive):
    * **Move to Special:** Append columns **A-AH** of these rows to the sheet `特殊的渠道订单`.
    * **Update Special Keys:**
        * In `特殊的渠道订单`, set AI = `D[row] & E[row]`.
        * In `特殊的渠道订单`, set AJ = `XLOOKUP(A[row], 账户流水!$R:$R, 账户流水!$R:$R)`.
    * **Cleanup:** Delete the matching rows from the `账户流水` sheet.

## 7. Technical Requirements
* **Shadow Calculation Rule:** The script must use `pandas` to perform all VLOOKUP/XLOOKUP logic internally to drive decision-making (Module B and D), as Excel formulas will not calculate until the file is opened by a user.
* **Formula Preservation:** **CRITICAL:** The script must write actual Excel formula strings (e.g., `ws.cell(row=r, column=c).value = "=A1+B1"`) using `openpyxl`.
* **Visibility:** Do not skip hidden, filtered, or grouped rows during ingestion.
* **Merchant IDs:** Treat all IDs and order numbers as **Strings** to prevent scientific notation conversion.
* **Optimization:** Use `pandas` for filtering/internal mapping and `openpyxl` for writing formulas.
* **Legacy Support:** Handle `.xls` using the `xlrd` engine within pandas.

---

# Revised Task: Multi-File Support & Dynamic Schema Update

## 1. File Discovery & Mapping (Regex Logic)
The script must now support multiple source files for each category. Implement a flexible discovery function (e.g., `get_source_files(directory, pattern)`) that returns a sorted list of paths based on the following mapping patterns:
* **Refunds (Type 1):** Matches `1.xls` or `1-*.xls` (e.g., `1-1.xls`, `1-2.xls`).
* **Consumption (Type 2):** Matches `2.xls` or `2-*.xls`.
* **Channel Orders (Type 3):** Matches `3.xls` or `3-*.xls`.
* **Baseline:** Remains `各通道需换汇情况汇总-*.xlsx` (Identify the most recent via date string).

## 2. Sequential Stacking & Data Ingestion (The "Append" Rule)
The script must iterate through the discovered files for each category and combine them into a single data pool for the target sheets.
* **Module A (Account Statement):** Sequentially append all data from every discovered **Type 1** file, immediately followed by all **Type 2** files, into the `账户流水` sheet. 
* **Module B (Channel Orders):** Append all discovered **Type 3** files into the `渠道订单` sheet before proceeding with filtering and column mapping.
* **Technical Requirement:** When iterating through files within a category, the script must calculate the `last_row` of the target dataframe or sheet to ensure each new batch of data begins immediately after the previous one without overwriting or leaving gaps.

## 3. Robust Header Detection & Duplicate Prevention
Since source files may or may not contain headers, implement a detection gate:
* **The Logic:** For every file loaded, check if Row 1 contains known keywords (e.g., for Type 3, check for "交易类型" or "订单号").
* **Action:** * If no header is detected, load the data using the script's internal schema.
    * If a header is detected, the script must **discard the header row for every file except the very first one in the stack** to prevent duplicate headers from appearing in the middle of the target sheet.
* *Note: Please define the expected header keywords for Types 1, 2, and 3 in a configuration section at the top of the script.*

## 4. Refactoring & Code Quality
* **Decoupling:** Define all file type mappings, naming patterns, and header keywords as global variables or a configuration dictionary at the top of the script. Do not hardcode "1.xls" or "2.xls" inside the functional logic.
* **Clear Old Data:** Before stacking begins, ensure the target sheets are cleared of old data while preserving the main header (Row 1).
* **Logging:** Update logs to report exactly how many files were found and stacked for each type (e.g., *"Found 3 files for Type 2 (Consumption). Appending batch to statement..."*).
* **Data Integrity:** Strictly treat all IDs, merchant numbers, and order numbers as **Strings** to prevent scientific notation conversion.
