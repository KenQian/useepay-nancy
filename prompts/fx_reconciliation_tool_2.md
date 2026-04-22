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


# Change this:
for r_idx, row in enumerate(account_statement_df.values, start=2):
    for c_idx, value in enumerate(row[:-1], start=1):
        ws_acc.cell(row=r_idx, column=c_idx).value = value
    ws_acc.cell(row=r_idx, column=13).value = f"=K{r_idx}-L{r_idx}" # Column M

# To this (adding a data check):
for r_idx, row in enumerate(account_statement_df.values, start=2):
    # Only write the row and formula if the first few columns aren't empty
    if pd.notna(row[0]) and str(row[0]).strip() != "":
        for c_idx, value in enumerate(row[:-1], start=1):
            ws_acc.cell(row=r_idx, column=c_idx).value = value
        
        # Only apply formula if there is data to calculate
        ws_acc.cell(row=r_idx, column=13).value = f"=K{r_idx}-L{r_idx}"
        ws_acc.cell(row=r_idx, column=18).value = f"=B{r_idx}&E{r_idx}"
        ws_acc.cell(row=r_idx, column=19).value = f"=G{r_idx}"