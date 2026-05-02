Master Plan

  - S1: Build source rows in 数据透视表
  - S2: Generate grouped summary in 数据透视表 from S1
  - S3: Publish S2 results into 1数透结果

  Rules that apply to all steps:

  - Use cell values, not formulas, for copying and calculations.
  - Treat this plan as the source of truth.
  - Implement in order: S1 -> S2 -> S3.
  - Do not start the next step until the current step is validated.

  S1: Copy filtered rows from 渠道订单 to 数据透视表

  - Purpose:
      - Rebuild the raw pivot-input area in 数据透视表 from 渠道订单.
  - Reads:
      - Sheet 渠道订单
      - Filter condition: column AO value equals 否
      - Source columns: AJ:AP
  - Writes:
      - Sheet 数据透视表
      - Preserve row 1
      - Clear rows 2 onward
      - Write source AJ:AP into target A:G
  - Dependency:
      - None
  - Output contract:
      - 数据透视表!A:G contains only rows copied from 渠道订单 where AO=否
      - Row 1 remains the existing header row
      - All written values are cell values, not formulas
  - Validation:
      - Rows 2+ in 数据透视表 are rebuilt from scratch
      - Row count matches the number of 渠道订单 rows where AO=否
      - Column mapping is exact: AJ->A, AK->B, AL->C, AM->D, AN->E, AO->F, AP->G

  S2: Generate grouped summary in 数据透视表

  - Purpose:
      - Build a pivot-table-like summary from the S1 output.
  - Reads:
      - Sheet 数据透视表
      - Source area: A:G produced by S1
  - Processing:
      - Group by columns A, C, D
      - Sum columns B and E
      - Sort by grouped keys in this order: A, then C, then D
  - Writes:
      - Sheet 数据透视表
      - Output columns:
          - A -> K
          - C -> L
          - D -> M
          - Sum(B) -> N
          - Sum(E) -> O
      - Append one final row after the grouped output:
          - K = Grand Total
          - N = total of all grouped N values
          - O = total of all grouped O values
      - Fill K:O of the Grand Total row with background color DarkSlateBlue
  - Dependency:
      - Depends on S1
  - Output contract:
      - 数据透视表!K:O contains the full grouped result set plus one Grand Total row
      - Grouping and totals are computed from values in S1, not from formulas
  - Validation:
      - Distinct grouped rows match unique (A,C,D) combinations from S1
      - N equals grouped sum of source B
      - O equals grouped sum of source E
      - Final row exists once and only once
      - Grand total row style is applied only to K:O of that row

  S3: Copy summary results to 1数透结果

  - Purpose:
      - Publish the S2 result into the final output sheet in two layouts.
  - Reads:
      - Sheet 数据透视表
      - Source area: K:O produced by S2
  - Writes:
      - Sheet `1数透结果`
      - Clear all rows in 1数透结果
      - Copy 数据透视表!K:O to 1数透结果!A:E, including header
      - Find the row in 1数透结果!A:E where:
          - A = CNY
          - B = USD
          - C = CNY
          - Fill that row with yellow background
      - Copy data from A:E to H:L with mapping:
          - A -> H
          - B -> J
          - C -> K
          - D -> I
          - E -> L
      - Exclude from the H:L copy:
          - the row where A=CNY, B=USD, C=CNY
          - the final Grand Total row
  - Dependency:
      - Depends on S2
  - Output contract:
      - 1数透结果!A:E is a full copy of S2 output
      - 1数透结果!H:L is a filtered/remapped copy excluding the highlighted row and Grand Total
  - Validation:
      - A:E matches 数据透视表!K:O
      - Exactly the target CNY/USD/CNY row is highlighted yellow if present
      - H:L excludes the highlighted row and excludes Grand Total
      - Column remapping is exact

  Implementation Order

  1. Implement S1 only.
  2. Validate S1 row count and column mapping.
  3. Implement S2 against accepted S1 output.
  4. Validate grouping, sorting, sums, and Grand Total.
  5. Implement S3 against accepted S2 output.
  6. Validate A:E copy, yellow-highlight row, and filtered H:L copy.

  Follow-up Prompt Pattern

  - “Implement S1 only.”
  - “Fix S1 only. Do not start S2.”
  - “S1 is accepted. Implement S2 only.”
  - “Fix S2 only. Keep S3 unchanged.”
  - “S2 is accepted. Implement S3 only.”

  Important Assumptions To Confirm Before Implementation

  - Whether “Delete all the rows from 1数透结果” means delete every row including any existing header row, or clear the sheet fully and rewrite without preserving a header.
  - Whether the yellow fill in 1数透结果 should apply to columns A:E only, or the full populated row.
  - Exact color code to use for DarkSlateBlue and Yellow, since Excel fill needs a concrete RGB/ARGB value.


-----------------------------------------------------------------------------------
We need to generate data for the sheet **`Estimated FX Summary`** (`预估换汇汇总`):

### 1. Clear Existing Data

* Remove all rows from the sheet except the header row.

---

### 2. Calculate `transactionDates` from **账户流水**

* Column A in **账户流水** contains datetime values in the format:
  `dd/MM/yyyy HH:mm:ss` (e.g., `20/04/2026 00:19:40`)
* Extract all **unique dates** (formatted as `yyyy-MM-dd`) from Column A.
* Sort the dates in ascending order.
* Generate a string `transactionDates` using the following rules:

  * If there is only one date → use that date directly.
  * If there are multiple dates →
    use the first date, then append `&dd` for each additional date.

    * Example:
      `2026-05-30, 2026-05-31, 2026-06-01` → `2026-05-30&31&01`

---

### 3. Copy and Transform Data from Source Sheet

Copy data from **`Source` (`1数透结果`)**, columns **H:L**, into **Target** (`Estimated FX Summary`), excluding the header row.

#### Column Mapping:

* **Target Column A** = `transactionDates` (calculated above)
* **Target Columns B–D** = Source Columns H–J
* **Target Columns G–H** = Source Columns K–L

---

### 4. Calculated Fields

* **Column E**:
  If `B[row] == D[row]`, then `1`, otherwise:

  ```
  IF(B[row]=D[row],1,XLOOKUP(B[row]&D[row],'每日汇率(oc系统中获取）'!I:I,'每日汇率(oc系统中获取）'!H:H))
  ```

* **Column F**:

  ```
  C[row] * E[row]
  ```

* **Column I**:

  ```
  H[row] * (1 - 3%)
  ```

* **Column J**:

  ```
  D[row] & G[row]
  ```
-----------------------------------------------------------------------------------
# Generate FX Transaction Analysis Sheets and Tables

---

## 1. Create a New Sheet
- Use the existing `transaction_dates` (e.g., `2026-05-30&31&01` or `2026-05-30`)
- Transform it by:
  - Removing `yyyy-`
  - Removing the `-` between `mm` and `dd`  
  - Example:  
    `2026-05-30&31&01` → `0530&31&01`
- Construct the sheet name as:  
  **`数透` + transformed date**  
  - Example: `数透0530&31&01`
- Insert this sheet as the **first sheet** in the workbook
- Use `fx_transaction` as the prefix for related variables

---

## 2. Create Data Blocks (Tables A–F)

### Common Formatting Rules
- Fill the **top-left cell of each table area** with **yellow**
- Fill **header rows** with color `FFF2F1F7`
- Tail (total) rows (if any) also use `FFF2F1F7`

---

### 2.1 Table A (`fx_transaction_table_a`)
**Purpose:** Pivot-style summary from `预估换汇汇总`

**Input:**
- Sheet `预估换汇汇总`

**Processing:**
- Group by columns **D and G**
- Aggregate:
  - Sum(E)
  - Sum(H)
- Sort by D, then G

**Output:**
- Location: starts from row 1
- **Write `表格A` in the top-left cell of the table area**
- Header:
  - 打款币种, 清算币种, 求和项:预估通道打款金额（已扣除手续费3.2%）, 求和项:清算净额（扣除收费）
- Mapping:
  - D → A
  - G → B
  - Sum(F) → C
  - Sum(H) → D
- Add final row:
  - A = `Grand Total`
  - C = total of Column C
  - D = total of Column D

---

### 2.2 Table B (`fx_transaction_table_b`)
**Purpose:** Currency pairs that do **NOT** require netting

**Input:**
- Table A

**Processing:**
- A pair (A, B) requires netting if a reverse pair (B, A) exists
- Otherwise, it does not require netting

**Output:**
- Position: 10 rows below Table A
- **Write `表格B(非轧差)` in the top-left cell of the table area**
- Header:
  - 打款币种, 清算币种, 求和项:预估通道打款金额（已扣除手续费3.2%）, 求和项:清算净额（扣除收费）
- Copy all rows that do not require netting

---

### 2.3 Table C (`fx_transaction_table_c`)
**Purpose:** Currency pairs that **require netting**

**Input:**
- Table A

**Output:**
- Position: right of Table B (start from column H), aligned to the same row
- **Write `表格C(轧差)` in the top-left cell of the table area**
- No header

**Processing:**
- For each reversible pair (A,B) & (B,A):
  - Row 1: first record
  - Row 2: reverse record
  - Row 3: net result:
    - J3 = J1 − J2 > 0
    - K3 = K1 − K2 > 0
    - H3 = H1, I3 = I1
  - Fill row 3 with **gray color**
  - Insert one empty row between groups

---

### 2.4 Table D (`fx_transaction_table_d`)
**Purpose:** Combine:
- Non-netting records (Table B)
- Netting results (Table C row 3 only)

**Processing:**
- Merge datasets
- Sort by:
  - First column in the merged dataset
  - Second column in the merged dataset

**Output:**
- Position: below the lower of Table B / Table C + 10 rows
- **Write `表格D(将轧差后的数据一起汇总)` in the top-left cell of the table area**
- Header:
  - 打款币种, 清算币种, 求和项:预估通道打款金额（已扣除手续费3.2%）, 求和项:清算净额（扣除收费）
- Highlight netting result rows (from Table C) in **red**

---

### 2.5 Table E (`fx_transaction_table_e`)
**Purpose:** Trade monitoring data

**Input:**
- Table D

**Output:**
- Position: 10 rows below Table D
- **Write `表格E(盯盘所需数据)` in the top-left cell of the table area**
- Header:
  - 卖出币种, 卖出金额, 买入币种, 买入金额
- Mapping:
  - A → A
  - C → B
  - B → C
  - D → D

---

### 2.6 Table F (`fx_transaction_table_f`)
**Purpose:** Final monitoring dataset

**Input:**
- Table E

**Output:**
- Position: right of Table E (start from column H), aligned to the same row
- **Write `表格F(最终数据)` in the top-left cell of the table area**
- Header:
  - 卖出币种, 卖出金额, 买入币种
- Mapping:
  - A → H
  - B → I (**round to the nearest hundred**)
  - C → J 

---

## 3. Create Summary Table in `预估换汇汇总`

Create table: **`fx_transaction_summary_table`**

**Input:**
- Table F (`fx_transaction_table_f`)

**Processing:**
- Find the last row in `预估换汇汇总`
- Insert the table after **10 empty rows**

---

### 3.1 Output: Summary Table
- Header:
  - 卖出币种, 卖出金额, 买入币种, 买入金额, 备注
- Mapping:
  - H → A
  - I → B
  - J → C

---

### 3.2 Generate Remarks (备注)
- Position:
  - On the **right side of `fx_transaction_summary_table`**', align with its header row.
  - Merge header row columns **G–J**

- Logic:
  - Iterate through **each row in `fx_transaction_summary_table`**
  - For each row:
    - If `B[row] == 0` → skip
    - If `B[row] > 0` →  
      `用B[row]A[row]换C[row]`
    - If `B[row] < 0` →  
      `用C[row]换B[row]A[row]`  
      (remove the negative sign from `B[row]`)

- Final Step:
  - Concatenate all generated strings using **Chinese comma (`，`)**
  - **Output the final joined string into the merged cell**


----------------------------
## Update Formatting and Layout

### 1. Update Bottom Table in `Estimated FX Summary` (`预估换汇汇总`)

The sheet contains two tables:
- Top table (from `1数透结果`)
- Bottom table (from `Table F (fx_transaction_table_f)`)

> ⚠️ Apply the following changes **only to the bottom table**

#### Formatting Rules
- **Header Row**
  - Set row height to **50**
  - Set alignment to:
    - Vertical: **middle**
    - Horizontal: **center**

- **Data Rows**
  - Columns **A and C**:
    - Set horizontal alignment to **center**
  - Column **B**:
    - Format: **Number**
    - Decimal places: **2**

---

### 2. Remarks Area (Right Side)
- Enable **wrap text** for the remarks cell(s)

---

### 3. Update Column Widths in Sheet `数透xxx`

> Apply **before creating Tables A–F**

Set column widths as follows:
- A: 16  
- B: 16  
- C: 40  
- D: 25  
- H: 15  
- I: 15  
- J: 15  
- K: 15  

---

### 4. Set Default Active Sheet
- Set **`预估换汇汇总`** as the active sheet at the end



fx_reconciliation_core.py is responsible for preparing for the data by importing them from other files, 
and then human needs to be involved to fill some missing data by looking into the production system. 
fx_consolidation_postprocess.py is responsible for processing these to generate/update multiple sheets with summarized data. 

fx_consolidation_postprocess.py might fail because the top table in 预估换汇汇总 needs to use B&D to lookup the column H from 每日汇率(oc系统中获取）.
Here is the data flow 渠道订单 => 数据透视表 => 1数透结果 => 预估换汇汇总. In order to avoid the lookup failure, we need to check if there's missing record in 每日汇率(oc系统中获取）.
Here is the requirement to add logic in fx_reconciliation_core.py:
- Find all the unique combination of AJ and AL in 渠道订单, and then find the record in 每日汇率(oc系统中获取）by its column I (its formula is D&E). 
  If no record can be found, append a row at the end of 每日汇率(oc系统中获取）, filling D = AJ, E = AL, I = D&E
- Include the information in the log summary and 处理摘要

## FX Reconciliation & Post-Processing Enhancement

### Context

- `fx_reconciliation_core.py`
  - Responsible for importing and preparing raw data from multiple sources
  - Requires **manual intervention** to fill in missing data based on the production system

- `fx_consolidation_postprocess.py`
  - Responsible for processing prepared data
  - Generates and updates multiple sheets with summarized results

---

### Problem

`fx_consolidation_postprocess.py` may fail due to missing lookup data.

- In sheet **`预估换汇汇总`**, the top table uses:
  - Columns **B & D** → to lookup column **H**
  - Lookup source: **`每日汇率(oc系统中获取）`**

- Data flow:
  `渠道订单 → 数据透视表 → 1数透结果 → 预估换汇汇总`


- Root cause:
  - Missing records in **`每日汇率(oc系统中获取）`** lead to lookup failures

---

### Solution Requirement

Add validation and auto-repair logic in `fx_reconciliation_core.py` to ensure all required lookup records exist.

---

### Implementation Logic

1. **Extract Required Keys**
 - From sheet **`渠道订单`**
 - Collect all **unique combinations of (AJ, AL)**

2. **Validate Against Exchange Rate Table**
 - Target sheet: **`每日汇率(oc系统中获取）`**
 - Lookup column: **Column I**
   - (Column I is derived from `D & E`)

3. **Handle Missing Records**
 - For each `(AJ, AL)` combination:
   - If no matching record exists in Column I:
     - Append a new row at the end of the sheet
     - Populate:
       - Column D = AJ
       - Column E = AL
       - Column I = D & E (same formula/logic)
 - Human will fill the data you appended later.
---

### Logging & Reporting

- Include all auto-added records in:
- **Log summary**
- **处理摘要**

- Each log entry should include:
  - Missing key (AJ, AL)
  - Whether a new row was inserted

---

### Expected Outcome

- Ensure all required lookup keys exist before post-processing
  - Prevent failures in `fx_consolidation_postprocess.py`
  - Improve data completeness and pipeline robustness


----

Then some fix is needed for appending records to sheets 二级商户号映射表-A07 and 打款币种. 

Following is the output:

--- log start ---
20:39:29 - INFO - 
Added back to 渠道订单 from 特殊的渠道订单
(no records)
20:39:29 - INFO - 
Added rows in 二级商户号映射表-A07
Row | A-二级商户号                          
----+----------------------------------
909 | MICROS-spiderfarmer.co.uk        
910 | MICROS-jimmy.eu                  
911 | MICROS-marshydroled.co.uk        
912 | VTOMAN-us.gooloo.com-klarna      
913 | SPHERE UK-pixarbikes.co.uk-klarna
914 | Oktyple-www.sposadresses.com     
915 | Nekoya FR-m.moonycozy.com-link   
916 | Oktyple US-okbridal.co.uk        
20:39:29 - INFO - 
Added rows in 打款币种
Row  | B-通道名称 | D-二级商户号                            | E-交易币种 | G-Column 7                                
-----+--------+------------------------------------+--------+-------------------------------------------
1720 | A07    | MICROS-spiderfarmer.co.uk          | GBP    | A07MICROS-spiderfarmer.co.ukGBP           
1721 | A07    | MICROS-jimmy.eu                    | EUR    | A07MICROS-jimmy.euEUR                     
1722 | A07    | MICROS-marshydroled.co.uk          | GBP    | A07MICROS-marshydroled.co.ukGBP           
1723 | A07    | ZONGHENG-TapRead                   | JPY    | A07USZONGHENG-TapReadJPY                  
1724 | A07    | VTOMAN-us.gooloo.com-klarna        | USD    | A07VTOMAN-us.gooloo.com-klarnaUSD         
1725 | A07    | TENZHUO HK-Readink                 | UYU    | A07HKTENZHUO HK-ReadinkUYU                
1726 | A07    | SPHERE UK-pixarbikes.co.uk-klarna  | GBP    | A07SPHERE UK-pixarbikes.co.uk-klarnaGBP   
1727 | A07    | Oktyple-www.sposadresses.com       | USD    | A07Oktyple-www.sposadresses.comUSD        
1728 | A07    | BABARONI-www.babaroni.co.uk-klarna | SEK    | A07GBBABARONI-www.babaroni.co.uk-klarnaSEK
1729 | A07    | Nekoya FR-m.moonycozy.com-link     | EUR    | A07Nekoya FR-m.moonycozy.com-linkEUR      
1730 | A07    | Oktyple US-okbridal.co.uk          | USD    | A07Oktyple US-okbridal.co.ukUSD           
1731 | A07    | NEXTMARVEL US- www.zeelool.com     | EUR    | A07USNEXTMARVEL US- www.zeelool.comEUR    
--- log end ---

Following records should not be added because "same"(only case differences) data is already there. Following are all start with `MICROS`, however, the records in the corresponding sheet start with `Micros`. 

909 | MICROS-spiderfarmer.co.uk        
910 | MICROS-jimmy.eu                  
911 | MICROS-marshydroled.co.uk  

1720 | A07    | MICROS-spiderfarmer.co.uk          | GBP    | A07MICROS-spiderfarmer.co.ukGBP           
1721 | A07    | MICROS-jimmy.eu                    | EUR    | A07MICROS-jimmy.euEUR                     
1722 | A07    | MICROS-marshydroled.co.uk          | GBP    | A07MICROS-marshydroled.co.ukGBP           
