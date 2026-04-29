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


---
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
