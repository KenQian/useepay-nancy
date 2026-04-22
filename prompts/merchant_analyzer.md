To incorporate your latest local changes (specifically the **Full History** requirement and the **3-Day Activity Filter**), we need to refine the prompt so it clearly distinguishes between the **Triggering Logic** (what causes a flag) and the **Output Content** (what rows end up in the result).

Here is the updated, "Ironclad" version of your prompt.

---

# Data Analysis Prompt: Merchant Anomaly Historical Report

## 1. Task Objective
Analyze the sheet **`2商户交易日报`** to identify merchants with transaction anomalies based on a 3-day window. For every merchant flagged, output their **entire available history** from the source data to provide full context.

## 2. Data Pre-processing Rules
* **Target Sheet:** `2商户交易日报`. Do not skip hidden, filtered, or grouped rows.
* **Merchant ID (商户号):** Treat as **String/Text** to prevent scientific notation.
* **Amount (支付成功金额USD):** Clean strings (remove commas/symbols) and convert to **Float**.
* **Count (支付成功笔数):** Convert to **Integer**.
* **Missing Dates:** If a merchant has a record for Today ($T$) but is missing a row for $T-1$ or $T-2$, treat that missing day as **\$0.00 amount and 0 count**.

## 3. The "3-Day Activity" Filter
Before performing anomaly checks, apply a noise filter. 
* **Rule:** For a given merchant, if the **Successful Payment Count** is less than **10** on **all three days** ($T$, $T-1$, and $T-2$), skip that merchant entirely.
* **Requirement:** The merchant must have at least one day in the 3-day window where the count is $\ge 10$.

## 4. Anomaly Detection Logic (The Trigger)
A merchant is "Flagged" if the comparison of **Today ($T$) vs. $T-1$** OR **Today ($T$) vs. $T-2$** meets any of these conditions:

1.  **Threshold Trigger (Zero/Missing):**
    * One day is $\$0$ (or missing) and the other is non-zero. 
    * Flag if the non-zero value is **$>\$1,000$** or the count is **$>10$**.
2.  **Ratio Trigger (Standard):**
    * Both days are non-zero.
    * Flag if Past ($T-1$ or $T-2$) is **$\ge 1.5x$** Today's value (a drop occurred today).
    * Flag if Past ($T-1$ or $T-2$) is **$\le 0.5x$** Today's value (a spike occurred today).

## 5. Output Rules (Full Context)
If a merchant is flagged based on the logic in Section 4:
1.  **Include ALL records:** Output **every row** for that Merchant ID found in the source sheet (including $T-3$, $T-4$, etc., if available).
2.  **Flag Reason Column:** Add a new column named **`Flag Reason`**.
    * Mark the Today row as `"Today"`.
    * Mark the specific past rows ($T-1$ or $T-2$) with the reason (e.g., `"Anomaly (Amount Ratio 1.8x drop vs [Date])"`).
    * Leave other historical rows (like $T-3$) blank in this column.
3.  **Formatting:** Generate a downloadable **.xlsx** file. 
    * **Sort:** By 商户号 (Ascending) and 日期 (Descending).
    * **Sheet Name:** `result`.

## 6. Validation Example
**Merchant 500000000010301:**
* *T (04-02):* \$3,493.95 (12 counts)
* *T-1 (04-01):* \$1,304.99 (5 counts)
* *T-2 (03-31):* \$2,723.96 (11 counts)
* **Logic:** $T-1$ vs $T$ Amount Ratio is $0.37x$ ($\le 0.5x$). 
* **Result:** Flagged. Output **all three rows** (and any other dates for this ID) to the result sheet.

---

### Key Improvements made to your prompt:
* **Clarified "Missing Data":** Explicitly told the AI to treat missing $T-1/T-2$ rows as zeros so it doesn't skip them.
* **The "Full History" Mandate:** Explicitly stated that once a merchant is caught, the *entire* available history must be exported.
* **The Activity Filter:** Integrated the `< 10 count` skip rule to ensure the AI doesn't waste time on low-volume data.
* **Clean Labels:** Instructions for the `Flag Reason` column ensure the report is professional and easy to read.