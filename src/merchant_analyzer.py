"""
PROJECT: Merchant Anomaly Historical Report
LOGIC SPECIFICATION:
1. Target: Sheet '2商户交易日报'.
2. 3-Day Activity Filter: If count < 10 for T, T-1, AND T-2, skip merchant.
3. Anomaly Trigger:
   - Ratio: Past (T-1/T-2) is >= 1.5x (drop) or <= 0.5x (spike) of Today.
   - Zero/Missing: If one is $0 and other is >$1000 or >10 counts.
4. Output: If flagged, export FULL available history for that Merchant ID.
5. Formatting: Sort by Merchant ID (Asc), Date (Desc). Column widths fixed.
"""

import pandas as pd
import argparse
import os


def run_anomaly_detection(input_file):
    # --- CONFIGURABLE THRESHOLDS ---
    AMT_THRESHOLD = 1000  # Minimum USD amount to trigger an alert
    CNT_THRESHOLD = 10    # Minimum transaction count to trigger an alert

    # --- EXCEL COLUMN WIDTHS ---
    COL_WIDTHS = {
        'A': 20,  # 商户号 (Merchant ID)
        'B': 35,  # 商户名称 (Merchant Name)
        'C': 15,  # 日期 (Date)
        'D': 22,  # 支付成功金额USD (Amount)
        'E': 15,  # 支付成功笔数 (Count)
        'F': 55   # Flag Reason (Detailed anomaly description)
    }
    # -------------------------------

    try:
        df = pd.read_excel(input_file, sheet_name='2商户交易日报')
    except Exception as e:
        print(f"Error loading file: {e}")
        return

    # 1. Robust Pre-processing
    df['商户号'] = df['商户号'].astype(str)

    def clean_currency(x):
        if isinstance(x, str):
            clean_val = x.replace(',', '').replace('$', '').strip()
            return pd.to_numeric(clean_val, errors='coerce')
        return x

    df['支付成功金额USD'] = df['支付成功金额USD'].apply(clean_currency).fillna(0.0)
    df['支付成功笔数'] = pd.to_numeric(df['支付成功笔数'], errors='coerce').fillna(0).astype(int)
    df['日期'] = pd.to_datetime(df['日期'])

    # 2. Identify Anchors
    today_t = df['日期'].max()
    t_minus_1 = today_t - pd.Timedelta(days=1)
    t_minus_2 = today_t - pd.Timedelta(days=2)

    merchants_today = df[df['日期'] == today_t]['商户号'].unique()
    all_flagged_data = []

    for mid in merchants_today:
        # Get ALL historical records for this merchant from the source
        m_data = df[df['商户号'] == mid].copy()
        m_data['Flag Reason'] = "" # Initialize empty reason for all rows

        # Identify the specific rows for the 3-day window check
        row_t_idx = m_data[m_data['日期'] == today_t].index

        # Skip if merchant has no record for today
        if row_t_idx.empty:
            continue

        # 3-Day Activity Check
        count_t = m_data.loc[row_t_idx[0], '支付成功笔数']
        count_t1 = m_data[m_data['日期'] == t_minus_1]['支付成功笔数'].sum()
        count_t2 = m_data[m_data['日期'] == t_minus_2]['支付成功笔数'].sum()

        # Skip if all 3 days are below CNT_THRESHOLD
        if count_t < CNT_THRESHOLD and count_t1 < CNT_THRESHOLD and count_t2 < CNT_THRESHOLD:
            continue

        # Mark "Today" for clarity
        m_data.loc[row_t_idx, 'Flag Reason'] = "Today"

        row_t = m_data.loc[row_t_idx[0]]
        merchant_has_anomaly = False

        # Check T-1 and T-2 against Today
        for past_date in [t_minus_1, t_minus_2]:
            past_rows = m_data[m_data['日期'] == past_date]

            # Use real data if exists, else compare against zero-baseline
            if not past_rows.empty:
                row_p = past_rows.iloc[0]
                p_idx = past_rows.index[0]
            else:
                row_p = pd.Series({'支付成功金额USD': 0.0, '支付成功笔数': 0})
                p_idx = None

            reason = ""
            sig_t = (row_t['支付成功金额USD'] > AMT_THRESHOLD or row_t['支付成功笔数'] > CNT_THRESHOLD)
            sig_p = (row_p['支付成功金额USD'] > AMT_THRESHOLD or row_p['支付成功笔数'] > CNT_THRESHOLD)

            if sig_t and row_p['支付成功金额USD'] == 0:
                reason = "Anomaly (Zero to Non-zero > Threshold)"
            elif sig_p and row_t['支付成功金额USD'] == 0:
                reason = "Anomaly (Non-zero to Zero > Threshold)"
            elif sig_t or sig_p:
                for col, label in [('支付成功金额USD', 'Amount'), ('支付成功笔数', 'Count')]:
                    val_t = row_t[col]
                    val_p = row_p[col]
                    if val_t > 0:
                        ratio = val_p / val_t
                        if ratio >= 1.5:
                            reason = f"Anomaly ({label} Ratio {ratio:.1f}x drop vs {today_t.date()})"
                            break
                        elif ratio <= 0.5:
                            reason = f"Anomaly ({label} Ratio {ratio:.1f}x spike vs {today_t.date()})"
                            break

            if reason:
                merchant_has_anomaly = True
                # If the date existed in source, mark the reason on that row
                if p_idx is not None:
                    m_data.loc[p_idx, 'Flag Reason'] = reason
                else:
                    # If date didn't exist, create a dummy row to show the anomaly
                    new_row = {
                        '商户号': mid, '商户名称': row_t['商户名称'], '日期': past_date,
                        '支付成功金额USD': 0.0, '支付成功笔数': 0, 'Flag Reason': reason
                    }
                    m_data = pd.concat([m_data, pd.DataFrame([new_row])], ignore_index=True)


        if merchant_has_anomaly:
            all_flagged_data.append(m_data)

    # 3. Format and Export
    if all_flagged_data:
        result_df = pd.concat(all_flagged_data).drop_duplicates(subset=['商户号', '日期'])
        result_df['日期'] = result_df['日期'].dt.date

        # Sort by Merchant, then Date (Latest First)
        result_df = result_df.sort_values(by=['商户号', '日期'], ascending=[True, False])

        base_name = os.path.splitext(input_file)[0]
        output_file = f"{base_name}_Anomaly_Report.xlsx"
        cols = ['商户号', '商户名称', '日期', '支付成功金额USD', '支付成功笔数', 'Flag Reason']

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            result_df[cols].to_excel(writer, index=False, sheet_name='Anomalies')
            worksheet = writer.sheets['Anomalies']
            for col_letter, width in COL_WIDTHS.items():
                worksheet.column_dimensions[col_letter].width = width

        print(f"\nAnalysis Complete. Thresholds: ${AMT_THRESHOLD} / {CNT_THRESHOLD} counts.")
        print(f"Report: {output_file}")
    else:
        print(f"No significant anomalies found for {today_t.date()}.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input", help="Input Excel file")
    args = parser.parse_args()
    run_anomaly_detection(args.input)


if __name__ == "__main__":
    main()