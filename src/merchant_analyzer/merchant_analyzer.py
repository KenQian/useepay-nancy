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

import argparse
import logging
import os
from datetime import datetime

import pandas as pd


LOG_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
DATE_FORMAT = '%H:%M:%S'


def configure_run_logging(log_path):
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    for handler in list(logger.handlers):
        if getattr(handler, '_merchant_file_handler', False):
            logger.removeHandler(handler)
            handler.close()

    file_handler = logging.FileHandler(log_path, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter(LOG_FORMAT, datefmt=DATE_FORMAT))
    file_handler._merchant_file_handler = True
    logger.addHandler(file_handler)


def ensure_result_dir(input_file):
    result_dir = os.path.join(os.path.dirname(os.path.abspath(input_file)), 'result')
    os.makedirs(result_dir, exist_ok=True)
    return result_dir


def run_merchant_analyzer(input_file, log_path=None):
    input_file = os.path.abspath(input_file)
    if not os.path.isfile(input_file):
        raise FileNotFoundError(f"Input Excel file not found: {input_file}")

    result_dir = ensure_result_dir(input_file)
    if log_path is None:
        timestamp = datetime.now().strftime('%Y%m%d-%H%M%S')
        log_path = os.path.join(result_dir, f'merchant_analyzer_{timestamp}.log')
    configure_run_logging(log_path)
    logging.info("Starting Merchant Analyzer...")
    logging.info("Input file: %s", input_file)

    amt_threshold = 1000
    cnt_threshold = 10
    col_widths = {
        'A': 20,
        'B': 35,
        'C': 15,
        'D': 22,
        'E': 15,
        'F': 55,
    }

    logging.info("Loading source workbook...")
    df = pd.read_excel(input_file, sheet_name='2商户交易日报')
    logging.info("Loaded %s rows from 2商户交易日报", len(df.index))

    df['商户号'] = df['商户号'].astype(str)

    def clean_currency(x):
        if isinstance(x, str):
            clean_val = x.replace(',', '').replace('$', '').strip()
            return pd.to_numeric(clean_val, errors='coerce')
        return x

    logging.info("Pre-processing source data...")
    df['支付成功金额USD'] = df['支付成功金额USD'].apply(clean_currency).fillna(0.0)
    df['支付成功笔数'] = pd.to_numeric(df['支付成功笔数'], errors='coerce').fillna(0).astype(int)
    df['日期'] = pd.to_datetime(df['日期'])

    today_t = df['日期'].max()
    t_minus_1 = today_t - pd.Timedelta(days=1)
    t_minus_2 = today_t - pd.Timedelta(days=2)

    merchants_today = df[df['日期'] == today_t]['商户号'].unique()
    all_flagged_data = []
    total_merchants = len(merchants_today)

    logging.info("Analyzing merchants for %s...", today_t.date())
    for idx, mid in enumerate(merchants_today, start=1):
        if idx == 1 or idx % 500 == 0 or idx == total_merchants:
            logging.info("Processing merchants: %s/%s", idx, total_merchants)

        m_data = df[df['商户号'] == mid].copy()
        m_data['Flag Reason'] = ""
        row_t_idx = m_data[m_data['日期'] == today_t].index

        if row_t_idx.empty:
            continue

        count_t = m_data.loc[row_t_idx[0], '支付成功笔数']
        count_t1 = m_data[m_data['日期'] == t_minus_1]['支付成功笔数'].sum()
        count_t2 = m_data[m_data['日期'] == t_minus_2]['支付成功笔数'].sum()

        if count_t < cnt_threshold and count_t1 < cnt_threshold and count_t2 < cnt_threshold:
            continue

        m_data.loc[row_t_idx, 'Flag Reason'] = "Today"
        row_t = m_data.loc[row_t_idx[0]]
        merchant_has_anomaly = False

        for past_date in [t_minus_1, t_minus_2]:
            past_rows = m_data[m_data['日期'] == past_date]

            if not past_rows.empty:
                row_p = past_rows.iloc[0]
                p_idx = past_rows.index[0]
            else:
                row_p = pd.Series({'支付成功金额USD': 0.0, '支付成功笔数': 0})
                p_idx = None

            reason = ""
            sig_t = (row_t['支付成功金额USD'] > amt_threshold or row_t['支付成功笔数'] > cnt_threshold)
            sig_p = (row_p['支付成功金额USD'] > amt_threshold or row_p['支付成功笔数'] > cnt_threshold)

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
                if p_idx is not None:
                    m_data.loc[p_idx, 'Flag Reason'] = reason
                else:
                    new_row = {
                        '商户号': mid,
                        '商户名称': row_t['商户名称'],
                        '日期': past_date,
                        '支付成功金额USD': 0.0,
                        '支付成功笔数': 0,
                        'Flag Reason': reason,
                    }
                    m_data = pd.concat([m_data, pd.DataFrame([new_row])], ignore_index=True)

        if merchant_has_anomaly:
            all_flagged_data.append(m_data)

    base_name = os.path.splitext(os.path.basename(input_file))[0]
    output_file = os.path.join(result_dir, f"{base_name}_Anomaly_Report.xlsx")

    if all_flagged_data:
        logging.info("Formatting anomaly report...")
        result_df = pd.concat(all_flagged_data).drop_duplicates(subset=['商户号', '日期'])
        result_df['日期'] = result_df['日期'].dt.date
        result_df = result_df.sort_values(by=['商户号', '日期'], ascending=[True, False])

        cols = ['商户号', '商户名称', '日期', '支付成功金额USD', '支付成功笔数', 'Flag Reason']
        logging.info("Saving anomaly report workbook...")
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            result_df[cols].to_excel(writer, index=False, sheet_name='Anomalies')
            worksheet = writer.sheets['Anomalies']
            for col_letter, width in col_widths.items():
                worksheet.column_dimensions[col_letter].width = width

        logging.info("Analysis complete. Thresholds: $%s / %s counts.", amt_threshold, cnt_threshold)
        logging.info("Report: %s", output_file)
        return {
            'output_file': output_file,
            'log_path': log_path,
            'today': str(today_t.date()),
            'flagged_merchants': len({frame['商户号'].iloc[0] for frame in all_flagged_data}),
            'anomaly_rows': len(result_df.index),
            'message': "Analysis Complete",
        }

    logging.info("No significant anomalies found for %s.", today_t.date())
    return {
        'output_file': "",
        'log_path': log_path,
        'today': str(today_t.date()),
        'flagged_merchants': 0,
        'anomaly_rows': 0,
        'message': f"No significant anomalies found for {today_t.date()}",
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input", help="Input Excel file")
    args = parser.parse_args()
    try:
        run_merchant_analyzer(args.input)
    except Exception as exc:
        logging.error("Merchant Analyzer failed: %s", exc)
        raise


if __name__ == "__main__":
    main()
