"""
FX Settlement Automation Tool
-----------------------------
Purpose:
    Automates the daily reconciliation of Foreign Exchange (FX) channel settlements.
    Processes account statements, validates channel orders against historical
    exceptions, and applies live Excel formulas for financial auditing.
"""

import os
import argparse
import shutil
import logging
import re
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

# --- Configuration & Decoupling ---
FILE_CONFIG = {
    'type_1': {'pattern': r'^1(-.*)?\.xls$', 'desc': 'Refunds', 'header_keys': ['商户号', '订单类型']},
    'type_2': {'pattern': r'^2(-.*)?\.xls$', 'desc': 'Consumption', 'header_keys': ['商户号', '订单类型']},
    'type_3': {'pattern': r'^3(-.*)?\.xls$', 'desc': 'Channel Orders', 'header_keys': ['交易类型', '商户订单号']},
}

# Setup logging configuration
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%H:%M:%S'
)

READ_EXCEL_TEXT_KWARGS = {
    'dtype': str,
    'keep_default_na': False,
}


def get_source_files(directory, pattern_str):
    """Finds all files in directory matching a regex pattern, sorted."""
    regex = re.compile(pattern_str, re.IGNORECASE)
    matched_files = [f for f in os.listdir(directory) if regex.match(f)]
    matched_files.sort()
    return [os.path.join(directory, f) for f in matched_files]


def file_has_header(path, header_keywords):
    """
    Detect whether the first row is a header row.

    We intentionally read with header=None first so stacked files always keep
    a consistent positional schema regardless of whether each source file has
    a header row.
    """
    preview_df = pd.read_excel(
        path,
        sheet_name=0,
        header=None,
        **READ_EXCEL_TEXT_KWARGS,
    )
    if preview_df.empty:
        return False

    first_row = preview_df.iloc[0].fillna('').astype(str)
    first_row_str = " ".join(first_row.tolist())
    return any(key in first_row_str for key in header_keywords)


def load_and_stack_files(file_paths, header_keywords):
    """Iteratively loads and stacks multiple files using a positional schema."""
    stacked_frames = []

    for path in file_paths:
        df = pd.read_excel(
            path,
            sheet_name=0,
            header=None,
            **READ_EXCEL_TEXT_KWARGS,
        )
        if df.empty:
            logging.info(f"Skipping empty file: {os.path.basename(path)}")
            continue

        has_header = file_has_header(path, header_keywords)
        if has_header:
            df = df.iloc[1:].reset_index(drop=True)

        # Normalize to a positional schema so mixed header/no-header source
        # files stack cleanly.
        df.columns = range(df.shape[1])

        # Drop fully blank rows that often appear at the tail of exported xls files.
        blank_row_mask = df.apply(
            lambda row: all(str(value).strip() == '' for value in row),
            axis=1,
        )
        df = df[~blank_row_mask].reset_index(drop=True)
        if df.empty:
            logging.info(f"Skipping file with only header/blank rows: {os.path.basename(path)}")
            continue

        logging.info(
            "Loaded %s rows from %s (header_detected=%s)",
            len(df),
            os.path.basename(path),
            has_header,
        )
        stacked_frames.append(df)

    if not stacked_frames:
        return pd.DataFrame()

    return pd.concat(stacked_frames, ignore_index=True)


def get_latest_baseline(root_dir):
    """Finds the most recent non-WIP baseline file in the directory."""
    files = [f for f in os.listdir(root_dir) if f.startswith('各通道需换汇情况汇总-') and f.endswith('.xlsx') and '-wip' not in f]
    if not files:
        raise FileNotFoundError("Baseline file not found.")
    files.sort(reverse=True)
    return files[0]


def main():
    parser = argparse.ArgumentParser(description="FX Channel Settlement Automation Tool")
    parser.add_argument("directory", help="The root directory containing source files.")

    import sys
    if len(sys.argv) == 1:
        parser.print_help(sys.stderr); sys.exit(1)

    args = parser.parse_args()
    root = args.directory

    logging.info("Starting FX Settlement Automation...")

    # 1. File Discovery
    try:
        latest_baseline = get_latest_baseline(root)
        baseline_path = os.path.join(root, latest_baseline)

        type1_files = get_source_files(root, FILE_CONFIG['type_1']['pattern'])
        type2_files = get_source_files(root, FILE_CONFIG['type_2']['pattern'])
        type3_files = get_source_files(root, FILE_CONFIG['type_3']['pattern'])

        logging.info(f"Baseline: {latest_baseline}")
        logging.info(
            "Found %s Refund files, %s Consumption files, %s Channel Order files",
            len(type1_files),
            len(type2_files),
            len(type3_files),
        )

        if not type1_files or not type2_files or not type3_files:
            logging.error("Missing one or more required source file types (1.xls, 2.xls, or 3.xls patterns).")
            return

    except Exception as e:
        logging.error(f"Initialization Failed: {e}"); return

    # 2. Workspace Creation
    today_str = datetime.now().strftime('%Y%m%d')
    wip_filename = f"各通道需换汇情况汇总-{today_str}-wip.xlsx"
    wip_path = os.path.join(root, wip_filename)
    shutil.copy2(baseline_path, wip_path)

    # 3. Data Ingestion & Stacking (Shadow Calculations)
    logging.info("Stacking %s Refund files for Module A...", len(type1_files))
    refunds_all = load_and_stack_files(type1_files, FILE_CONFIG['type_1']['header_keys'])
    logging.info("Stacking %s Consumption files for Module A...", len(type2_files))
    consumption_all = load_and_stack_files(type2_files, FILE_CONFIG['type_2']['header_keys'])

    account_statement_df = pd.concat([consumption_all, refunds_all], ignore_index=True)
    account_statement_df['Internal_Key_R'] = account_statement_df.iloc[:, 1].fillna('') + account_statement_df.iloc[:, 4].fillna('')

    logging.info("Stacking %s Channel Order files for Module B...", len(type3_files))
    channel_orders_raw = load_and_stack_files(type3_files, FILE_CONFIG['type_3']['header_keys'])
    channel_orders_filtered = channel_orders_raw[~channel_orders_raw.iloc[:, 9].isin(['预授权申请', '预授权撤销'])]

    # Module B Source 2: Re-validation
    special_orders_df = pd.read_excel(
        baseline_path,
        sheet_name='特殊的渠道订单',
        **READ_EXCEL_TEXT_KWARGS,
    )
    valid_from_special = pd.DataFrame()
    if not special_orders_df.empty:
        special_orders_df['Calc_AI_Key'] = special_orders_df.iloc[:, 3].fillna('') + special_orders_df.iloc[:, 4].fillna('')
        valid_from_special = special_orders_df[special_orders_df['Calc_AI_Key'].isin(account_statement_df['Internal_Key_R'])]

    # 4. Writing to Excel
    wb = load_workbook(wip_path)

    # Update Account Statement Sheet
    ws_acc = wb['账户流水']
    ws_acc.delete_rows(2, ws_acc.max_row) # Clear Old Data

    keys_to_remove_from_statement = set()
    ws_chan = wb['渠道订单']
    ws_chan.delete_rows(2, ws_chan.max_row)

    ws_spec = wb['特殊的渠道订单']
    if ws_spec.max_row > 1:
        ws_spec.delete_rows(2, ws_spec.max_row)

    # Prepare Target Data Pool
    target_data = []
    for _, row in channel_orders_filtered.iterrows():
        new_row = list(row.iloc[0:4]) + list(row.iloc[5:34]) + [row.iloc[4]]
        target_data.append(new_row)
    if not valid_from_special.empty:
        for _, row in valid_from_special.iterrows():
            target_data.append(list(row.iloc[0:34]))

    # Mapping Sheets
    mapping_sheets = {
        '打款币种': pd.read_excel(wip_path, sheet_name='打款币种', **READ_EXCEL_TEXT_KWARGS),
        '渠道名称': pd.read_excel(wip_path, sheet_name='渠道名称', **READ_EXCEL_TEXT_KWARGS)
    }

    # Module D & Logic Resolution
    curr_row = 2
    exceptions_moved = 0
    chan_map = mapping_sheets['渠道名称'].set_index(mapping_sheets['渠道名称'].iloc[:, 0]).iloc[:, 1]

    for data_row in target_data:
        col_f, col_d, col_e = data_row[5], data_row[3], data_row[4]
        ai_key = f"{col_d}{col_e}"
        ap_val = str(chan_map.get(col_f, "")).strip()
        is_na = account_statement_df[account_statement_df['Internal_Key_R'] == ai_key].empty

        if is_na and ap_val.lower() not in ["paypal", "afterpay直连"]:
            # Move to Special
            next_spec = ws_spec.max_row + 1
            for c_idx, val in enumerate(data_row, start=1):
                ws_spec.cell(row=next_spec, column=c_idx).value = val
            ws_spec.cell(row=next_spec, column=35).value = f"=D{next_spec}&E{next_spec}"
            ws_spec.cell(row=next_spec, column=36).value = f"=XLOOKUP(A{next_spec},账户流水!$R:$R,账户流水!$R:$R)"
            keys_to_remove_from_statement.add(ai_key)
            exceptions_moved += 1
            continue

        # Write to Main Channel Orders
        for c_idx, val in enumerate(data_row, start=1):
            ws_chan.cell(row=curr_row, column=c_idx).value = val

        # Formulas
        ws_chan.cell(row=curr_row, column=35).value = f"=D{curr_row}&E{curr_row}"
        ws_chan.cell(row=curr_row, column=36).value = f"=N{curr_row}"
        ws_chan.cell(row=curr_row, column=37).value = f'=IF(I{curr_row}="退款", -M{curr_row}*(1-0.032), M{curr_row}*(1-0.032))'
        ws_chan.cell(row=curr_row, column=38).value = f"=XLOOKUP(AR{curr_row}, 打款币种!$G:$G, 打款币种!$F:$F)"
        ws_chan.cell(row=curr_row, column=39).value = f"=VLOOKUP(AI{curr_row}, 账户流水!$R:$T, 2, 0)"
        ws_chan.cell(row=curr_row, column=40).value = f"=VLOOKUP(AI{curr_row}, 账户流水!$R:$T, 3, 0)"
        ws_chan.cell(row=curr_row, column=41).value = f"=IF(AL{curr_row}=AM{curr_row}, \"是\", \"否\")"
        ws_chan.cell(row=curr_row, column=42).value = f"=VLOOKUP(F{curr_row}, 渠道名称!$A:$B, 2, 0)"

        aq_f = (f'IF(AP{curr_row}="2号通道", XLOOKUP(AH{curr_row}, \'二级商户号映射表-A01\'!$A:$A, \'二级商户号映射表-A01\'!$B:$B), '
                f'IF(AP{curr_row}="A07", XLOOKUP(AH{curr_row}, \'二级商户号映射表-A07\'!$A:$A, \'二级商户号映射表-A07\'!$C:$C) & AH{curr_row}, '
                f'IF(AP{curr_row}="7号通道", AB{curr_row}, "")))')
        ws_chan.cell(row=curr_row, column=43).value = f"={aq_f}"
        ws_chan.cell(row=curr_row, column=44).value = f"=AP{curr_row}&AQ{curr_row}&AJ{curr_row}"
        curr_row += 1

    # Apply Cleanup & Final Statement Write
    if keys_to_remove_from_statement:
        account_statement_df = account_statement_df[~account_statement_df['Internal_Key_R'].isin(keys_to_remove_from_statement)]

    for r_idx, row in enumerate(account_statement_df.values, start=2):
        # Only write the row and formula if the first few columns aren't empty
        if pd.notna(row[0]) and str(row[0]).strip() != "":
            for c_idx, value in enumerate(row[:-1], start=1):
                ws_acc.cell(row=r_idx, column=c_idx).value = value

            # Only apply formula if there is data to calculate
            ws_acc.cell(row=r_idx, column=13).value = f"=K{r_idx}-L{r_idx}"
            ws_acc.cell(row=r_idx, column=18).value = f"=B{r_idx}&E{r_idx}"
            ws_acc.cell(row=r_idx, column=19).value = f"=G{r_idx}"
            ws_acc.cell(row=r_idx, column=20).value = f"=M{r_idx}"

    logging.info(f"Processing finished. Exceptions: {exceptions_moved}")
    wb.save(wip_path)
    final_path = wip_path.replace("-wip.xlsx", ".xlsx")
    if os.path.exists(final_path): os.remove(final_path)
    os.rename(wip_path, final_path)
    logging.info(f"COMPLETED. File: {final_path}")

if __name__ == "__main__":
    main()
