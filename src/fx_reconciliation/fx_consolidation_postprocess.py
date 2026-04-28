"""
FX Consolidation Post-Processing Tool
-------------------------------------
Purpose:
    Rebuilds 数据透视表 and 1数透结果 from a completed FX workbook after
    the user has finished manual lookup inputs and saved the file.
"""

import argparse
import logging
import os
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import PatternFill


LOG_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
DATE_FORMAT = '%H:%M:%S'
SETTLEMENT_FLOW_INPUT_SHEET_NAME = '数据透视表'
SETTLEMENT_FLOW_OUTPUT_SHEET_NAME = '1数透结果'
CHANNEL_ORDER_SHEET_NAME = '渠道订单'
PIVOT_TOTAL_FILL = PatternFill(fill_type='solid', fgColor='FF483D8B')
PIVOT_HIGHLIGHT_FILL = PatternFill(fill_type='solid', fgColor='FFFFFF00')
SETTLEMENT_FLOW_INPUT_HEADERS = [
    '支付币种',
    '支付金额（扣除通道成本）',
    '打款币种',
    '清算币种',
    '清算金额',
    '打款币种与清算币种是否一致',
    '通道名称',
]
SETTLEMENT_FLOW_OUTPUT_HEADERS = [
    '支付币种',
    '打款币种',
    '清算币种',
    '求和项:支付金额（扣除通道成本）',
    '求和项:清算金额',
]


logging.basicConfig(
    level=logging.INFO,
    format=LOG_FORMAT,
    datefmt=DATE_FORMAT,
)


#######################################################
#  Common Utils
#######################################################
def configure_run_logging(log_path):
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    for handler in list(logger.handlers):
        if getattr(handler, '_fx_consolidation_postprocess_file_handler', False):
            logger.removeHandler(handler)
            handler.close()

    file_handler = logging.FileHandler(log_path, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter(LOG_FORMAT, datefmt=DATE_FORMAT))
    file_handler._fx_consolidation_postprocess_file_handler = True
    logger.addHandler(file_handler)


def normalize_cell_text(value):
    if value is None:
        return ""
    return str(value).strip()


def numeric_cell_value(value, default=0.0):
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if text == "":
        return default

    text = text.replace(",", "").replace("$", "")
    try:
        return float(text)
    except ValueError:
        return default


def to_excel_cell_value(value):
    if value is None:
        return None
    if isinstance(value, str) and value == "":
        return None
    return value


def normalize_comparable_value(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def write_matrix(ws, start_row, start_col, rows):
    for row_offset, row_values in enumerate(rows):
        for col_offset, value in enumerate(row_values):
            ws.cell(row=start_row + row_offset, column=start_col + col_offset).value = to_excel_cell_value(value)


def write_header_row(ws, start_col, headers, fill):
    for col_offset, header in enumerate(headers):
        cell = ws.cell(row=1, column=start_col + col_offset)
        cell.value = header
        cell.fill = fill


def get_last_data_row(ws, min_col, max_col, min_row=1):
    for row_number in range(ws.max_row, min_row - 1, -1):
        for col_number in range(min_col, max_col + 1):
            if normalize_cell_text(ws.cell(row=row_number, column=col_number).value) != "":
                return row_number
    return min_row - 1


def get_range_values(ws, start_row, end_row, start_col, end_col):
    values = []
    if end_row < start_row:
        return values

    for row_number in range(start_row, end_row + 1):
        row_values = []
        for col_number in range(start_col, end_col + 1):
            row_values.append(ws.cell(row=row_number, column=col_number).value)
        values.append(row_values)
    return values


#######################################################
#  数据透视表 Creation and Input Rebuild
#######################################################
def recreate_pivot_source_sheet(wb):
    if SETTLEMENT_FLOW_INPUT_SHEET_NAME in wb.sheetnames:
        sheet_index = wb.sheetnames.index(SETTLEMENT_FLOW_INPUT_SHEET_NAME)
        old_sheet = wb[SETTLEMENT_FLOW_INPUT_SHEET_NAME]
        wb.remove(old_sheet)
    else:
        sheet_index = len(wb.sheetnames)

    ws_pivot = wb.create_sheet(SETTLEMENT_FLOW_INPUT_SHEET_NAME, sheet_index)
    write_header_row(ws_pivot, 1, SETTLEMENT_FLOW_INPUT_HEADERS, PIVOT_HIGHLIGHT_FILL)
    write_header_row(ws_pivot, 11, SETTLEMENT_FLOW_OUTPUT_HEADERS, PIVOT_HIGHLIGHT_FILL)
    return ws_pivot


def build_pivot_source_rows(ws_channel_values):
    pivot_rows = []

    for row_number in range(2, ws_channel_values.max_row + 1):
        if normalize_cell_text(ws_channel_values.cell(row=row_number, column=1).value) == "":
            continue

        ao_value = normalize_cell_text(ws_channel_values.cell(row=row_number, column=41).value)
        if ao_value != "否":
            continue

        pivot_rows.append([
            ws_channel_values.cell(row=row_number, column=36).value,
            ws_channel_values.cell(row=row_number, column=37).value,
            ws_channel_values.cell(row=row_number, column=38).value,
            ws_channel_values.cell(row=row_number, column=39).value,
            ws_channel_values.cell(row=row_number, column=40).value,
            ws_channel_values.cell(row=row_number, column=41).value,
            ws_channel_values.cell(row=row_number, column=42).value,
        ])

    return pivot_rows


def validate_settlement_flow_input_sheet(ws_pivot, expected_rows):
    actual_last_row = get_last_data_row(ws_pivot, 1, 7, min_row=2)
    actual_rows = get_range_values(ws_pivot, 2, actual_last_row, 1, 7) if actual_last_row >= 2 else []
    normalized_actual = [[normalize_comparable_value(value) for value in row] for row in actual_rows]
    normalized_expected = [[normalize_comparable_value(value) for value in row] for row in expected_rows]

    if len(actual_rows) != len(expected_rows):
        raise ValueError(
            f"Settlement flow input validation failed: expected {len(expected_rows)} rows, "
            f"found {len(actual_rows)} rows in 数据透视表!A:G."
        )
    if normalized_actual != normalized_expected:
        raise ValueError(
            "Settlement flow input validation failed: 数据透视表!A:G values do not match "
            "the expected AJ:AP value mapping from 渠道订单."
        )


def rebuild_settlement_flow_input_sheet(write_wb, ws_channel_values):
    ws_pivot = recreate_pivot_source_sheet(write_wb)
    pivot_rows = build_pivot_source_rows(ws_channel_values)
    write_matrix(ws_pivot, 2, 1, pivot_rows)
    validate_settlement_flow_input_sheet(ws_pivot, pivot_rows)

    logging.info("Settlement flow input validated: rebuilt %s rows in 数据透视表!A:G", len(pivot_rows))
    return ws_pivot, pivot_rows


#######################################################
#  数据透视表 Summary Build
#######################################################
def build_grouped_pivot_rows(source_rows):
    grouped = {}
    for row in source_rows:
        group_key = (
            normalize_cell_text(row[0]),
            normalize_cell_text(row[2]),
            normalize_cell_text(row[3]),
        )
        bucket = grouped.setdefault(group_key, {'sum_b': 0.0, 'sum_e': 0.0})
        bucket['sum_b'] += numeric_cell_value(row[1])
        bucket['sum_e'] += numeric_cell_value(row[4])

    grouped_rows = []
    for group_key in sorted(grouped.keys()):
        totals = grouped[group_key]
        grouped_rows.append([
            group_key[0],
            group_key[1],
            group_key[2],
            totals['sum_b'],
            totals['sum_e'],
        ])

    grand_total_n = sum(row[3] for row in grouped_rows)
    grand_total_o = sum(row[4] for row in grouped_rows)
    grand_total_row = ["Grand Total", None, None, grand_total_n, grand_total_o]
    return grouped_rows, grand_total_row


def validate_settlement_flow_summary(ws_pivot, grouped_rows, grand_total_row):
    actual_last_row = get_last_data_row(ws_pivot, 11, 15, min_row=2)
    actual_rows = get_range_values(ws_pivot, 2, actual_last_row, 11, 15) if actual_last_row >= 2 else []
    expected_rows = grouped_rows + [grand_total_row]
    normalized_actual = [[normalize_comparable_value(value) for value in row] for row in actual_rows]
    normalized_expected = [[normalize_comparable_value(value) for value in row] for row in expected_rows]

    if normalized_actual != normalized_expected:
        raise ValueError("Settlement flow summary validation failed: 数据透视表!K:O does not match the grouped summary output.")

    grand_total_matches = [idx for idx, row in enumerate(actual_rows, start=2) if normalize_cell_text(row[0]) == "Grand Total"]
    if len(grand_total_matches) != 1:
        raise ValueError("Settlement flow summary validation failed: expected exactly one Grand Total row in 数据透视表!K:O.")

    grand_total_row_number = grand_total_matches[0]
    for col_number in range(11, 16):
        cell = ws_pivot.cell(row=grand_total_row_number, column=col_number)
        if cell.fill.fill_type != PIVOT_TOTAL_FILL.fill_type or cell.fill.fgColor.rgb != PIVOT_TOTAL_FILL.fgColor.rgb:
            raise ValueError("Settlement flow summary validation failed: Grand Total fill was not applied correctly to K:O.")


def build_settlement_flow_summary(ws_pivot, source_rows):
    grouped_rows, grand_total_row = build_grouped_pivot_rows(source_rows)

    for row_number in range(2, ws_pivot.max_row + 1):
        for col_number in range(11, 16):
            ws_pivot.cell(row=row_number, column=col_number).value = None
            ws_pivot.cell(row=row_number, column=col_number).fill = PatternFill(fill_type=None)

    output_rows = grouped_rows + [grand_total_row]
    write_matrix(ws_pivot, 2, 11, output_rows)

    grand_total_row_number = len(grouped_rows) + 2
    for col_number in range(11, 16):
        ws_pivot.cell(row=grand_total_row_number, column=col_number).fill = PIVOT_TOTAL_FILL

    validate_settlement_flow_summary(ws_pivot, grouped_rows, grand_total_row)
    logging.info(
        "Settlement flow summary validated: wrote %s grouped rows plus Grand Total to 数据透视表!K:O",
        len(grouped_rows),
    )


#######################################################
#  1数透结果 Publish
#######################################################
def validate_settlement_flow_results(ws_result, source_rows, remapped_rows, highlighted_row_count):
    actual_ae_last_row = get_last_data_row(ws_result, 1, 5, min_row=1)
    actual_ae_rows = get_range_values(ws_result, 1, actual_ae_last_row, 1, 5) if actual_ae_last_row >= 1 else []
    normalized_actual_ae = [[normalize_comparable_value(value) for value in row] for row in actual_ae_rows]
    normalized_expected_ae = [[normalize_comparable_value(value) for value in row] for row in source_rows]
    if normalized_actual_ae != normalized_expected_ae:
        raise ValueError("Settlement flow result validation failed: 1数透结果!A:E does not match 数据透视表!K:O.")

    if highlighted_row_count > 1:
        raise ValueError("Settlement flow result validation failed: more than one CNY/USD/CNY row was highlighted.")

    actual_hl_last_row = get_last_data_row(ws_result, 8, 12, min_row=1)
    actual_hl_rows = get_range_values(ws_result, 1, actual_hl_last_row, 8, 12) if actual_hl_last_row >= 1 else []
    normalized_actual_hl = [[normalize_comparable_value(value) for value in row] for row in actual_hl_rows]
    normalized_expected_hl = [[normalize_comparable_value(value) for value in row] for row in remapped_rows]
    if normalized_actual_hl != normalized_expected_hl:
        raise ValueError("Settlement flow result validation failed: 1数透结果!H:L does not match the filtered/remapped copy.")


def publish_settlement_flow_results(write_wb, ws_pivot):
    if SETTLEMENT_FLOW_OUTPUT_SHEET_NAME not in write_wb.sheetnames:
        raise KeyError(f"Target sheet not found in workbook: {SETTLEMENT_FLOW_OUTPUT_SHEET_NAME}")

    ws_result = write_wb[SETTLEMENT_FLOW_OUTPUT_SHEET_NAME]
    source_last_row = get_last_data_row(ws_pivot, 11, 15, min_row=1)
    source_rows = get_range_values(ws_pivot, 1, source_last_row, 11, 15) if source_last_row >= 1 else []

    if ws_result.max_row > 0:
        ws_result.delete_rows(1, ws_result.max_row)

    write_matrix(ws_result, 1, 1, source_rows)

    highlighted_row_count = 0
    remapped_rows = []
    for row_index, row_values in enumerate(source_rows, start=1):
        col_a = normalize_cell_text(row_values[0]) if len(row_values) > 0 else ""
        col_b = normalize_cell_text(row_values[1]) if len(row_values) > 1 else ""
        col_c = normalize_cell_text(row_values[2]) if len(row_values) > 2 else ""

        is_grand_total = col_a == "Grand Total"
        is_target_highlight = col_a == "CNY" and col_b == "USD" and col_c == "CNY"

        if is_target_highlight:
            highlighted_row_count += 1
            for col_number in range(1, 6):
                ws_result.cell(row=row_index, column=col_number).fill = PIVOT_HIGHLIGHT_FILL

        if row_index == 1 or (not is_target_highlight and not is_grand_total):
            remapped_rows.append([
                row_values[0] if len(row_values) > 0 else None,
                row_values[3] if len(row_values) > 3 else None,
                row_values[1] if len(row_values) > 1 else None,
                row_values[2] if len(row_values) > 2 else None,
                row_values[4] if len(row_values) > 4 else None,
            ])

    write_matrix(ws_result, 1, 8, remapped_rows)
    validate_settlement_flow_results(ws_result, source_rows, remapped_rows, highlighted_row_count)
    logging.info(
        "Settlement flow results validated: copied %s rows to 1数透结果!A:E and %s rows to 1数透结果!H:L",
        len(source_rows),
        len(remapped_rows),
    )


#######################################################
#  Main Orchestration
#######################################################
def run_fx_consolidation_postprocess(workbook_path, log_path=None):
    workbook_path = os.path.abspath(workbook_path)
    if not os.path.isfile(workbook_path):
        raise FileNotFoundError(f"Input Excel file not found: {workbook_path}")

    if log_path is None:
        timestamp = datetime.now().strftime('%Y%m%d-%H%M%S')
        workbook_dir = os.path.dirname(workbook_path)
        log_path = os.path.join(workbook_dir, f'fx_consolidation_postprocess_{timestamp}.log')

    configure_run_logging(log_path)
    logging.info("Starting FX consolidation post-processing...")
    logging.info("Input workbook: %s", workbook_path)

    # Load the completed workbook for value reads and formula-preserving writes.
    read_wb = load_workbook(workbook_path, data_only=True)
    write_wb = load_workbook(workbook_path)

    try:
        if CHANNEL_ORDER_SHEET_NAME not in read_wb.sheetnames:
            raise KeyError(f"Source sheet not found in workbook: {CHANNEL_ORDER_SHEET_NAME}")

        ws_channel_values = read_wb[CHANNEL_ORDER_SHEET_NAME]

        # Rebuild 数据透视表 input rows from 渠道订单.
        ws_pivot, source_rows = rebuild_settlement_flow_input_sheet(write_wb, ws_channel_values)

        # Build the grouped settlement-flow summary in 数据透视表.
        build_settlement_flow_summary(ws_pivot, source_rows)

        # Publish the final settlement-flow outputs to 1数透结果.
        publish_settlement_flow_results(write_wb, ws_pivot)

        # Save the updated workbook.
        write_wb.save(workbook_path)
        logging.info("Completed FX consolidation post-processing: %s", workbook_path)
    finally:
        read_wb.close()
        write_wb.close()

    return {
        'workbook_path': workbook_path,
        'log_path': log_path,
    }


def main():
    parser = argparse.ArgumentParser(description="Rebuild 数据透视表 and 1数透结果 for a specified FX workbook.")
    parser.add_argument("workbook", help="Path to the completed Excel workbook.")
    args = parser.parse_args()

    try:
        run_fx_consolidation_postprocess(args.workbook)
    except Exception as exc:
        logging.error("FX consolidation post-processing failed: %s", exc)
        raise


if __name__ == "__main__":
    main()
