import argparse
import csv
import logging
from pathlib import Path
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%H:%M:%S",
)

MISMATCH_OUTPUT_THRESHOLD = 0
DEFAULT_ENCODINGS = ("utf-8-sig", "utf-8", "gb18030")
DIFF_VALUE_SEPARATOR = "; "


def parse_args():
    parser = argparse.ArgumentParser(
        description="Compare two CSV files using exact string matching."
    )
    parser.add_argument("source_file", help="Path to source CSV file")
    parser.add_argument("target_file", help="Path to target CSV file")
    parser.add_argument(
        "key_col",
        help="Unique key column, either Excel-style letter (e.g. C) or header name",
    )
    parser.add_argument(
        "--decimal-places",
        type=int,
        default=5,
        help="Decimal places for comparing non-key numeric values (default: 5)",
    )
    return parser.parse_args()


def resolve_key_column(fieldnames, key_col):
    if not fieldnames:
        raise ValueError("CSV file has no header row")

    key_col = key_col.strip()
    if not key_col:
        raise ValueError("Key column cannot be blank")

    upper_key = key_col.upper()
    if upper_key.isalpha():
        col_idx = column_letter_to_index(upper_key)
        if col_idx < len(fieldnames):
            return fieldnames[col_idx]

    for fieldname in fieldnames:
        if fieldname == key_col:
            return fieldname

    raise ValueError(f"Unable to resolve key column: {key_col}")


def column_letter_to_index(column_letter):
    index = 0
    for char in column_letter:
        if not ("A" <= char <= "Z"):
            raise ValueError(f"Invalid column letter: {column_letter}")
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def column_index_to_letter(column_index):
    index = column_index
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def load_csv_rows(path):
    last_error = None
    for encoding in DEFAULT_ENCODINGS:
        try:
            with open(path, "r", encoding=encoding, newline="") as handle:
                reader = csv.DictReader(handle)
                if reader.fieldnames is None:
                    raise ValueError(f"{path} has no header row")

                fieldnames = [field.strip() if field else "" for field in reader.fieldnames]
                rows = []
                for row_number, raw_row in enumerate(reader, start=2):
                    normalized_row = {}
                    for field in fieldnames:
                        value = raw_row.get(field, "")
                        normalized_row[field] = value.strip() if value is not None else ""
                    normalized_row["_RowNumber"] = row_number
                    rows.append(normalized_row)

                return fieldnames, rows
        except UnicodeDecodeError as exc:
            last_error = exc

    raise UnicodeDecodeError(
        last_error.encoding if last_error else "unknown",
        last_error.object if last_error else b"",
        last_error.start if last_error else 0,
        last_error.end if last_error else 0,
        f"Unable to decode {path} with supported encodings: {DEFAULT_ENCODINGS}",
    )


def build_lookup(rows, key_name, label):
    filtered_rows = []
    blank_key_count = 0

    for row in rows:
        if row.get(key_name, "") == "":
            blank_key_count += 1
            continue
        filtered_rows.append(row)

    if blank_key_count:
        logging.info("Dropped %s rows with blank key from %s", blank_key_count, label)

    lookup = {}
    for row in filtered_rows:
        key = row[key_name]
        lookup.setdefault(key, row)

    return filtered_rows, lookup


def write_csv(path, fieldnames, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def normalize_decimal_string(value, decimal_places):
    quantizer = Decimal("1").scaleb(-decimal_places)
    normalized = value.quantize(quantizer, rounding=ROUND_HALF_UP)
    return format(normalized, "f")


def values_match(source_value, target_value, decimal_places):
    if source_value == target_value:
        return True

    if source_value == "" or target_value == "":
        return False

    try:
        source_decimal = Decimal(source_value)
        target_decimal = Decimal(target_value)
    except InvalidOperation:
        return False

    return normalize_decimal_string(source_decimal, decimal_places) == normalize_decimal_string(
        target_decimal,
        decimal_places,
    )


def trim_trailing_empty_values(values):
    trimmed = list(values)
    while trimmed and trimmed[-1][1] == "":
        trimmed.pop()
    return trimmed


def display_value(value):
    return value if value != "" else "<empty>"


def display_field_name(field_name, position):
    column_label = column_index_to_letter(position)
    if field_name:
        return f"{column_label}-{field_name}"
    return f"{column_label}-<blank-header-column-{position}>"


def compare_csv_files():
    args = parse_args()
    output_prefix = f"{Path(args.source_file).stem}_"

    logging.info("Loading Source: %s", args.source_file)
    source_fields, source_rows = load_csv_rows(args.source_file)

    logging.info("Loading Target: %s", args.target_file)
    target_fields, target_rows = load_csv_rows(args.target_file)

    source_key_name = resolve_key_column(source_fields, args.key_col)
    target_key_name = resolve_key_column(target_fields, args.key_col)

    source_rows, source_lookup = build_lookup(source_rows, source_key_name, "Source")
    target_rows, target_lookup = build_lookup(target_rows, target_key_name, "Target")

    source_keys = set(source_lookup)
    target_keys = set(target_lookup)

    in_source_not_in_target = [
        row for row in source_rows if row[source_key_name] not in target_keys
    ]
    in_target_not_in_source = [
        row for row in target_rows if row[target_key_name] not in source_keys
    ]

    in_source_not_in_target_file = f"{output_prefix}InSourceNotInTarget.csv"
    write_csv(in_source_not_in_target_file, source_fields, in_source_not_in_target)
    logging.info(
        "Saved %s (%s records)",
        in_source_not_in_target_file,
        len(in_source_not_in_target),
    )

    in_target_not_in_source_file = f"{output_prefix}InTargetNotInSource.csv"
    write_csv(in_target_not_in_source_file, target_fields, in_target_not_in_source)
    logging.info(
        "Saved %s (%s records)",
        in_target_not_in_source_file,
        len(in_target_not_in_source),
    )

    logging.info("Comparing overlapping keys...")
    mismatch_rows = []
    mismatch_fieldnames = [source_key_name, "DiffValues"]

    for key in source_keys & target_keys:
        source_row = source_lookup[key]
        target_row = target_lookup[key]
        diff_values = get_row_diff_values(
            source_row,
            target_row,
            source_fields,
            target_fields,
            source_key_name,
            target_key_name,
            args.decimal_places,
        )
        if not diff_values:
            continue

        mismatch_rows.append(
            {
                source_key_name: source_row[source_key_name],
                "DiffValues": DIFF_VALUE_SEPARATOR.join(diff_values),
            }
        )

        if MISMATCH_OUTPUT_THRESHOLD > 0 and len(mismatch_rows) >= MISMATCH_OUTPUT_THRESHOLD:
            logging.info(
                "Mismatch threshold reached (%s records). Stopping early.",
                MISMATCH_OUTPUT_THRESHOLD,
            )
            break

    mismatch_file = f"{output_prefix}InSourceAndInTarget.csv"
    if MISMATCH_OUTPUT_THRESHOLD <= 0 or len(mismatch_rows) >= MISMATCH_OUTPUT_THRESHOLD:
        write_csv(mismatch_file, mismatch_fieldnames, mismatch_rows)
        logging.info("Saved %s (%s records)", mismatch_file, len(mismatch_rows))
    else:
        write_csv(mismatch_file, mismatch_fieldnames, [])
        logging.info(
            "Mismatch count below threshold (%s). Writing header-only %s.",
            MISMATCH_OUTPUT_THRESHOLD,
            mismatch_file,
        )


def rows_match(
    source_row,
    target_row,
    source_fields,
    target_fields,
    source_key_name,
    target_key_name,
    decimal_places,
):
    return not get_row_diff_values(
        source_row,
        target_row,
        source_fields,
        target_fields,
        source_key_name,
        target_key_name,
        decimal_places,
    )


def get_row_diff_values(
    source_row,
    target_row,
    source_fields,
    target_fields,
    source_key_name,
    target_key_name,
    decimal_places,
):
    source_values = []
    for field_position, field in enumerate(source_fields, start=1):
        if field_position == source_fields.index(source_key_name) + 1:
            continue
        source_values.append((field, source_row.get(field, ""), field_position))

    target_values = []
    for field_position, field in enumerate(target_fields, start=1):
        if field_position == target_fields.index(target_key_name) + 1:
            continue
        target_values.append((field, target_row.get(field, ""), field_position))

    source_values = trim_trailing_empty_values(source_values)
    target_values = trim_trailing_empty_values(target_values)

    if len(source_values) != len(target_values):
        max_len = max(len(source_values), len(target_values))
        source_values.extend([("", "", len(source_values) + idx + 1) for idx in range(max_len - len(source_values))])
        target_values.extend([("", "", len(target_values) + idx + 1) for idx in range(max_len - len(target_values))])

    diff_values = []
    for (source_field, source_value, source_position), (target_field, target_value, target_position) in zip(source_values, target_values):
        if values_match(source_value, target_value, decimal_places):
            continue
        diff_field = display_field_name(source_field or target_field, max(source_position, target_position))
        diff_values.append(
            f"{diff_field}:{display_value(source_value)}|{display_value(target_value)}"
        )

    return diff_values


if __name__ == "__main__":
    compare_csv_files()
