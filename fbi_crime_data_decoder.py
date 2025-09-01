#!/usr/bin/env python3
"""
ASR Master File Decoder

uses simple variables at the top of the file to set the input/output paths
Edit INPUT_FILE and OUTPUT_FILE as needed, then run (lines 17 and 18)

if you get errors that mention any modules from lines 10-13, youll have to import them
thonny has a package manager in a menu somewhere that'll let you import them
otherwise, you can run "pip install ___" in the terminal to install any libraries
that may not be an issue
"""

from pathlib import Path
import pandas as pd
import sys
import logging

# ------------------ USER CONFIGURATION ------------------------------------
# Set these two variables to the path of your ASR master file and the desired output Excel file.
INPUT_FILE = Path(r'C:\Users\lovey\Downloads\asr-2024\asr-2024\2024_ASR1MON_NATIONAL_MASTER_FILE.txt')   # <-- change this to your input file path
OUTPUT_FILE = Path(r'C:\Users\lovey\OneDrive\Desktop\decoded_fbi_data.xlsx')  # <-- change this if you want a different output path
RECORD_LENGTH = 564  # expected fixed width record length (adjust if needed)
HEADER_OFFENSE_CODE = '000'  # offense code indicating a header record (adjust if data differs)

# ------------------ Logging ------------------------------------------------
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

# ------------------ Helpers ------------------------------------------------

def slice1(line: str, a: int, b: int) -> str:
    """Return substring using 1-indexed inclusive positions [a..b]."""
    if len(line) < b:
        # pad with spaces so slicing won't fail
        line = line.ljust(b)
    return line[a-1:b]


def safe_int(s: str) -> int:
    s = s.strip()
    if not s:
        return 0
    try:
        return int(s)
    except ValueError:
        filtered = ''.join(ch for ch in s if ch.isdigit())
        return int(filtered) if filtered else 0

# ------------------ Parsers -----------------------------------------------

def parse_asr_detail_record(line: str) -> dict:
    if len(line) < RECORD_LENGTH:
        line = line.ljust(RECORD_LENGTH)

    rec = {}
    rec['identifier'] = slice1(line, 1, 1)
    rec['state_code'] = slice1(line, 2, 3)
    rec['ori_code'] = slice1(line, 4, 10).strip()
    rec['group'] = slice1(line, 11, 12)
    rec['division'] = slice1(line, 13, 13)
    rec['year'] = slice1(line, 14, 15)
    rec['msa'] = slice1(line, 16, 18).strip()
    rec['card1_indicator_adult_male'] = slice1(line, 19, 19)
    rec['card2_indicator_adult_female'] = slice1(line, 20, 20)
    rec['card3_indicator_juvenile'] = slice1(line, 21, 21)
    rec['adjustment'] = slice1(line, 22, 22)
    rec['offense_code'] = slice1(line, 23, 25)

    male_groups = [
        'male_under_10','male_10_12','male_13_14','male_15','male_16','male_17','male_18','male_19',
        'male_20','male_21','male_22','male_23','male_24','male_25_29','male_30_34','male_35_39',
        'male_40_44','male_45_49','male_50_54','male_55_59','male_60_64','male_over_64'
    ]
    start = 41
    for name in male_groups:
        val = slice1(line, start, start+8)
        rec[name] = safe_int(val)
        start += 9

    female_groups = [g.replace('male', 'female') for g in male_groups]
    start = 239
    for name in female_groups:
        val = slice1(line, start, start+8)
        rec[name] = safe_int(val)
        start += 9

    juvenile_labels = ['juvenile_white','juvenile_black','juvenile_indian','juvenile_asian','juvenile_hispanic','juvenile_non_hispanic']
    start = 437
    for name in juvenile_labels:
        val = slice1(line, start, start+8)
        rec[name] = safe_int(val)
        start += 9

    adult_labels = ['adult_white','adult_black','adult_indian','adult_asian','adult_hispanic','adult_non_hispanic']
    start = 491
    for name in adult_labels:
        val = slice1(line, start, start+8)
        rec[name] = safe_int(val)
        start += 9

    return rec


def parse_asr_header_record(line: str) -> dict:
    if len(line) < RECORD_LENGTH:
        line = line.ljust(RECORD_LENGTH)
    rec = {}
    rec['identifier'] = slice1(line, 1, 1)
    rec['state_code'] = slice1(line, 2, 3)
    rec['ori_code'] = slice1(line, 4, 10).strip()
    rec['group'] = slice1(line, 11, 12)
    rec['division'] = slice1(line, 13, 13)
    rec['year'] = slice1(line, 14, 15)
    rec['raw_line'] = line
    return rec

# ------------------ File processing --------------------------------------

def process_file(input_path: Path):
    if not input_path.exists():
        logging.error('Input file does not exist: %s', input_path)
        raise FileNotFoundError(f'Input file does not exist: {input_path}')

    detail_rows = []
    header_rows = []
    total_lines = 0

    with input_path.open('r', encoding='utf-8', errors='replace') as fh:
        for raw in fh:
            total_lines += 1
            line = raw.rstrip('')
            if not line.strip():
                continue
            if len(line) < RECORD_LENGTH:
                line = line.ljust(RECORD_LENGTH)
            offense_code = slice1(line, 23, 25)
            if offense_code == HEADER_OFFENSE_CODE:
                header_rows.append(parse_asr_header_record(line))
            else:
                detail_rows.append(parse_asr_detail_record(line))

    logging.info('Read %d non-empty lines (details=%d, headers=%d)', total_lines, len(detail_rows), len(header_rows))
    return detail_rows, header_rows


def write_to_excel(detail_rows, header_rows, output_path: Path):
    df_details = pd.DataFrame(detail_rows) if detail_rows else pd.DataFrame()
    df_headers = pd.DataFrame(header_rows) if header_rows else pd.DataFrame()

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        if not df_details.empty:
            df_details.to_excel(writer, sheet_name='DETAILS', index=False)
        if not df_headers.empty:
            df_headers.to_excel(writer, sheet_name='HEADERS', index=False)

    logging.info('Wrote Excel workbook: %s', output_path)

# ------------------ Run --------------------------------------------------

if __name__ == '__main__':
    logging.info('Starting ASR decode using INPUT_FILE=%s OUTPUT_FILE=%s', INPUT_FILE, OUTPUT_FILE)
    details, headers = process_file(INPUT_FILE)
    write_to_excel(details, headers, OUTPUT_FILE)
    logging.info('Finished. Output written to %s', OUTPUT_FILE)
