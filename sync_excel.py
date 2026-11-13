import os
import json
import io
import re
import time as timer
import requests
import openpyxl
import gspread
from datetime import datetime, date, time
from decimal import Decimal
from openpyxl.utils.datetime import to_excel
from google.oauth2.service_account import Credentials

DROPBOX_URL_PLANTING = "https://www.dropbox.com/scl/fi/en98shfi3xvkbn8slaa5w/Database-for-Planting-Copy.xlsx?rlkey=1hncx3evuuz30haccwc9ph1wf&dl=1"
DROPBOX_URL_FOLLOWUP = "https://www.dropbox.com/scl/fi/av13w2gl8etjfs6ai6e0i/Potential-Farmer-for-ARR.xlsx?rlkey=qeg53qzba72yguf2nt2d2ek1t&dl=1"
DROPBOX_URL_CT_COMPILE = "https://www.dropbox.com/scl/fi/bg82woa41npj9sfc6l58h/CT-Compile.xlsx?rlkey=yu29wz7rjpgvzh7vttahexc0n&dl=1"
DROPBOX_URL_BT_COMPILE = "https://www.dropbox.com/scl/fi/5au2sa502w64h32in2azi/BT-Compile.xlsx?rlkey=8do0u62sevgy3vcovrvd4ptox&dl=1"
DROPBOX_URL_COMDEV = "https://www.dropbox.com/scl/fi/bkkqvxk9nnqkxr84qry5f/ComDev-Database_Update_20260818.xlsx?rlkey=t5691lpd5x0pax5bazvf9n8st&dl=1"
DROPBOX_URL_NURSERY = "https://www.dropbox.com/scl/fi/4wo13modlf34h2fdmpeus/New-Nursery-Database_.xlsx?rlkey=lgj3j98jq6fbnlq3bn1h6g5et&dl=1"
DROPBOX_URL_PLANTING_NEW = "https://www.dropbox.com/scl/fi/yw173h2zgnm6h3965i2ra/Database-for-Planting.xlsx?rlkey=9tfv2u3u2f12w641iauyoqc6w&dl=1"
DROPBOX_URL_HSE = "https://www.dropbox.com/scl/fi/g4gxslp755l98nxo5d5av/Dashboard-HSE-Final.xlsx?rlkey=4c1xl96yor90724jswlcsg3v1&dl=1"

TARGET_SPREADSHEET_ID = "1-LHCIUy1iD154WczY4Da_-iRaILtvL0mOW4gRW6YQac"
CAMERATRAP_SPREADSHEET_ID = "1Zkm82R3ct1HY1ov_HL35M9qEJudKPCkf0wGSNGxwdQ4"
BT_SPREADSHEET_ID = "1NJPa0nnULKqlihlkZo7J9M6gALi71LZ2_cs-0qz8SuI"
COMDEV_SPREADSHEET_ID = "1GB64g8bGV_IFjXcpRL7bu4lj9FtgRLKcUXqpNMzg50c"
NURSERY_SPREADSHEET_ID = "1or3IIjtpas5QD8PRAfO8WdhW6iTc-xH7v0B7hNOAF5M"
KTH_SPREADSHEET_ID = "1_8jR1z131_-9k3QVbsaGFRn8MY2JPPiEAgWOZliKo_k"
LOG_MONITORING_SPREADSHEET_ID = "1DRYnz-86Gx7mIzGaeNzUUd9xz_HoxwCgF4W-lLsOYjk"
PLANTING_NEW_SHEET_NAME = "ARR_SUMMARY"
PLANTING_NEW_SPREADSHEET_ID = "1CwnGELzoAz98rcVjrCPGoAu2FljK8FYKs3PEMf0OnFU"
HSE_SPREADSHEET_ID = "1ROa7ApqWNewNiI4wtV3pPHMdTDc02XrTypadPrkbM0s"

SUMMARY_HEADER = [
    "No", "Grid", "KTH", "Luas", "Tipe", "Petani", "Jenis Kelamin", "Desa",
    "ID_Petak", "Tgl_Partisipasi", "Tgl_Realisasi",
    "Agro_Reg", "Agro_BPO", "Temb_Reg", "Temb_MTL",
    "Total_Tanaman_Karbon", "link",
    "Insentif_T1", "Insentif_T2", "Insentif_T3", "Insentif_T4", "Insentif_T5",
    "Insentif_T6", "Insentif_T7", "Insentif_T8", "Insentif_T9", "Insentif_T10", "Insentif_T11",
    "Nominal_T1", "Nominal_T2", "Nominal_T3", "Nominal_T4", "Nominal_T5",
    "Nominal_T6", "Nominal_T7", "Nominal_T8", "Nominal_T9", "Nominal_T10", "Nominal_T11",
    "PupukPadat", "PupukCair", "tanggal_bpo", "insentifMTL", "tanggal_mtl",
    "dusun", "luas_agro", "luas_temb", "ENL", "ENP", "Total_Tanaman",
]
ARR_SUMMARY_DATE_COLS = ["J", "K", "AR"]

ARR_INSERTED_COLS = ("BW", "CX", "DX")
ARR_ASAM_E_COL = ("CX", "E-Asam Keranji")
ARR_ASAM_P_COL = ("DX", "P-Asam Keranji")
ARR_PUPUK_COL = ("OI", "Aplikasi Pupuk")
ARR_EXTRA_E_COLS = ("BY", "CU")
ARR_EXTRA_P_COLS = ("CX", "DT")

TARGET_CAPAIAN_MAPPING = {
    "Target Penanaman Tembawang (Ha)": "tembawang",
    "Target Penanaman Agroforestry (Ha)": "agroforestri",
    "Target Pengkayaan LOA": "LOA",
    "Target Pengkayaan PA": "PA",
}

FOLLOW_UP_HEADER = [
    "No", "Nama Petani", "Desa", "Dusun", "Lokasi", "Minat Program",
    "est_tot", "est_agr_reg", "est_tbw", "est_bpo", "est_mtl",
    "ukur_agr_reg", "ukur_tbw", "ukur_bpo", "ukur_mtl",
    "pot_agr_reg", "pot_tbw", "pot_bpo", "pot_mtl",
    "tgl_update", "status", "keterangan",
]
FOLLOW_UP_DATE_COLS = ["T"]


def _serial_number(value):
    number = to_excel(value)
    if float(number).is_integer():
        return int(number)
    return round(float(number), 6)


def clean_cell(value):
    if value is None:
        return ""
    if isinstance(value, (datetime, date)) and not isinstance(value, bool) and value.year < 1950:
        return _serial_number(value)
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, time):
        return value.strftime("%H:%M:%S")
    if isinstance(value, Decimal):
        return float(value)
    return value


def clean_rows(rows):
    return [[clean_cell(c) for c in row] for row in rows]


def safe_get(row, idx):
    if idx < 0 or idx >= len(row):
        return ""
    return row[idx]


def col_to_index(col_letters):
    result = 0
    for ch in col_letters:
        result = result * 26 + (ord(ch.upper()) - ord("A") + 1)
    return result - 1


WRITE_LIMIT_PER_MINUTE = 50
_write_times = []


def _throttle():
    now = timer.time()
    while _write_times and now - _write_times[0] >= 60:
        _write_times.pop(0)
    if len(_write_times) >= WRITE_LIMIT_PER_MINUTE:
        wait = 60 - (now - _write_times[0]) + 1
        if wait > 0:
            timer.sleep(wait)
    _write_times.append(timer.time())


def call_with_retry(func, *args, **kwargs):
    attempts = 6
    for attempt in range(attempts):
        _throttle()
        try:
            return func(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            status = getattr(getattr(e, "response", None), "status_code", None)
            retryable = status in (429, 500, 502, 503) or "429" in str(e)
            if not retryable or attempt == attempts - 1:
                raise
            timer.sleep(30 + attempt * 15)


def write_to_sheet(target_spreadsheet, sheet_name, data, date_cols=None):
    try:
        sheet = call_with_retry(target_spreadsheet.worksheet, sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        sheet = call_with_retry(target_spreadsheet.add_worksheet, title=sheet_name, rows=100, cols=20)

    call_with_retry(sheet.clear)
    if data:
        call_with_retry(sheet.update, range_name="A1", values=data, value_input_option="USER_ENTERED")

    target_rows = max(len(data), 1)
    target_cols = max((len(row) for row in data), default=1)
    target_cols = max(target_cols, 1)
    try:
        call_with_retry(sheet.resize, rows=target_rows, cols=target_cols)
    except Exception:
        pass

    if date_cols and len(data) >= 2:
        for col in date_cols:
            try:
                call_with_retry(sheet.format, f"{col}2:{col}{len(data)}", {
                    "numberFormat": {"type": "DATE", "pattern": "m/d/yyyy"}
                })
            except Exception:
                pass

    return sheet


def download_workbook(url, label):
    response = requests.get(url)
    response.raise_for_status()
    return openpyxl.load_workbook(io.BytesIO(response.content), data_only=True)


def _is_blank_cell(value):
    return value is None or (isinstance(value, str) and value.strip() == "")


def _is_blank_row(row):
    return all(_is_blank_cell(c) for c in row)


def read_sheet_rows(sheet, empty_row_gap=100):
    rows = []
    empty_streak = 0
    for row in sheet.iter_rows(values_only=True):
        row = list(row)
        if _is_blank_row(row):
            empty_streak += 1
            if empty_streak >= empty_row_gap:
                break
        else:
            empty_streak = 0
        rows.append(row)

    while rows and _is_blank_row(rows[-1]):
        rows.pop()

    if not rows:
        return rows

    max_cols = max(len(r) for r in rows)

    def col_is_blank(idx):
        return all(_is_blank_cell(r[idx]) if idx < len(r) else True for r in rows)

    last_col = max_cols
    while last_col > 0 and col_is_blank(last_col - 1):
        last_col -= 1

    return [row[:last_col] for row in rows]


def process_arr(source_workbook, target_spreadsheet):
    if "ARR" not in source_workbook.sheetnames:
        return

    raw_data = read_sheet_rows(source_workbook["ARR"])

    header_row_idx = next(
        (i for i, row in enumerate(raw_data)
         if all(k in [str(v).strip() if v is not None else "" for v in row]
                for k in ("No", "Kompartemen", "Nama KTH"))),
        None,
    )
    if header_row_idx is None:
        raise ValueError("Header tidak ditemukan di sheet ARR")

    data = raw_data[header_row_idx:]
    header = [str(v).strip() if v is not None else "" for v in data[0]]
    if "No" not in header:
        raise ValueError("Kolom 'No' tidak ditemukan di sheet ARR")
    no_col_index = header.index("No")

    data = [row[no_col_index:] for row in data]
    data = [row for row in data if any(str(c).strip() != "" for c in row if c is not None)]
    data = [["" if c == "#REF!" else c for c in row] for row in data]

    if len(data) >= 2:
        del data[1]

    write_to_sheet(target_spreadsheet, "ARR", clean_rows(data))

    asam_e_idx = col_to_index(ARR_ASAM_E_COL[0]) - no_col_index
    asam_p_idx = col_to_index(ARR_ASAM_P_COL[0]) - no_col_index
    asam_values = [(safe_get(r, asam_e_idx), safe_get(r, asam_p_idx)) for r in data]
    pupuk_idx = col_to_index(ARR_PUPUK_COL[0])
    pupuk_values = [safe_get(r, pupuk_idx) for r in data]

    inserted_idx = {col_to_index(c) - no_col_index for c in ARR_INSERTED_COLS}
    data = [[v for j, v in enumerate(r) if j not in inserted_idx] for r in data]

    header_row = data[0]
    e_start, e_end = col_to_index(ARR_EXTRA_E_COLS[0]), col_to_index(ARR_EXTRA_E_COLS[1])
    p_start, p_end = col_to_index(ARR_EXTRA_P_COLS[0]), col_to_index(ARR_EXTRA_P_COLS[1])
    extra_e_indices = list(range(e_start, e_end + 1))
    extra_p_indices = list(range(p_start, p_end + 1))
    extra_e_headers = [f"E-{safe_get(header_row, i)}" for i in extra_e_indices]
    extra_p_headers = [f"P-{safe_get(header_row, i)}" for i in extra_p_indices]

    summary = [
        SUMMARY_HEADER
        + extra_e_headers + [ARR_ASAM_E_COL[1]]
        + extra_p_headers + [ARR_ASAM_P_COL[1]]
        + [ARR_PUPUK_COL[1]]
    ]
    for row, (asam_e, asam_p), pupuk in zip(data[1:], asam_values[1:], pupuk_values[1:]):
        gender = "Laki-laki" if safe_get(row, 8) == 1 else "Perempuan" if safe_get(row, 9) == 1 else ""
        tipe = {"TBW": "Tembawang", "AGR": "Agroforestri"}.get(safe_get(row, 5), safe_get(row, 5))

        summary_row = [
            safe_get(row, 0), safe_get(row, 1), safe_get(row, 3), safe_get(row, 6),
            tipe, safe_get(row, 7), gender, safe_get(row, 11),
            safe_get(row, 13), safe_get(row, 14), safe_get(row, 15),
            safe_get(row, 21), safe_get(row, 22), safe_get(row, 23), safe_get(row, 24),
            safe_get(row, 48), safe_get(row, 368),
            safe_get(row, 334), safe_get(row, 337), safe_get(row, 340),
            safe_get(row, 343), safe_get(row, 346), safe_get(row, 349),
            safe_get(row, 352), safe_get(row, 355), safe_get(row, 358),
            safe_get(row, 361), safe_get(row, 364),
            safe_get(row, 336), safe_get(row, 339), safe_get(row, 342),
            safe_get(row, 345), safe_get(row, 348), safe_get(row, 351),
            safe_get(row, 354), safe_get(row, 357), safe_get(row, 360),
            safe_get(row, 363), safe_get(row, 366),
            safe_get(row, 39), safe_get(row, 40), safe_get(row, 41),
            safe_get(row, 43), safe_get(row, 44),
            safe_get(row, 12), safe_get(row, 16), safe_get(row, 17),
            safe_get(row, 18), safe_get(row, 19), safe_get(row, 47),
        ]
        summary_row += [safe_get(row, i) for i in extra_e_indices] + [asam_e]
        summary_row += [safe_get(row, i) for i in extra_p_indices] + [asam_p]
        summary_row.append(pupuk)
        summary.append(clean_rows([summary_row])[0])

    write_to_sheet(target_spreadsheet, "ARR_SUMMARY", summary, ARR_SUMMARY_DATE_COLS)


def process_planting_new(source_workbook, target_spreadsheet):
    sheet_name = "ARR"
    if sheet_name not in source_workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' tidak ditemukan. "
            f"Sheet yang tersedia: {source_workbook.sheetnames}"
        )

    raw_data = read_sheet_rows(source_workbook[sheet_name])
    if len(raw_data) < 8:
        raise ValueError(f"Sheet '{sheet_name}' tidak punya data mulai baris 8.")

    start, end = col_to_index("B"), col_to_index("EH")
    width = end - start + 1

    def pick(row):
        row = list(row) + [None] * (end + 1 - len(row))
        out = []
        for v in row[start:end + 1]:
            if isinstance(v, time):
                v = _serial_number(v)
            elif isinstance(v, datetime) and v.time() == time(0, 0):
                v = v.date()
            if v == "#REF!":
                v = ""
            out.append(v)
        return out

    header = pick(raw_data[5])
    rows = [pick(r) for r in raw_data[7:]]
    rows = [r for r in rows if not _is_blank_row(r)]
    if not rows:
        raise ValueError(f"Sheet '{sheet_name}' tidak punya data mulai baris 8.")

    write_to_sheet(target_spreadsheet, PLANTING_NEW_SHEET_NAME, clean_rows([header] + rows))


def process_target_capaian(source_workbook, target_spreadsheet):
    if "Target Capaian" not in source_workbook.sheetnames:
        return

    raw_data = read_sheet_rows(source_workbook["Target Capaian"])
    data = raw_data[1:]
    if not data:
        return

    for row in data[1:]:
        uraian = str(row[0]).strip() if row[0] is not None else ""
        if uraian in TARGET_CAPAIAN_MAPPING:
            row[0] = TARGET_CAPAIAN_MAPPING[uraian]

    write_to_sheet(target_spreadsheet, "TARGET_CAPAIAN", clean_rows(data))


def process_cumulative_area(source_workbook, target_spreadsheet):
    sheet_name = "Cumulative Area"
    if sheet_name not in source_workbook.sheetnames:
        return

    raw_data = read_sheet_rows(source_workbook[sheet_name])

    data = [row[1:] for row in raw_data[3:]]

    write_to_sheet(target_spreadsheet, "Planting Cumulative Area", clean_rows(data))


def process_log_monitoring(source_workbook, target_spreadsheet):
    sheet_name = "Log - Monitoring"
    if sheet_name not in source_workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' tidak ditemukan. "
            f"Sheet yang tersedia: {source_workbook.sheetnames}"
        )

    raw_data = read_sheet_rows(source_workbook[sheet_name])
    b, n = col_to_index("B"), col_to_index("N")

    rows = [row[b:n + 1] for row in raw_data]
    width = n - b + 1
    rows = [r + [None] * (width - len(r)) for r in rows]

    data = rows[3:4] + rows[5:]
    if len(data) < 2:
        raise ValueError(f"Sheet '{sheet_name}' tidak punya data untuk ditulis.")

    write_to_sheet(target_spreadsheet, "Sheet1", clean_rows(data))


def process_kth(source_workbook, target_spreadsheet):
    sheet_name = "KTH"
    if sheet_name not in source_workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' tidak ditemukan. "
            f"Sheet yang tersedia: {source_workbook.sheetnames}"
        )

    raw_data = read_sheet_rows(source_workbook[sheet_name])

    a_start, a_end = col_to_index("A"), col_to_index("I")
    n_start, n_end = col_to_index("N"), col_to_index("S")

    data_ai = [row[a_start:a_end + 1] for row in raw_data]
    data_nr = [row[n_start:n_end + 1] for row in raw_data]

    if not data_ai:
        raise ValueError(f"Sheet '{sheet_name}' kosong, tidak ada data untuk ditulis.")

    write_to_sheet(target_spreadsheet, "Sheet1", clean_rows(data_ai))
    write_to_sheet(target_spreadsheet, "Sheet2", clean_rows(data_nr))


CT_KEY_HEADERS = ("Year", "Batch", "Site")


def read_until_key_blank(sheet, key_headers, header_row=1, fallback_key_idx=(1, 2, 7)):
    rows = []
    key_idx = None
    for i, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        row = list(row)
        if i == header_row:
            names = [str(v).strip() if v is not None else "" for v in row]
            found = [names.index(k) for k in key_headers if k in names]
            key_idx = found if found else list(fallback_key_idx)
            rows.append(row)
            continue
        if i < header_row:
            rows.append(row)
            continue
        if key_idx is None:
            break
        if all(_is_blank_cell(safe_get(row, k)) for k in key_idx):
            break
        rows.append(row)

    if not rows:
        return rows

    max_cols = max(len(r) for r in rows)
    last_col = max_cols
    while last_col > 0 and all(
        _is_blank_cell(r[last_col - 1]) if last_col - 1 < len(r) else True for r in rows
    ):
        last_col -= 1
    return [row[:last_col] for row in rows]


def process_ct_compile(source_workbook, target_spreadsheet):
    sheet_name = "CT Compile"
    if sheet_name not in source_workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' tidak ditemukan. "
            f"Sheet yang tersedia: {source_workbook.sheetnames}"
        )

    data = read_until_key_blank(source_workbook[sheet_name], CT_KEY_HEADERS)
    if not data:
        raise ValueError(f"Sheet '{sheet_name}' kosong, tidak ada data untuk ditulis.")

    write_to_sheet(target_spreadsheet, "db-CT Compile", clean_rows(data))


def process_bt_compile(source_workbook, target_spreadsheet):
    sheet_name = "BT"
    if sheet_name not in source_workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' tidak ditemukan. "
            f"Sheet yang tersedia: {source_workbook.sheetnames}"
        )

    data = read_sheet_rows(source_workbook[sheet_name])
    if not data:
        raise ValueError(f"Sheet '{sheet_name}' kosong, tidak ada data untuk ditulis.")

    write_to_sheet(target_spreadsheet, "db-BT", clean_rows(data))


def process_comdev(source_workbook, target_spreadsheet):
    sheet_name = "ComDev"
    if sheet_name not in source_workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' tidak ditemukan. "
            f"Sheet yang tersedia: {source_workbook.sheetnames}"
        )

    raw_data = read_sheet_rows(source_workbook[sheet_name])

    data = [row[1:] for row in raw_data[3:]]
    if not data:
        raise ValueError(f"Sheet '{sheet_name}' kosong, tidak ada data untuk ditulis.")

    write_to_sheet(target_spreadsheet, "db-comdev", clean_rows(data))


COMDEV_SHEET2_NAME = "db-comdev-2"
COMDEV_ACTIVITY_SHEET_NAME = "db-comdev-kegiatan"
COMDEV_EXTRA_START_COL = "AW"
MONTH_NAMES = {
    "jan", "january", "januari", "feb", "february", "februari", "mar", "march", "maret",
    "apr", "april", "may", "mei", "jun", "june", "juni", "jul", "july", "juli",
    "aug", "august", "agu", "agt", "agustus", "sep", "sept", "september",
    "oct", "october", "okt", "oktober", "nov", "november", "dec", "december", "des", "desember",
}


def _is_numeric_only(value):
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float, Decimal)):
        return True
    if isinstance(value, (datetime, date, time)):
        return True
    text = str(value).strip().replace(",", ".")
    try:
        float(text)
        return True
    except ValueError:
        return False


def _is_month_only(value):
    return str(value).strip().rstrip(".").lower() in MONTH_NAMES


def _has_total_peserta(value):
    return re.search(r"total[\s_\-]*peserta", str(value), re.IGNORECASE) is not None


_PAREN_GROUP = re.compile(r"\([^()]*\)")


def clean_comdev_title(value):
    text = _PAREN_GROUP.sub(" ", str(value))
    return re.sub(r"\s+", " ", text).strip()


_DATE_FORMATS = ("%d %b %Y", "%d %B %Y", "%d-%b-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d %b %y")
_ID_MONTHS = {
    "januari": "january", "februari": "february", "maret": "march", "mei": "may",
    "juni": "june", "juli": "july", "agustus": "august", "agu": "aug", "agt": "aug",
    "oktober": "october", "okt": "oct", "desember": "december", "des": "dec",
}


def parse_comdev_date(value):
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    text = re.sub(r"\s+", " ", str(value)).strip()
    words = [_ID_MONTHS.get(w.lower(), w) for w in text.split(" ")]
    text2 = " ".join(words)
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text2, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return text


def _split_paren_segments(inner):
    return [seg.strip() for seg in re.split(r";|\s+-\s+", inner) if seg.strip()]


def _is_date_segment(segment):
    tokens = re.findall(r"[A-Za-z]+", segment)
    if any(t.lower() in MONTH_NAMES for t in tokens):
        return True
    if re.search(r"\b(?:19|20)\d{2}\b", segment):
        return True
    return re.fullmatch(r"[\d\s/.\-]+", segment) is not None


def parse_comdev_header(header_value, row7_value):
    raw = str(header_value)
    village, date_parts = "", []
    for group in _PAREN_GROUP.findall(raw):
        segments = _split_paren_segments(group[1:-1])
        if not segments:
            continue
        first = segments[0]
        if not village and not _is_date_segment(first):
            village = first
        date_parts.extend(seg for seg in segments if _is_date_segment(seg))
    title = clean_comdev_title(raw)
    if not _is_blank_cell(row7_value):
        date_value = parse_comdev_date(row7_value)
    elif date_parts:
        date_value = parse_comdev_date(" - ".join(date_parts))
    else:
        date_value = ""
    return title, village, date_value


def _abbr_title(title):
    words = [w for w in re.findall(r"[A-Za-z]+", title)]
    return "".join(w[0].upper() for w in words)[:8] or "KEG"


def _abbr_village(village):
    words = re.findall(r"[A-Za-z]+", village)
    if not words:
        return "XXX"
    if len(words) == 1:
        return words[0][:3].upper()
    return "".join(w[0].upper() for w in words)[:4]


def build_comdev_code(title, village, date_value, used):
    digits = re.sub(r"\D", "", date_value) if re.fullmatch(r"\d{4}-\d{2}-\d{2}", date_value) else "0000"
    code = f"{_abbr_title(title)}-{_abbr_village(village)}-{digits}"
    base, n = code, 1
    while code in used:
        n += 1
        code = f"{base}-{n}"
    used.add(code)
    return code


def comdev_extra_header_ok(header_value, row7_value):
    if _is_blank_cell(header_value) or _is_blank_cell(row7_value):
        return False
    if _is_numeric_only(header_value):
        return False
    if _is_month_only(header_value):
        return False
    if _has_total_peserta(header_value):
        return False
    return True


def process_comdev_sheet2(source_workbook, target_spreadsheet):
    sheet_name = "ComDev"
    if sheet_name not in source_workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' tidak ditemukan. "
            f"Sheet yang tersedia: {source_workbook.sheetnames}"
        )

    raw_data = read_sheet_rows(source_workbook[sheet_name])
    b, i = col_to_index("B"), col_to_index("I")
    n, u = col_to_index("N"), col_to_index("U")
    extra_start = col_to_index(COMDEV_EXTRA_START_COL)

    if len(raw_data) < 6:
        raise ValueError(f"Sheet '{sheet_name}' tidak punya baris header ke-6.")

    header_src = raw_data[5]
    row7_src = raw_data[6] if len(raw_data) >= 7 else []
    width = max(len(r) for r in raw_data)
    extra_idx = [
        c for c in range(extra_start, width)
        if comdev_extra_header_ok(safe_get(header_src, c), safe_get(row7_src, c))
        and clean_comdev_title(safe_get(header_src, c))
    ]

    def pick(row):
        row = list(row) + [None] * (max(u, width - 1) + 1 - len(row))
        return row[b:i + 1] + row[n:u + 1] + [row[c] for c in extra_idx]

    header = pick(header_src)
    base_len = len(header) - len(extra_idx)
    used = set()
    activities = []
    codes = []
    for c in extra_idx:
        title, village, date_value = parse_comdev_header(safe_get(header_src, c), safe_get(row7_src, c))
        code = build_comdev_code(title, village, date_value, used)
        codes.append(code)
        activities.append([code, title, village, date_value])

    header = header[:base_len] + codes
    data = [header] + [pick(r) for r in raw_data[8:]]
    if len(data) < 2:
        raise ValueError(f"Sheet '{sheet_name}' tidak punya data mulai baris 9.")

    write_to_sheet(target_spreadsheet, COMDEV_SHEET2_NAME, clean_rows(data))

    activity_rows = [["ID_Kode", "Judul_Kegiatan", "Desa", "Tanggal"]] + activities
    write_to_sheet(target_spreadsheet, COMDEV_ACTIVITY_SHEET_NAME, clean_rows(activity_rows))


def process_nursery_stock(source_workbook, target_spreadsheet):
    sheet_name = "Nursery Stock"
    if sheet_name not in source_workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' tidak ditemukan. "
            f"Sheet yang tersedia: {source_workbook.sheetnames}"
        )

    raw_data = read_sheet_rows(source_workbook[sheet_name])

    data = [row[1:] for row in raw_data[3:]]
    if not data:
        raise ValueError(f"Sheet '{sheet_name}' kosong, tidak ada data untuk ditulis.")

    write_to_sheet(target_spreadsheet, "Nursery Stock", clean_rows(data))


def process_log(source_workbook, target_spreadsheet):
    sheet_name = "Log"
    if sheet_name not in source_workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' tidak ditemukan. "
            f"Sheet yang tersedia: {source_workbook.sheetnames}"
        )

    raw_data = read_sheet_rows(source_workbook[sheet_name])

    data = [row[1:] for row in raw_data[3:]]

    if len(data) >= 2:
        del data[1]

    if not data:
        raise ValueError(f"Sheet '{sheet_name}' kosong, tidak ada data untuk ditulis.")

    write_to_sheet(target_spreadsheet, "Log", clean_rows(data))


def process_hse_sheet(source_workbook, target_spreadsheet, sheet_name, last_col):
    if sheet_name not in source_workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' tidak ditemukan. "
            f"Sheet yang tersedia: {source_workbook.sheetnames}"
        )

    raw_data = read_sheet_rows(source_workbook[sheet_name])
    end = col_to_index(last_col) + 1
    data = [row[:end] for row in raw_data[3:]]
    if not data:
        raise ValueError(f"Sheet '{sheet_name}' kosong, tidak ada data untuk ditulis.")

    write_to_sheet(target_spreadsheet, sheet_name, clean_rows(data))


def process_follow_up_petani(source_workbook, target_spreadsheet):
    if "Follow Up Petani" not in source_workbook.sheetnames:
        return

    data = read_sheet_rows(source_workbook["Follow Up Petani"])
    result = [FOLLOW_UP_HEADER]

    for row in data[3:]:
        if not (safe_get(row, 1) or safe_get(row, 2) or safe_get(row, 3)):
            continue
        result.append([safe_get(row, i) for i in range(1, 23)])

    write_to_sheet(target_spreadsheet, "FOLLOW_UP_PETANI", clean_rows(result), FOLLOW_UP_DATE_COLS)


def run_step(label, errors, func):
    start = timer.time()
    try:
        result = func()
        elapsed = timer.time() - start
        print(f"[OK] {label} ({elapsed:.1f}s)")
        return result
    except Exception as e:
        elapsed = timer.time() - start
        print(f"[GAGAL] {label} ({elapsed:.1f}s): {e}")
        errors.append((label, str(e)))
        return None


def main():
    errors = []
    run_start = timer.time()

    creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)
    target_spreadsheet = gc.open_by_key(TARGET_SPREADSHEET_ID)
    planting_new_spreadsheet = run_step("Buka spreadsheet Planting Baru", errors, lambda: gc.open_by_key(PLANTING_NEW_SPREADSHEET_ID))

    planting_wb = run_step(
        "Download Database-for-Planting", errors,
        lambda: download_workbook(DROPBOX_URL_PLANTING, "Database-for-Planting"),
    )
    if planting_wb:
        run_step("process_arr", errors, lambda: process_arr(planting_wb, target_spreadsheet))
        if planting_new_spreadsheet:
            run_step("process_target_capaian", errors, lambda: process_target_capaian(planting_wb, planting_new_spreadsheet))
            run_step("process_cumulative_area", errors, lambda: process_cumulative_area(planting_wb, planting_new_spreadsheet))

        log_monitoring_spreadsheet = run_step("Buka spreadsheet Log Monitoring", errors, lambda: gc.open_by_key(LOG_MONITORING_SPREADSHEET_ID))
        if log_monitoring_spreadsheet:
            run_step("process_log_monitoring", errors, lambda: process_log_monitoring(planting_wb, log_monitoring_spreadsheet))

        kth_spreadsheet = run_step("Buka spreadsheet KTH", errors, lambda: gc.open_by_key(KTH_SPREADSHEET_ID))
        if kth_spreadsheet:
            run_step("process_kth", errors, lambda: process_kth(planting_wb, kth_spreadsheet))

    planting_new_wb = run_step(
        "Download Database-for-Planting (baru)", errors,
        lambda: download_workbook(DROPBOX_URL_PLANTING_NEW, "Database-for-Planting (baru)"),
    )
    if planting_new_wb and planting_new_spreadsheet:
        run_step("process_planting_new", errors, lambda: process_planting_new(planting_new_wb, planting_new_spreadsheet))

    followup_wb = run_step(
        "Download Potential-Farmer-for-ARR", errors,
        lambda: download_workbook(DROPBOX_URL_FOLLOWUP, "Potential-Farmer-for-ARR"),
    )
    if followup_wb and planting_new_spreadsheet:
        run_step("process_follow_up_petani", errors, lambda: process_follow_up_petani(followup_wb, planting_new_spreadsheet))

    ct_compile_wb = run_step(
        "Download CT Compile", errors,
        lambda: download_workbook(DROPBOX_URL_CT_COMPILE, "CT Compile"),
    )
    if ct_compile_wb:
        cameratrap_spreadsheet = run_step("Buka spreadsheet cameratrap db", errors, lambda: gc.open_by_key(CAMERATRAP_SPREADSHEET_ID))
        if cameratrap_spreadsheet:
            run_step("process_ct_compile", errors, lambda: process_ct_compile(ct_compile_wb, cameratrap_spreadsheet))

    bt_compile_wb = run_step(
        "Download BT Compile", errors,
        lambda: download_workbook(DROPBOX_URL_BT_COMPILE, "BT Compile"),
    )
    if bt_compile_wb:
        bt_spreadsheet = run_step("Buka spreadsheet BT", errors, lambda: gc.open_by_key(BT_SPREADSHEET_ID))
        if bt_spreadsheet:
            run_step("process_bt_compile", errors, lambda: process_bt_compile(bt_compile_wb, bt_spreadsheet))

    comdev_wb = run_step(
        "Download ComDev Database", errors,
        lambda: download_workbook(DROPBOX_URL_COMDEV, "ComDev Database"),
    )
    if comdev_wb:
        comdev_spreadsheet = run_step("Buka spreadsheet ComDev", errors, lambda: gc.open_by_key(COMDEV_SPREADSHEET_ID))
        if comdev_spreadsheet:
            run_step("process_comdev", errors, lambda: process_comdev(comdev_wb, comdev_spreadsheet))
            run_step("process_comdev_sheet2", errors, lambda: process_comdev_sheet2(comdev_wb, comdev_spreadsheet))

    nursery_wb = run_step(
        "Download New Nursery Database", errors,
        lambda: download_workbook(DROPBOX_URL_NURSERY, "New Nursery Database"),
    )
    if nursery_wb:
        nursery_spreadsheet = run_step("Buka spreadsheet Nursery", errors, lambda: gc.open_by_key(NURSERY_SPREADSHEET_ID))
        if nursery_spreadsheet:
            run_step("process_nursery_stock", errors, lambda: process_nursery_stock(nursery_wb, nursery_spreadsheet))
            run_step("process_log", errors, lambda: process_log(nursery_wb, nursery_spreadsheet))

    hse_wb = run_step(
        "Download Dashboard HSE", errors,
        lambda: download_workbook(DROPBOX_URL_HSE, "Dashboard HSE"),
    )
    if hse_wb:
        hse_spreadsheet = run_step("Buka spreadsheet HSE", errors, lambda: gc.open_by_key(HSE_SPREADSHEET_ID))
        if hse_spreadsheet:
            run_step("process_hse_register_insiden", errors, lambda: process_hse_sheet(hse_wb, hse_spreadsheet, "Register Insiden", "AO"))
            run_step("process_hse_manhours", errors, lambda: process_hse_sheet(hse_wb, hse_spreadsheet, "Manhours", "Q"))

    total_elapsed = timer.time() - run_start
    print()
    if errors:
        print(f"Selesai dalam {total_elapsed:.1f}s dengan {len(errors)} bagian gagal (dilewati, bagian lain tetap berhasil ditulis):")
        for label, msg in errors:
            print(f" - {label}: {msg}")
    else:
        print(f"Selesai dalam {total_elapsed:.1f}s, semua bagian berhasil.")


if __name__ == "__main__":
    main()
