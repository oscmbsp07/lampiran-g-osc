# app.py
# Lampiran G Generator (Padu v2) - agenda filter ikut FAIL INDUK + PTJ auto exclude
# Dependencies: streamlit, pandas, openpyxl, python-docx, pillow, pytesseract

import io
import re
import math
import zipfile
import datetime as dt
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple, Iterable

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image
import pytesseract


# =========================
# CONFIG / CONSTANTS
# =========================
APP_TITLE = "Generator Lampiran G - OSC MBSP (Padu v2)"

DAERAH_ORDER = {"SPU": 0, "SPS": 1, "SPT": 2}

ALLOWED_SHEETS = {
    "SERENTAK",
    "PKM",
    "TKR-GUNA",
    "TKR",
    "PKM TUKARGUNA",
    "BGN",
    "BGN EVCB",
    "EVCB",
    "EV",
    "TELCO",
    "PS",
    "SB",
    "CT",
    "PL",
    "KTUP",
    "JP",
    "LJUP",
}

# Kod canonical (untuk susun serentak)
CANON_CODES_ORDER = [
    "PKM",
    "TKR-GUNA",
    "TKR",
    "124A",
    "204D",
    "PS",
    "SB",
    "CT",
    "KTUP",
    "LJUP",
    "JP",
    "PL",
    "BGN",
    "EVCB",
    "EV",
    "TELCO",
]

KNOWN_CODES = set(CANON_CODES_ORDER)

PB_CODES = {"PKM", "TKR-GUNA", "TKR", "124A", "204D", "PS", "SB", "CT"}
KEJ_CODES = {"KTUP", "LJUP", "JP"}
JL_CODES = {"PL"}

# UT scope
UT_ALLOWED_SHEETS = {"SERENTAK", "PKM", "BGN", "BGN EVCB", "TKR-GUNA", "PKM TUKARGUNA"}
SERENTAK_UT_ALLOWED_INDUK = {"PB", "PKM", "BGN"}  # ikut logic asal

# Header hints & col candidates
HEADER_HINTS = [
    "No. Rujukan OSC",
    "No. Rujukan",
    "Rujukan OSC",
    "Pemaju",
    "Pemohon",
    "Daerah",
    "Mukim",
    "Seks",
    "Lot",
    "Tempoh Untuk Proses",
    "Tempoh Untuk Diberi",
    "Tarikh Keputusan",
]

COL_CANDIDATES = {
    "fail_no": ["norujukanosc", "no rujukan osc", "rujukan osc", "fail no", "failno", "no rujukan"],
    "pemohon": ["pemajupemohon", "pemaju/pemohon", "pemaju", "pemohon", "tetuan"],
    "mukim": ["mukimseksyen", "mukim/seksyen", "mukim", "seksyen"],
    "lot": ["lot"],
    "km": ["tempohuntukprosesolehjabataninduk", "tempohuntukproses"],
    "ut": ["tempohuntukdiberiulasanolehjabatanteknikal", "tempohuntukdiberiulasan"],
    "belum": ["jabatanindukteknikalygbelummemberikeputusanulasansehinggakini", "belummemberikeputusanulasan", "belummemberikeputusan", "belummemberi"],
    "keputusan": ["tarikhkeputusankuasa", "tarikhkeputusan"],
}

# Agenda parsing
AGENDA_BLOCK_SPLIT = re.compile(r"(KERTAS\s+MESYUARAT\s+BIL\.[^\n\r]+)", re.IGNORECASE)


# =========================
# UTIL
# =========================
def is_nan(v) -> bool:
    return v is None or (isinstance(v, float) and math.isnan(v)) or (isinstance(v, str) and v.strip().lower() == "nan")


def clean_str(v) -> str:
    if is_nan(v):
        return ""
    return str(v).strip()


def clean_fail_no(v) -> str:
    if is_nan(v):
        return ""
    s = str(v)
    s = re.sub(r"[\s\r\n\t]+", "", s)
    return s.strip()


def norm_basic(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.strip().lower()
    s = re.sub(r"[\s\r\n\t]+", " ", s)
    s = re.sub(r"[^a-z0-9]+", "", s)
    return s


def sheet_norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip()).upper()


def parse_date_from_cell(val) -> Optional[dt.date]:
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return None
    if isinstance(val, dt.datetime):
        return val.date()
    if isinstance(val, dt.date):
        return val

    # excel serial
    if isinstance(val, (int, float)) and 20000 < float(val) < 60000:
        base = dt.date(1899, 12, 30)
        return base + dt.timedelta(days=int(val))

    s = str(val).strip()
    if not s or s.lower() == "nan":
        return None

    m = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
    if m:
        d, mo, y = map(int, m.groups())
        try:
            return dt.date(y, mo, d)
        except Exception:
            return None

    m = re.search(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", s)
    if m:
        y, mo, d = map(int, m.groups())
        try:
            return dt.date(y, mo, d)
        except Exception:
            return None

    return None


def is_blankish_text(v) -> bool:
    if v is None or is_nan(v):
        return True
    s = str(v).strip()
    if s == "":
        return True
    if s.lower() in {"-", "—", "–", "n/a", "na", "nil", "tiada"}:
        return True
    if re.fullmatch(r"[-–—\s]+", s):
        return True
    return False


def keputusan_is_empty(v) -> bool:
    if v is None or is_nan(v):
        return True
    s = str(v).strip()
    if s == "" or s.lower() in {"-", "tiada", "nil", "n/a", "na", "—", "–"}:
        return True
    if re.fullmatch(r"[-–—\s]+", s):
        return True
    if parse_date_from_cell(s) is not None:
        return False
    return False


def canon_serentak_codes(codes: Set[str]) -> List[str]:
    return [c for c in CANON_CODES_ORDER if c in codes]


def is_serentak(sheet_name: str, fail_no: str) -> bool:
    if sheet_norm(sheet_name) == "SERENTAK":
        return True
    return "SERENTAK" in str(fail_no or "").upper()


# =========================
# FAIL NO NORMALIZATION (MATCHING vs DISPLAY)
# =========================
def unify_prefix_for_match(s: str) -> str:
    """
    Untuk matching sahaja:
    - buang whitespace
    - unify MBPS/MPSP -> MBSP (kalau ada variasi legacy)
    """
    s = re.sub(r"[\s\r\n\t]+", "", str(s or "")).upper()
    s = re.sub(r"^(MPSP|MBPS)", "MBSP", s)
    return s


def norm_key_for_match(s: str) -> str:
    s = unify_prefix_for_match(s)
    s = re.sub(r"[-/\\()\[\]{}+.,:;]", "", s)
    return s.lower()


def split_fail_induk(fail_no: str) -> str:
    """
    Split FAIL INDUK dari FAIL NO.
    Penting: agenda filter wajib guna INDUK, bukan suffix jabatan.
    """
    s = str(fail_no or "").strip()
    if not s:
        return s
    s = re.sub(r"[\s\r\n\t]+", "", s)

    # Cari '-' terakhir yang menandakan mula suffix kod
    for i in range(len(s) - 1, 0, -1):
        if s[i] == "-":
            suffix = s[i + 1 :].upper()
            if any(code in suffix for code in KNOWN_CODES):
                return s[:i]
    return s


def prefix_key_from_induk_for_match(induk: str) -> str:
    """
    Prefix (untuk fallback typo): sampai slash terakhir.
    Contoh: MBSP/15/U24-2511/2540 -> MBSP/15/U24-2511/
    """
    s = unify_prefix_for_match(induk)
    if "/" in s:
        return s[: s.rfind("/") + 1]
    return s


def lot_tokens(s: str) -> Set[str]:
    s = str(s or "").replace("&", ",")
    return set(re.findall(r"\d+", s))


def simplify_pemohon_tokens(name: str) -> List[str]:
    """
    Token pemohon untuk fuzzy match:
    - buang tetuan/tuan/puan
    - buang sdn bhd/bhd/enterprise/plc dsb
    - return token alfabet/nombor yg relevan
    """
    s = (name or "").lower()
    s = re.sub(r"\b(tetuan|tuan|puan)\b", " ", s)
    s = re.sub(r"\b(sdn\.?\s*bhd\.?|sdn\s*bhd|bhd|berhad|enterprise|enterprises|plc|llp|ltd)\b", " ", s)
    toks = re.findall(r"[a-z0-9]+", s)
    # buang token terlalu pendek (kecuali angka besar)
    toks = [t for t in toks if len(t) >= 3 or t.isdigit()]
    return toks


def pemohon_similar(a: str, b: str) -> bool:
    """
    Fuzzy match pemohon (selamat) sebab akan digabung dengan prefix+lot.
    - True jika intersection token >= 2
    - atau substring panjang (>=8)
    """
    a0 = re.sub(r"[^a-z0-9]+", "", (a or "").lower())
    b0 = re.sub(r"[^a-z0-9]+", "", (b or "").lower())
    if a0 and b0:
        if (a0 in b0 and len(a0) >= 8) or (b0 in a0 and len(b0) >= 8):
            return True

    ta = set(simplify_pemohon_tokens(a))
    tb = set(simplify_pemohon_tokens(b))
    if len(ta & tb) >= 2:
        return True
    return False


def tidy_fail_no_display(raw: str, sheet_u: str) -> str:
    """
    DISPLAY:
    - buang whitespace
    - kekalkan prefix asal (MBSP/MBPS/MPSP) supaya output ikut data asal
    - buang nota tertentu yang memang “ganggu” (contoh JILID2, TCO)
    - (optional) ikut pattern Aidil: PKM(TG) dari sheet PKM kadang dibersihkan -> PKM
    """
    s = re.sub(r"[\s\r\n\t]+", "", str(raw or ""))

    # buang parentheses yg mengandungi JILID / TCO (match partial)
    def _paren_repl(m):
        inner = m.group(1)
        if re.search(r"(?i)jilid", inner) or re.search(r"(?i)\btco\b", inner):
            return ""
        return "(" + inner + ")"

    s = re.sub(r"\(([^)]*)\)", _paren_repl, s)

    # pattern: PKM(TG) dari sheet PKM -> buang (TG) (supaya sama gaya Aidil utk kes tertentu)
    if sheet_u == "PKM":
        s = s.replace("PKM(TG)", "PKM")

    # kemas ++
    s = re.sub(r"\+{2,}", "+", s)
    return s.strip()


# =========================
# CODE EXTRACTION (ELAK 204D PALS U DARI NOTA)
# =========================
def _normalize_fail_for_code_scan(fail_no: str) -> str:
    s = str(fail_no or "").upper()

    # Kes nota yang selalu menyebabkan 204D tersalah pickup:
    # 1) SB(204D) -> SB sahaja
    s = re.sub(r"\b(SB|CT|PS)\s*\(\s*204D\s*\)", r"\1", s)
    # 2) (SB)204D -> SB sahaja
    s = re.sub(r"\(\s*(SB|CT|PS)\s*\)\s*204D", r"\1", s)

    # Nota umum selepas kod (JP(JILID2), PKM(PIN), BGN(PIN), PKM(TG), dsb)
    # Simpan kod sahaja, buang isi kurungan
    s = re.sub(r"\b(PKM|BGN|JP|KTUP|LJUP|PS|SB|CT|PL|TKR|TKR-GUNA|EVCB|EV|TELCO)\s*\([^)]*\)", r"\1", s)

    return s


def extract_codes(fail_no: str, sheet_name: str) -> Set[str]:
    s = _normalize_fail_for_code_scan(fail_no)
    tokens = re.split(r"[\s\+\-/\\(),]+", s)
    codes = set()
    for t in tokens:
        t = t.strip()
        if t in KNOWN_CODES:
            codes.add(t)

    sn = sheet_norm(sheet_name)
    if sn == "BGN EVCB":
        codes.update({"BGN", "EVCB"})
    elif sn in KNOWN_CODES:
        codes.add(sn)

    return codes


# =========================
# EXCEL READER
# =========================
def find_header_row(excel_bytes: bytes, sheet: str) -> Tuple[Optional[int], int]:
    raw = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=sheet, header=None, engine="openpyxl", nrows=80)
    best_idx, best_score = None, 0
    for i in range(len(raw)):
        row = raw.iloc[i].astype(str).fillna("")
        joined = " | ".join(row.tolist())
        score = sum(1 for h in HEADER_HINTS if h.lower() in joined.lower())
        if score > best_score:
            best_score, best_idx = score, i
    return best_idx, best_score


def detect_columns(df: pd.DataFrame) -> Dict[str, str]:
    norm_map = {col: norm_basic(col) for col in df.columns}
    found = {}
    for key, needles in COL_CANDIDATES.items():
        for needle in needles:
            for col, ncol in norm_map.items():
                if needle in ncol:
                    found[key] = col
                    break
            if key in found:
                break
    return found


def read_kertas_excel(excel_bytes: bytes, daerah_label: str) -> List[dict]:
    out = []
    xl = pd.ExcelFile(io.BytesIO(excel_bytes), engine="openpyxl")
    allowed_upper = {s.upper() for s in ALLOWED_SHEETS}

    for sheet in xl.sheet_names:
        if (sheet or "").strip().upper() not in allowed_upper:
            continue

        hdr_idx, score = find_header_row(excel_bytes, sheet)
        if hdr_idx is None or score == 0:
            continue

        df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=sheet, header=hdr_idx, engine="openpyxl").dropna(how="all")
        if df.empty:
            continue

        cols = detect_columns(df)
        if "fail_no" not in cols or "pemohon" not in cols:
            continue

        for _, row in df.iterrows():
            fail = row.get(cols["fail_no"])
            pem = row.get(cols["pemohon"])
            if (is_nan(fail) or str(fail).strip() == "") and (is_nan(pem) or str(pem).strip() == ""):
                continue

            km_raw = row.get(cols["km"]) if "km" in cols else None
            rec = {
                "daerah": daerah_label,
                "sheet": (sheet or "").strip(),
                "sheet_u": sheet_norm(sheet),
                "fail_no_raw": clean_fail_no(fail),
                "pemohon": clean_str(pem),
                "mukim": clean_str(row.get(cols["mukim"])) if "mukim" in cols else "",
                "lot": clean_str(row.get(cols["lot"])) if "lot" in cols else "",
                "km_date": parse_date_from_cell(km_raw) if "km" in cols else None,
                "ut_date": parse_date_from_cell(row.get(cols["ut"])) if "ut" in cols else None,
                "belum": clean_str(row.get(cols["belum"])) if "belum" in cols else "",
                "keputusan": clean_str(row.get(cols["keputusan"])) if "keputusan" in cols else "",
                "km_raw": km_raw,
            }

            rec["serentak"] = is_serentak(rec["sheet"], rec["fail_no_raw"])
            rec["fail_induk"] = split_fail_induk(rec["fail_no_raw"])
            rec["fail_induk_norm"] = norm_key_for_match(rec["fail_induk"])
            rec["codes"] = extract_codes(rec["fail_no_raw"], rec["sheet"])

            out.append(rec)

    return out


# =========================
# AGENDA PARSER (FAIL INDUK FILTER + PTJ AUTO EXCLUDE)
# =========================
@dataclass
class AgendaItem:
    kertas_code: str
    osc_raw: str
    induk_norm: str  # normalized fail_induk for match
    prefix_norm: str  # normalized prefix for typo match
    pemohon_raw: str
    lot_raw: str
    lot_toks: Set[str]


def docx_text_with_ocr(docx_bytes: bytes) -> str:
    """
    Extract text from docx paragraphs + tables.
    Jika sangat sedikit text (contoh: agenda scan), buat OCR pada imej dalam docx.
    """
    doc = Document(io.BytesIO(docx_bytes))
    chunks: List[str] = []

    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            chunks.append(t)

    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                ct = cell.text.strip()
                if ct:
                    chunks.append(ct)

    text = "\n".join(chunks).strip()

    # OCR fallback jika nampak “kosong”
    if len(text) < 200:
        try:
            with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as z:
                imgs = [f for f in z.namelist() if f.startswith("word/media/")]
                ocr_out = []
                for name in imgs:
                    data = z.read(name)
                    im = Image.open(io.BytesIO(data))
                    if im.mode != "RGB":
                        im = im.convert("RGB")
                    ocr_txt = pytesseract.image_to_string(im)
                    if ocr_txt.strip():
                        ocr_out.append(ocr_txt)
                if ocr_out:
                    text = (text + "\n" + "\n".join(ocr_out)).strip()
        except Exception:
            pass

    return text


def parse_agenda_blocks(full_text: str) -> List[Tuple[str, str]]:
    parts = AGENDA_BLOCK_SPLIT.split(full_text)
    blocks = []
    for i in range(1, len(parts), 2):
        head = parts[i].strip()
        body = parts[i + 1] if i + 1 < len(parts) else ""
        blocks.append((head, body))
    return blocks


def parse_agenda_item(head: str, body: str) -> AgendaItem:
    head_u = head.upper()

    # kertas code: OSC/PKM/..., OSC/BGN/..., OSC/PTJ/...
    m = re.search(r"OSC/([^/]+)/", head_u)
    kertas_code = m.group(1).strip() if m else ""

    # No Rujukan OSC
    m2 = re.search(r"No\.?\s*Rujukan\s*OSC\s*[:\t ]+\s*([^\n\r]+)", body, flags=re.IGNORECASE)
    osc_raw = m2.group(1).strip() if m2 else ""
    osc_raw = osc_raw.strip(" :\t")
    if osc_raw in {"-", "—", "–"}:
        osc_raw = ""

    # pemohon
    m3 = re.search(r"(Pemohon|Tetuan)\s*[:\t ]+\s*([^\n\r]+)", body, flags=re.IGNORECASE)
    pem_raw = m3.group(2).strip() if m3 else ""

    # lot
    lot_raw = ""
    m4 = re.search(r"Di\s+Atas[^\n\r]{0,120}?Lot\s*([^\n\r]+)", body, flags=re.IGNORECASE)
    if m4:
        lot_raw = m4.group(1).strip()
        lot_raw = re.split(r"\bMukim\b|\bDaerah\b|\bBandar\b", lot_raw, flags=re.IGNORECASE)[0].strip(" ,.;")

    lot_toks = lot_tokens(lot_raw)

    # build match keys
    if osc_raw:
        # matching guna fail_induk (bukan suffix)
        # display clean minimal utk split (buang whitespace sahaja; kekal kurungan)
        osc_disp = re.sub(r"[\s\r\n\t]+", "", osc_raw)
        induk = split_fail_induk(osc_disp)
        induk_norm = norm_key_for_match(induk)
        prefix_norm = norm_key_for_match(prefix_key_from_induk_for_match(induk))
    else:
        induk_norm = ""
        prefix_norm = ""

    return AgendaItem(
        kertas_code=kertas_code,
        osc_raw=osc_raw,
        induk_norm=induk_norm,
        prefix_norm=prefix_norm,
        pemohon_raw=pem_raw,
        lot_raw=lot_raw,
        lot_toks=lot_toks,
    )


class AgendaIndex:
    def __init__(self, items: List[AgendaItem]):
        self.items = items

        # PTJ auto exclude
        self.non_ptj = [it for it in items if it.kertas_code.upper() != "PTJ"]

        self.induk_set = {it.induk_norm for it in self.non_ptj if it.induk_norm}

        # index prefix -> list
        self.by_prefix: Dict[str, List[AgendaItem]] = {}
        self.by_noosc: Dict[str, List[AgendaItem]] = {}  # key by pemohon raw bucket (fuzzy later)

        for it in self.non_ptj:
            if it.induk_norm:
                self.by_prefix.setdefault(it.prefix_norm, []).append(it)
            else:
                # group by first significant token to reduce scan (still fuzzy)
                toks = simplify_pemohon_tokens(it.pemohon_raw)
                key = toks[0] if toks else ""
                self.by_noosc.setdefault(key, []).append(it)

    def should_filter_row(self, row: dict) -> bool:
        """
        Filter rule (ikut gaya Aidil):
        1) Exact match FAIL INDUK (normalized) -> filter
        2) Fallback typo: prefix sama + lot overlap + pemohon fuzzy -> filter
        3) Kes agenda tiada OSC: lot overlap + pemohon fuzzy -> filter
        PTJ: tidak trigger (auto exclude di index).
        """
        if row["fail_induk_norm"] in self.induk_set:
            return True

        # fallback prefix+lot+pemohon fuzzy
        pref_norm = norm_key_for_match(prefix_key_from_induk_for_match(row["fail_induk"]))
        lots = lot_tokens(row.get("lot", ""))
        pem = row.get("pemohon", "")

        # prefix candidates
        for it in self.by_prefix.get(pref_norm, []):
            if it.lot_toks and lots and (it.lot_toks & lots):
                if pemohon_similar(pem, it.pemohon_raw):
                    return True

        # no-osc blocks: reduce scan by first token
        rtoks = simplify_pemohon_tokens(pem)
        bucket = rtoks[0] if rtoks else ""
        for it in self.by_noosc.get(bucket, []):
            if it.lot_toks and lots and (it.lot_toks & lots):
                if pemohon_similar(pem, it.pemohon_raw):
                    return True

        return False


def parse_agenda_docx(docx_bytes: bytes) -> AgendaIndex:
    text = docx_text_with_ocr(docx_bytes)
    blocks = parse_agenda_blocks(text)

    items: List[AgendaItem] = []
    for head, body in blocks:
        try:
            items.append(parse_agenda_item(head, body))
        except Exception:
            continue

    return AgendaIndex(items)


# =========================
# CATEGORY LOGIC
# =========================
def tindakan_ut(belum_text: str) -> str:
    if is_blankish_text(belum_text):
        return ""
    raw = str(belum_text).strip()
    parts = [p.strip() for p in re.split(r"[,&/]+", raw) if p.strip()]

    internal_map = {
        "KEJ": "Pengarah Kejuruteraan",
        "PB": "Pengarah Perancang Bandar",
        "BGN": "Pengarah Bangunan",
        "COB": "Pengarah COB",
        "KES": "Pengarah Kesihatan",
        "PEN": "Pengarah Penilaian",
        "PBRN": "Pengarah Perbandaran",
        "LESEN": "Pengarah Pelesenan",
        "JL": "Pengarah Landskap",
    }

    internal, external = [], []
    for p in parts:
        key = re.sub(r"\s+", "", p.upper())
        if key in internal_map:
            internal.append(internal_map[key])
        else:
            external.append(p.upper())

    def dedup(seq: Iterable[str]) -> List[str]:
        seen = set()
        out = []
        for x in seq:
            if x not in seen:
                seen.add(x)
                out.append(x)
        return out

    return "\n".join(dedup(internal) + dedup(external)).strip()


def parse_induk_code_from_km_raw(val) -> str:
    if val is None or is_nan(val):
        return ""
    s = str(val).strip()
    toks = re.findall(r"[A-Z]{2,5}", s.upper())
    return toks[-1] if toks else ""


def sheet_is_ut_allowed(sheet_u: str) -> bool:
    if sheet_u in UT_ALLOWED_SHEETS:
        return True
    if "GUNA" in sheet_u and ("TKR" in sheet_u or "TUKAR" in sheet_u or sheet_u == "TG"):
        return True
    return False


def compute_serentak_fail_display(fail_induk: str, union_codes: Set[str], raw_candidates: List[dict]) -> str:
    """
    Serentak FAIL NO:
    - jika ada pattern khas (SB)204D, guna raw (supaya tak hilang nota)
    - kalau tiada, rebuild ikut canon order (macam output Aidil)
    """
    special = any(re.search(r"\((SB|CT|PS)\)\s*204D", (r.get("fail_no_raw") or ""), re.IGNORECASE) for r in raw_candidates)
    if special:
        # pilih row SERENTAK kalau ada, kalau tak pilih yg paling panjang
        raw_candidates = sorted(raw_candidates, key=lambda r: (0 if r.get("sheet_u") == "SERENTAK" else 1, -len(r.get("fail_no_raw") or "")))
        best = raw_candidates[0].get("fail_no_raw", "")
        return tidy_fail_no_display(best, "SERENTAK")

    codes_join = "+".join(canon_serentak_codes(union_codes))
    induk_disp = re.sub(r"[\s\r\n\t]+", "", fail_induk)
    return f"{induk_disp}-{codes_join}" if codes_join else induk_disp


def perkara_3lines(d: Optional[dt.date]) -> str:
    dd = d.strftime("%d.%m.%Y") if d else ""
    return f"Penyediaan Kertas\nMesyuarat Tamat Tempoh\n{dd}"


def build_categories(
    rows: List[dict],
    agenda_idx: Optional[AgendaIndex],
    km_start: dt.date,
    km_end: dt.date,
    ut_start: dt.date,
    ut_end: dt.date,
    ut_enabled: bool,
) -> Tuple[List[dict], List[dict], List[dict], List[dict], List[dict]]:
    # 1) drop yang ada keputusan
    rows = [r for r in rows if keputusan_is_empty(r.get("keputusan"))]

    # 2) agenda filter (FAIL INDUK), PTJ auto exclude dalam AgendaIndex
    if agenda_idx is not None:
        rows = [r for r in rows if not agenda_idx.should_filter_row(r)]

    # 3) induk_code for UT serentak constraint
    for r in rows:
        r["induk_code"] = parse_induk_code_from_km_raw(r.get("km_raw"))

    # 4) group by fail_induk_norm (bukan fail_no_raw)
    by: Dict[str, List[dict]] = {}
    for r in rows:
        by.setdefault(r["fail_induk_norm"], []).append(r)

    cat1, cat2, cat3, cat4, cat5 = [], [], [], [], []

    def make_rec(tindakan: str, base: dict, jenis: str, fail_no: str, perkara: str) -> dict:
        return {
            "tindakan": tindakan,
            "jenis": jenis,
            "fail_no": fail_no,
            "pemohon": base.get("pemohon", ""),
            "mukim": base.get("mukim", ""),
            "lot": base.get("lot", ""),
            "daerah": base.get("daerah", ""),
            "perkara": perkara,
        }

    for _, grp in by.items():
        is_ser = any(g.get("serentak") for g in grp)
        union_codes: Set[str] = set()
        km_dates = [g.get("km_date") for g in grp if g.get("km_date")]
        for g in grp:
            union_codes |= set(g.get("codes") or set())

        km_date = min(km_dates) if km_dates else None
        fail_induk = grp[0].get("fail_induk", "")
        fail_no_ser = compute_serentak_fail_display(fail_induk, union_codes, grp)

        jenis_ser = "+".join(canon_serentak_codes(union_codes))
        jenis_ser = (jenis_ser + " (Serentak)").strip() if jenis_ser else "(Serentak)"

        # ========= CAT 1 (PB/BGN) =========
        if is_ser and km_date and (km_start <= km_date <= km_end):
            if union_codes & (PB_CODES - {"PS", "SB", "CT"}):
                cat1.append(make_rec("Pengarah Perancang Bandar", grp[0], jenis_ser, fail_no_ser, perkara_3lines(km_date)))
            if union_codes & {"BGN", "EVCB", "EV", "TELCO"}:
                cat1.append(make_rec("Pengarah Bangunan", grp[0], jenis_ser, fail_no_ser, perkara_3lines(km_date)))

        if not is_ser:
            for g in grp:
                d = g.get("km_date")
                if not (d and (km_start <= d <= km_end)):
                    continue
                disp_fail = tidy_fail_no_display(g.get("fail_no_raw", ""), g.get("sheet_u", ""))

                if set(g.get("codes") or set()) & {"PKM", "TKR", "TKR-GUNA"}:
                    cat1.append(make_rec("Pengarah Perancang Bandar", g, g.get("sheet_u", ""), disp_fail, perkara_3lines(d)))

                if set(g.get("codes") or set()) & {"BGN", "EVCB", "EV", "TELCO"}:
                    cat1.append(make_rec("Pengarah Bangunan", g, g.get("sheet_u", ""), disp_fail, perkara_3lines(d)))

        # ========= CAT 2 (UT) =========
        if ut_enabled:
            for g in grp:
                sheet_u = g.get("sheet_u", "")
                if not sheet_is_ut_allowed(sheet_u):
                    continue
                d = g.get("ut_date")
                if not (d and (ut_start <= d <= ut_end)):
                    continue
                if is_blankish_text(g.get("belum")):
                    continue
                if sheet_u == "SERENTAK" and (g.get("induk_code") or "") not in SERENTAK_UT_ALLOWED_INDUK:
                    continue

                tindakan = tindakan_ut(g.get("belum", ""))
                if is_blankish_text(tindakan):
                    continue

                disp_fail = fail_no_ser if is_ser else tidy_fail_no_display(g.get("fail_no_raw", ""), sheet_u)
                disp_jenis = jenis_ser if is_ser else sheet_u
                cat2.append(
                    make_rec(
                        tindakan,
                        g,
                        disp_jenis,
                        disp_fail,
                        f"Ulasan teknikal belum dikemukakan. Tamat Tempoh {d.strftime('%d.%m.%Y')}.",
                    )
                )

        # ========= CAT 3/4/5 =========
        if is_ser and km_date and (km_start <= km_date <= km_end):
            if union_codes & KEJ_CODES:
                cat3.append(make_rec("Pengarah Kejuruteraan", grp[0], jenis_ser, fail_no_ser, perkara_3lines(km_date)))
            if union_codes & JL_CODES:
                cat4.append(make_rec("Pengarah Landskap", grp[0], jenis_ser, fail_no_ser, perkara_3lines(km_date)))
            if union_codes & {"124A", "204D"}:
                cat5.append(make_rec("Pengarah Perancang Bandar", grp[0], jenis_ser, fail_no_ser, perkara_3lines(km_date)))

        if not is_ser:
            for g in grp:
                d = g.get("km_date")
                if not (d and (km_start <= d <= km_end)):
                    continue
                sheet_u = g.get("sheet_u", "")
                disp_fail = tidy_fail_no_display(g.get("fail_no_raw", ""), sheet_u)

                if sheet_u in {"KTUP", "JP", "LJUP"}:
                    cat3.append(make_rec("Pengarah Kejuruteraan", g, sheet_u, disp_fail, perkara_3lines(d)))
                if sheet_u == "PL":
                    cat4.append(make_rec("Pengarah Landskap", g, sheet_u, disp_fail, perkara_3lines(d)))
                if sheet_u in {"PS", "SB", "CT"}:
                    cat5.append(make_rec("Pengarah Perancang Bandar", g, sheet_u, disp_fail, perkara_3lines(d)))

    # sort
    def sort_key(r):
        return (
            DAERAH_ORDER.get(r.get("daerah", ""), 99),
            r.get("tindakan", ""),
            r.get("fail_no", ""),
        )

    cat1.sort(key=sort_key)
    cat2.sort(key=sort_key)
    cat3.sort(key=sort_key)
    cat4.sort(key=sort_key)
    cat5.sort(key=sort_key)

    return cat1, cat2, cat3, cat4, cat5


# =========================
# WORD EXPORT
# =========================
def add_title(doc: Document, text: str):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(14)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def add_heading(doc: Document, text: str):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(11)


def make_table(doc: Document, rows: List[dict]):
    headers = [
        "BIL",
        "TINDAKAN",
        "JENIS PERMOHONAN",
        "FAIL NO",
        "PEMAJU/PEMOHON",
        "MUKIM/SEKSYEN",
        "LOT",
        "PERKARA",
    ]
    tbl = doc.add_table(rows=1, cols=len(headers))
    tbl.style = "Table Grid"
    hdr = tbl.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = h

    for i, r in enumerate(rows, start=1):
        cells = tbl.add_row().cells
        cells[0].text = str(i)
        cells[1].text = r.get("tindakan", "")
        cells[2].text = r.get("jenis", "")
        cells[3].text = r.get("fail_no", "")
        cells[4].text = r.get("pemohon", "")
        cells[5].text = r.get("mukim", "")
        cells[6].text = r.get("lot", "")
        cells[7].text = r.get("perkara", "")


def build_docx(cat1, cat2, cat3, cat4, cat5) -> bytes:
    doc = Document()
    add_title(doc, "LAMPIRAN G")

    add_heading(doc, "KATEGORI 1: PENYEDIAAN KERTAS MESYUARAT TAMAT TEMPOH (PB/BGN)")
    make_table(doc, cat1)
    doc.add_paragraph("")

    add_heading(doc, "KATEGORI 2: ULASAN TEKNIKAL BELUM DIKEMUKAKAN TAMAT TEMPOH")
    make_table(doc, cat2)
    doc.add_paragraph("")

    add_heading(doc, "KATEGORI 3: PENYEDIAAN KERTAS MESYUARAT TAMAT TEMPOH (KEJ)")
    make_table(doc, cat3)
    doc.add_paragraph("")

    add_heading(doc, "KATEGORI 4: PENYEDIAAN KERTAS MESYUARAT TAMAT TEMPOH (LANDSKAP)")
    make_table(doc, cat4)
    doc.add_paragraph("")

    add_heading(doc, "KATEGORI 5: PENYEDIAAN KERTAS MESYUARAT TAMAT TEMPOH (PS/SB/CT / 124A/204D SERENTAK)")
    make_table(doc, cat5)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# =========================
# STREAMLIT UI
# =========================
st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)

st.markdown(
    """
**Pembaikan utama (Padu v2):**
- **Agenda filter ikut FAIL INDUK** (bukan FAIL NO + suffix jabatan) → isu “dalam agenda tapi terlepas masuk” selesai.
- **PTJ auto dikecualikan** daripada agenda filtering (tiada toggle / tick).
- Robust normalisasi + fallback **prefix fail + lot + pemohon (fuzzy)** untuk kes typo / variasi pemohon.
- Serentak: rebuild FAIL NO ikut susunan kod canon (macam output standard), kecuali kes nota khas `(SB)204D`.
"""
)

colA, colB = st.columns(2)

with colA:
    km_start = st.date_input("Tempoh Kertas Mesyuarat (Mula)", value=dt.date(2026, 1, 8))
    km_end = st.date_input("Tempoh Kertas Mesyuarat (Tamat)", value=dt.date(2026, 1, 27))

with colB:
    ut_enabled = st.checkbox("Aktifkan Kategori 2 (Ulasan Teknikal Tamat Tempoh)", value=True)
    ut_start = st.date_input("Tempoh Ulasan Teknikal (Mula)", value=dt.date(2025, 12, 23))
    ut_end = st.date_input("Tempoh Ulasan Teknikal (Tamat)", value=dt.date(2026, 1, 12))

st.divider()

st.subheader("Muat naik fail")
c1, c2 = st.columns(2)
with c1:
    agenda_file = st.file_uploader("Agenda (DOCX)", type=["docx"])
with c2:
    kertas_files = st.file_uploader("Kertas Maklumat (Excel) - SPU/SPS/SPT", type=["xlsx"], accept_multiple_files=True)

if agenda_file and kertas_files:
    try:
        agenda_bytes = agenda_file.read()
        agenda_idx = parse_agenda_docx(agenda_bytes)

        # read excel
        all_rows: List[dict] = []
        for f in kertas_files:
            name = (f.name or "").upper()
            daerah = "SPU" if "SPU" in name else "SPS" if "SPS" in name else "SPT" if "SPT" in name else "UNKNOWN"
            all_rows += read_kertas_excel(f.read(), daerah)

        cat1, cat2, cat3, cat4, cat5 = build_categories(
            all_rows,
            agenda_idx,
            km_start,
            km_end,
            ut_start,
            ut_end,
            ut_enabled=ut_enabled,
        )

        st.success(
            f"Siap jana. Bilangan rekod: Cat1={len(cat1)}, Cat2={len(cat2)}, Cat3={len(cat3)}, Cat4={len(cat4)}, Cat5={len(cat5)}."
        )

        with st.expander("Debug ringkas (berapa yang ditapis oleh agenda)"):
            # kira anggaran
            rows_no_keputusan = [r for r in all_rows if keputusan_is_empty(r.get("keputusan"))]
            removed = sum(1 for r in rows_no_keputusan if agenda_idx.should_filter_row(r))
            st.write(f"Rekod (tiada keputusan): {len(rows_no_keputusan)}")
            st.write(f"Ditapis sebab agenda (FAIL INDUK / fallback): {removed}")
            st.write("Nota: PTJ tidak menapis (auto exclude).")

        out_docx = build_docx(cat1, cat2, cat3, cat4, cat5)
        st.download_button(
            "Muat turun Lampiran G (DOCX)",
            data=out_docx,
            file_name="Lampiran_G_Padu_v2.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as e:
        st.error(f"Ralat semasa proses: {e}")
else:
    st.info("Sila muat naik Agenda (DOCX) dan sekurang-kurangnya 1 fail Kertas Maklumat (xlsx).")
