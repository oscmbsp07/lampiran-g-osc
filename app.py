import io
import math
import os
import re
import base64
import datetime as dt
import zipfile
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple, Set

import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from openpyxl import load_workbook

from docx import Document
from docx.enum.section import WD_ORIENTATION, WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt


# ============================================================
# CONFIG
# ============================================================
st.set_page_config(page_title="Lampiran G Unit OSC", layout="wide")

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

# Tapisan agenda TERHAD (ikut arahan terbaru user):
# ONLY: SERENTAK, PKM, TKR-GUNA/TG, PKM TUKARGUNA, BGN, TELCO, EVCB cluster (BGN EVCB/EVCB/EV)
AGENDA_FILTER_SHEETS = {
    "SERENTAK",
    "PKM",
    "TKR-GUNA",
    "PKM TUKARGUNA",
    "BGN",
    "TELCO",
    "EVCB",
    "BGN EVCB",
    "EV",
}

DAERAH_ORDER = {"SPU": 0, "SPS": 1, "SPT": 2}

KNOWN_CODES = [
    "PKM", "TKR-GUNA", "TKR", "124A", "204D", "PS", "SB", "CT",
    "KTUP", "LJUP", "JP", "PL",
    "BGN", "EVCB", "EV", "TELCO",
]

PB_CODES = {"PKM", "TKR-GUNA", "TKR", "124A", "204D", "PS", "SB", "CT"}
KEJ_CODES = {"KTUP", "LJUP", "JP"}
JL_CODES = {"PL"}

# UT rules (kekal)
UT_ALLOWED_SHEETS = {"SERENTAK", "PKM", "BGN", "BGN EVCB", "TKR-GUNA", "PKM TUKARGUNA", "TKR"}
SERENTAK_UT_ALLOWED_INDUK = {"PB", "PKM", "BGN"}


# ============================================================
# UI HELPERS (BACKGROUND + CSS)
# ============================================================
def _inject_bg_and_css(img_path: str) -> bool:
    try:
        with open(img_path, "rb") as f:
            data = f.read()
    except Exception:
        data = None

    b64 = base64.b64encode(data).decode("utf-8") if data else ""
    ext = os.path.splitext(img_path)[1].lower().replace(".", "")
    if ext in {"jpg", "jpeg"}:
        mime = "image/jpeg"
    elif ext == "png":
        mime = "image/png"
    else:
        mime = "image/*"

    bg_css = ""
    if data:
        bg_css = f"""
        .stApp::before {{
            content: "";
            position: fixed;
            inset: 0;
            z-index: -2;
            background-image:
                linear-gradient(rgba(0,0,0,0.48), rgba(0,0,0,0.48)),
                url("data:{mime};base64,{b64}");
            background-size: cover;
            background-position: center center;
            background-repeat: no-repeat;
            background-attachment: fixed;
            transform: translateZ(0);
        }}
        """

    css = f"""
    <style>
      html, body {{ height: 100%; }}
      body {{
        overflow-y: scroll;
        scrollbar-width: none;
        -ms-overflow-style: none;
      }}
      body::-webkit-scrollbar {{
        width: 0px;
        height: 0px;
        background: transparent;
      }}
      header, footer {{
        visibility: hidden;
        height: 0;
      }}
      .stApp {{
        background: transparent !important;
      }}
      {bg_css}
      section.main > div.block-container {{
        max-width: 1200px;
        padding-top: 0.8rem;
        padding-bottom: 0.8rem;
      }}
      .app-title {{
        text-align: center;
        font-weight: 900;
        letter-spacing: 1px;
        margin: 0.9rem 0 0.2rem 0;
        text-transform: uppercase;
        color: white;
        text-shadow: 0px 2px 14px rgba(0,0,0,0.55);
      }}
      .hero-spacer {{ height: 22vh; }}
      div[data-testid="stVerticalBlockBorderWrapper"] {{
        background: rgba(0,0,0,0.44) !important;
        border: 1px solid rgba(255,255,255,0.12) !important;
        border-radius: 18px !important;
        padding: 14px 16px 12px 16px !important;
        box-shadow: 0 10px 30px rgba(0,0,0,0.25);
        backdrop-filter: blur(2px);
      }}
      h1 a, h2 a, h3 a {{ display: none !important; }}
      label {{ font-size: 0.85rem !important; }}
      .stTextInput input {{ height: 2.35rem !important; }}
      div.stButton > button {{
        width: 100%;
        border-radius: 14px;
        font-weight: 800;
        letter-spacing: 0.8px;
      }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)
    return bool(data)


def _parse_ddmmyyyy(s: str) -> Optional[dt.date]:
    s = (s or "").strip()
    if not s:
        return None
    try:
        return dt.datetime.strptime(s, "%d/%m/%Y").date()
    except Exception:
        return None


# ============================================================
# UTIL - NORMALISASI & PARSING
# ============================================================
def is_nan(v) -> bool:
    return v is None or (isinstance(v, float) and math.isnan(v)) or (isinstance(v, str) and v.strip().lower() == "nan")


def clean_fail_no(v) -> str:
    if is_nan(v):
        return ""
    s = str(v)
    s = re.sub(r"[\s\r\n\t]+", "", s)
    return s.strip()


def clean_str(v) -> str:
    if is_nan(v):
        return ""
    return str(v).strip()


def is_blankish_text(v) -> bool:
    if v is None or is_nan(v):
        return True
    s = str(v).strip()
    if s == "":
        return True
    s2 = s.lower()
    if s2 in {"-", "—", "–", "n/a", "na", "nil", "tiada"}:
        return True
    if re.fullmatch(r"[-–—\s]+", s):
        return True
    return False


def parse_date_from_cell(val) -> Optional[dt.date]:
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return None
    if isinstance(val, dt.datetime):
        return val.date()
    if isinstance(val, dt.date):
        return val
    if isinstance(val, (int, float)) and 20000 < float(val) < 60000:
        base = dt.date(1899, 12, 30)
        return base + dt.timedelta(days=int(val))

    s = str(val).strip()
    if not s or s.lower() == "nan":
        return None

    m = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return dt.date(y, mo, d)
        except Exception:
            return None

    m = re.search(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return dt.date(y, mo, d)
        except Exception:
            return None

    return None


def parse_induk_code(val) -> str:
    if val is None or is_nan(val):
        return ""
    s = str(val).strip()
    if not s:
        return ""
    toks = re.findall(r"[A-Z]{2,5}", s.upper())
    return toks[-1] if toks else ""


def in_range(d: Optional[dt.date], start: dt.date, end: dt.date) -> bool:
    return d is not None and start <= d <= end


def normalize_osc_prefix(s: str) -> str:
    if not s:
        return ""
    s2 = str(s).strip()
    s2 = re.sub(r"[\s\r\n\t]+", "", s2)
    s2 = re.sub(r"^(MBPS|MPSP)", "MBSP", s2, flags=re.IGNORECASE)
    s2 = re.sub(r"^M\.?B\.?S\.?P", "MBSP", s2, flags=re.IGNORECASE)
    s2 = re.sub(r"^M\.?B\.?P\.?S", "MBSP", s2, flags=re.IGNORECASE)
    return s2.upper()


def sheet_norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip()).upper()


def canonical_sheet_name(sheet: str) -> str:
    """
    Canonicalize variasi nama sheet dalam excel supaya:
    - "TG" dianggap "TKR-GUNA"
    - "TKR GUNA" dianggap "TKR-GUNA"
    - "PKM TUKAR GUNA" dianggap "PKM TUKARGUNA"
    - "BGN-EVCB" dianggap "BGN EVCB"
    - "E V" jadi "EV"
    """
    s = sheet_norm(sheet).replace("_", " ")
    s = re.sub(r"\s+", " ", s).strip()

    if s in {"E V"}:
        return "EV"
    if s in {"TG", "TUKARGUNA", "TUKAR GUNA"}:
        return "TKR-GUNA"
    if s in {"TKR GUNA", "TKR-GUNA"}:
        return "TKR-GUNA"
    if s in {"PKM TUKAR GUNA", "PKM TUKARGUNA"}:
        return "PKM TUKARGUNA"
    if s in {"BGN-EVCB", "BGN EVCB"}:
        return "BGN EVCB"
    return s


def extract_tail_only(fail_no: str) -> str:
    s = normalize_osc_prefix(fail_no)
    m = re.search(r"/(\d{3,5})(?:[-A-Z\(]|$)", s)
    return m.group(1) if m else ""


def extract_series_tail_key(fail_no: str) -> str:
    s = normalize_osc_prefix(fail_no)
    m = re.search(r"^MBSP/\d+/([^/]+)/(\d{3,5})", s)
    if not m:
        return ""
    return f"{m.group(1)}|{m.group(2)}"


def extract_osc_head(fail_no: str) -> str:
    s = normalize_osc_prefix(fail_no)
    m = re.search(r"^(MBSP)/(\d+)/([^/]+)/(\d{3,5})", s)
    if not m:
        return ""
    return f"{m.group(1)}/{m.group(2)}/{m.group(3)}/{m.group(4)}"


def osc_norm(x: str) -> str:
    s = normalize_osc_prefix(str(x or ""))
    s = s.lower()
    s = re.sub(r"[\s\r\n\t]+", "", s)
    s = re.sub(r"[-/\\()\[\]{}+.,:;]", "", s)
    return s


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


def is_serentak(sheet_name: str, fail_no: str) -> bool:
    if canonical_sheet_name(sheet_name) == "SERENTAK":
        return True
    return "SERENTAK" in str(fail_no or "").upper()


def _sheet_implied_codes(sheet_u: str) -> Set[str]:
    s = sheet_u.upper()
    out = set()
    if "PKM" in s:
        out.add("PKM")
    if "TKR-GUNA" in s or s == "TG":
        out.add("TKR-GUNA")
    elif re.fullmatch(r"TKR", s):
        out.add("TKR")
    if "BGN" in s:
        out.add("BGN")
    if "EVCB" in s:
        out.add("EVCB")
    if re.fullmatch(r"EV", s):
        out.add("EV")
    if "TELCO" in s:
        out.add("TELCO")
    if re.fullmatch(r"PS", s):
        out.add("PS")
    if re.fullmatch(r"SB", s):
        out.add("SB")
    if re.fullmatch(r"CT", s):
        out.add("CT")
    if re.fullmatch(r"PL", s):
        out.add("PL")
    if re.fullmatch(r"KTUP", s):
        out.add("KTUP")
    if re.fullmatch(r"JP", s):
        out.add("JP")
    if re.fullmatch(r"LJUP", s):
        out.add("LJUP")
    return out


def extract_codes(fail_no: str, sheet_name: str) -> Set[str]:
    s = normalize_osc_prefix(str(fail_no or ""))
    tokens = re.split(r"[\s\+\-/\\(),]+", s.upper())
    codes: Set[str] = set()
    for t in tokens:
        if t in KNOWN_CODES:
            codes.add(t)

    sn = canonical_sheet_name(sheet_name)
    codes |= _sheet_implied_codes(sn)

    if sn == "BGN EVCB":
        codes.add("BGN")
        codes.add("EVCB")
    return codes


def split_fail_induk(fail_no: str) -> str:
    s = normalize_osc_prefix(str(fail_no or "")).strip()
    if not s:
        return s
    for i in range(len(s) - 1, 0, -1):
        if s[i] == "-":
            suffix = s[i + 1:].upper()
            if any(code in suffix for code in KNOWN_CODES):
                return s[:i]
    return s


def canon_serentak_codes(codes: Set[str]) -> List[str]:
    order = [
        "PKM", "TKR-GUNA", "TKR", "124A", "204D", "PS", "SB", "CT",
        "KTUP", "LJUP", "JP",
        "PL",
        "BGN", "EVCB", "EV", "TELCO",
    ]
    return [c for c in order if c in codes]


def perkara_3lines(d: Optional[dt.date]) -> str:
    dd = d.strftime("%d.%m.%Y") if d else ""
    return f"Penyediaan Kertas\nMesyuarat Tamat Tempoh\n{dd}"


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
        if is_blankish_text(p):
            continue
        key = re.sub(r"\s+", "", p.upper())
        if key in internal_map:
            internal.append(internal_map[key])
        else:
            external.append(p.upper())

    def dedup(seq):
        seen, out = set(), []
        for x in seq:
            if x not in seen:
                seen.add(x)
                out.append(x)
        return out

    internal = dedup(internal)
    external = dedup(external)
    return "\n".join(internal + external).strip()


def pemohon_norm(x: str) -> str:
    s = str(x or "").lower().strip()
    s = re.sub(r"\b(tetuan|tuan|puan)\b", "", s)
    s = re.sub(r"\b(sdn\.?\s*bhd\.?|sdn\s*bhd|bhd|berhad|enterprise|enterprises|plc|llp|ltd)\b", "", s)
    s = re.sub(r"[^a-z0-9]+", "", s)
    return s


def lot_tokens(x: str) -> Set[str]:
    toks = re.findall(r"\d{2,6}", str(x or ""))
    return set(toks)


# ============================================================
# AGENDA PARSER (WORD .docx) — tail-focused
# ============================================================
@dataclass
class AgendaBlock:
    is_ptj: bool
    codes: Set[str]
    osc_heads: List[str]
    tails: Set[str]
    series_tail_keys: Set[str]
    pemohon_key: str
    lot_set: Set[str]
    has_osc: bool


@dataclass
class AgendaIndex:
    tails_all: Set[str]
    series_tail_all: Set[str]
    osc_head_norm_all: Set[str]
    blocks: List[AgendaBlock]


HEADER_ANYWHERE_RE = re.compile(r"(?i)\bKERTAS\s+MESYUARAT\s+BIL\.\s*OSC/")
HEADER_CODE_RE = re.compile(r"(?i)OSC/([A-Z]{2,12}(?:-[A-Z]{2,12})?)/")

OSC_HEAD_RE = re.compile(
    r"(?i)\b(MBSP|MBPS|MPSP)\s*/\s*(\d+)\s*/\s*([A-Z0-9\-]+)\s*/\s*(\d{3,5})\b"
)

NO_RUJ_OSC_LINE_RE = re.compile(r"(?i)No\.?\s*Rujukan\s*OSC\s*:?\s*(.+)")
PEMOHON_LINE_RE = re.compile(r"(?i)Pemohon\s*:?\s*(.+)")


def _docx_collect_text(doc: Document) -> str:
    chunks: List[str] = []
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t:
            chunks.append(t)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                t = (cell.text or "").strip()
                if t:
                    chunks.append(t)
    return "\n".join(chunks)


def _extract_images_from_docx_bytes(file_bytes: bytes) -> List[bytes]:
    out = []
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            for name in z.namelist():
                if name.lower().startswith("word/media/") and name.lower().endswith(
                    (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff")
                ):
                    out.append(z.read(name))
    except Exception:
        return []
    return out


def _try_ocr_images(images: List[bytes]) -> str:
    if not images:
        return ""
    try:
        import pytesseract  # type: ignore
    except Exception:
        return ""

    texts = []
    for b in images:
        try:
            im = Image.open(io.BytesIO(b)).convert("RGB")
            txt = pytesseract.image_to_string(im, lang="eng")
            if txt and txt.strip():
                texts.append(txt)
        except Exception:
            continue
    return "\n".join(texts)


def _split_into_blocks(full_text: str) -> List[str]:
    lines = full_text.splitlines()
    idx = []
    for i, line in enumerate(lines):
        if HEADER_ANYWHERE_RE.search(line):
            idx.append(i)
    if not idx:
        return [full_text] if full_text.strip() else []

    blocks = []
    for j, start in enumerate(idx):
        end = idx[j + 1] if j + 1 < len(idx) else len(lines)
        blk = "\n".join(lines[start:end]).strip()
        if blk:
            blocks.append(blk)
    return blocks


def _parse_block_codes(header_line: str) -> Set[str]:
    codes = set()
    m = HEADER_CODE_RE.search(header_line)
    if not m:
        return codes
    raw = m.group(1).upper().strip()
    if raw in KNOWN_CODES:
        codes.add(raw)
    if raw in {"BGN-EVCB", "BGN EVCB"}:
        codes.add("BGN")
        codes.add("EVCB")
    return codes


def _parse_agenda_block(block_text: str) -> AgendaBlock:
    lines = block_text.splitlines()
    header_line = ""
    for ln in lines[:3]:
        if HEADER_ANYWHERE_RE.search(ln):
            header_line = ln.strip()
            break
    if not header_line:
        header_line = (lines[0].strip() if lines else "")

    is_ptj = bool(re.search(r"(?i)\bOSC/PTJ/", header_line))
    codes = _parse_block_codes(header_line)

    osc_heads: List[str] = []
    tails: Set[str] = set()
    series_tail_keys: Set[str] = set()
    has_osc = False

    for m in OSC_HEAD_RE.finditer(block_text):
        prefix = normalize_osc_prefix(m.group(1))
        yy = m.group(2)
        series = m.group(3).upper()
        tail = m.group(4)
        head = f"{prefix}/{yy}/{series}/{tail}"
        osc_heads.append(head)
        tails.add(tail)
        series_tail_keys.add(f"{series}|{tail}")
        has_osc = True

    for m in NO_RUJ_OSC_LINE_RE.finditer(block_text):
        rhs = (m.group(1) or "").strip()
        if not rhs:
            continue
        rhs2 = normalize_osc_prefix(rhs)
        if rhs2 in {"-", "—", "–"}:
            continue
        mm = OSC_HEAD_RE.search(rhs2)
        if mm:
            prefix = normalize_osc_prefix(mm.group(1))
            yy = mm.group(2)
            series = mm.group(3).upper()
            tail = mm.group(4)
            head = f"{prefix}/{yy}/{series}/{tail}"
            osc_heads.append(head)
            tails.add(tail)
            series_tail_keys.add(f"{series}|{tail}")
            has_osc = True

    seen = set()
    osc_heads2 = []
    for x in osc_heads:
        if x not in seen:
            seen.add(x)
            osc_heads2.append(x)
    osc_heads = osc_heads2

    pem = ""
    m = PEMOHON_LINE_RE.search(block_text)
    if m:
        pem = (m.group(1) or "").strip()
    else:
        m2 = re.search(r"(?i)\bTetuan\b\s*:?\s*(.+)", block_text)
        if m2:
            pem = (m2.group(1) or "").strip()
    pem_key = pemohon_norm(pem)

    lot_candidates = []
    for mm in re.finditer(r"(?i)\b(?:di\s+atas\s+)?lot\b[^.\n\r]{0,160}", block_text):
        lot_candidates.append(mm.group(0))
    for mm in re.finditer(r"(?i)\bPT\s*\d{1,6}\b", block_text):
        lot_candidates.append(mm.group(0))
    lot_s = " ".join(lot_candidates) if lot_candidates else block_text
    lot_set = set(re.findall(r"\d{2,6}", lot_s))

    return AgendaBlock(
        is_ptj=is_ptj,
        codes=codes,
        osc_heads=osc_heads,
        tails=tails,
        series_tail_keys=series_tail_keys,
        pemohon_key=pemohon_norm(pem),
        lot_set=lot_set,
        has_osc=has_osc,
    )


@st.cache_data(show_spinner=False)
def parse_agenda_docx(file_bytes: bytes, enable_ocr: bool = False) -> AgendaIndex:
    doc = Document(io.BytesIO(file_bytes))
    text_main = _docx_collect_text(doc)

    text_ocr = ""
    if enable_ocr:
        imgs = _extract_images_from_docx_bytes(file_bytes)
        text_ocr = _try_ocr_images(imgs)

    full_text = (text_main + "\n" + (text_ocr or "")).strip()

    blocks_raw = _split_into_blocks(full_text)
    blocks: List[AgendaBlock] = []

    tails_all: Set[str] = set()
    series_tail_all: Set[str] = set()
    osc_head_norm_all: Set[str] = set()

    for blk_text in blocks_raw:
        blk = _parse_agenda_block(blk_text)
        blocks.append(blk)

        if blk.is_ptj:
            continue

        tails_all |= set(blk.tails)
        series_tail_all |= set(blk.series_tail_keys)
        for h in blk.osc_heads:
            osc_head_norm_all.add(osc_norm(h))

    return AgendaIndex(
        tails_all=tails_all,
        series_tail_all=series_tail_all,
        osc_head_norm_all=osc_head_norm_all,
        blocks=blocks,
    )


# ============================================================
# EXCEL READER — FAST (tahan fail Google Sheets bengkak)
#   - load_workbook(read_only=True) => lebih ringan
#   - tidak read sheet 2 kali
#   - stop awal kalau banyak row kosong (formatting bengkak)
# ============================================================
HEADER_HINTS = [
    "No. Rujukan OSC",
    "No. Rujukan",
    "Rujukan OSC",
    "Pemaju",
    "Pemohon",
    "Daerah",
    "Mukim",
    "Lot",
    "Tempoh Untuk Proses",
    "Tempoh Untuk Diberi",
    "Tarikh Keputusan",
]

COL_CANDIDATES = {
    "fail_no": ["norujukanosc", "norujukan", "rujukanosc", "failno", "fail no", "no rujukan osc", "no rujukan"],
    "pemohon": ["pemajupemohon", "pemaju/pemohon", "pemaju", "pemohon", "tetuan"],
    "mukim": ["mukimseksyen", "mukim/seksyen", "mukim", "seksyen"],
    "lot": ["lot"],
    "km": ["tempohuntukprosesolehjabataninduk", "tempohuntukproses"],
    "ut": ["tempohuntukdiberiulasanolehjabatanteknikal", "tempohuntukdiberiulasan"],
    "belum": ["jabatanindukteknikalygbelummemberikeputusanulasansehinggakini", "belummemberikeputusanulasan", "ygbelummemberikeputusan", "belummemberi"],
    "keputusan": ["tarikhkeputusankuasa", "tarikhkeputusan"],
}


def norm_basic(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.strip().lower()
    s = re.sub(r"[\s\r\n\t]+", " ", s)
    s = re.sub(r"[^a-z0-9]+", "", s)
    return s


def _row_score(values: List[str]) -> int:
    joined = " | ".join(values).lower()
    score = 0
    for h in HEADER_HINTS:
        if h.lower() in joined:
            score += 1
    return score


def _find_header_row_ws(ws, max_scan_rows: int = 80) -> Tuple[Optional[int], int]:
    best_idx, best_score = None, 0
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan_rows, values_only=True), start=1):
        vals = [str(v).strip() if v is not None else "" for v in row]
        score = _row_score(vals)
        if score > best_score:
            best_score = score
            best_idx = i
    return best_idx, best_score


def _detect_col_indices(header_row_values: List[object]) -> Dict[str, int]:
    norm_headers = [norm_basic(v) for v in header_row_values]
    found: Dict[str, int] = {}
    for key, needles in COL_CANDIDATES.items():
        for needle in needles:
            needle_n = norm_basic(needle)
            for idx, h in enumerate(norm_headers):
                if needle_n and needle_n in h:
                    found[key] = idx
                    break
            if key in found:
                break
    return found


@st.cache_data(show_spinner=False)
def read_kertas_excel_fast(excel_bytes: bytes, daerah_label: str) -> List[dict]:
    out: List[dict] = []
    bio = io.BytesIO(excel_bytes)

    # read_only=True => cepat & ringan untuk fail besar
    wb = load_workbook(bio, read_only=True, data_only=True)
    allowed_upper = {s.upper() for s in ALLOWED_SHEETS}

    for sheet_name in wb.sheetnames:
        raw_name = (sheet_name or "").strip()
        sheet_clean = canonical_sheet_name(raw_name)

        if sheet_clean.upper() not in allowed_upper:
            continue

        ws = wb[sheet_name]
        hdr_row, score = _find_header_row_ws(ws, max_scan_rows=80)
        if hdr_row is None or score == 0:
            continue

        # ambil header row
        header_vals = []
        for cell in ws.iter_rows(min_row=hdr_row, max_row=hdr_row, values_only=True):
            header_vals = list(cell)
            break

        col_idx = _detect_col_indices(header_vals)
        if "fail_no" not in col_idx or "pemohon" not in col_idx:
            continue

        # indices yang kita perlukan sahaja
        idx_fail = col_idx.get("fail_no")
        idx_pem = col_idx.get("pemohon")
        idx_muk = col_idx.get("mukim")
        idx_lot = col_idx.get("lot")
        idx_km = col_idx.get("km")
        idx_ut = col_idx.get("ut")
        idx_belum = col_idx.get("belum")
        idx_kep = col_idx.get("keputusan")

        # NOTE: Google Sheets eksport kadang ada "max_row" besar sebab formatting.
        # Kita stop awal jika jumpa streak kosong panjang.
        empty_streak = 0
        EMPTY_STREAK_LIMIT = 250  # boleh naik/turun ikut real data

        for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
            # ambil nilai ikut kolum
            fail = row[idx_fail] if idx_fail is not None and idx_fail < len(row) else None
            pem = row[idx_pem] if idx_pem is not None and idx_pem < len(row) else None

            fail_s = str(fail).strip() if fail is not None else ""
            pem_s = str(pem).strip() if pem is not None else ""

            if fail_s == "" and pem_s == "":
                empty_streak += 1
                if empty_streak >= EMPTY_STREAK_LIMIT:
                    break
                continue

            empty_streak = 0

            km_raw = row[idx_km] if idx_km is not None and idx_km < len(row) else None
            ut_raw = row[idx_ut] if idx_ut is not None and idx_ut < len(row) else None

            fail_raw = normalize_osc_prefix(clean_fail_no(fail_s))

            rec = {
                "daerah": daerah_label,
                "sheet": sheet_clean,
                "fail_no_raw": fail_raw,
                "pemohon": clean_str(pem),
                "mukim": clean_str(row[idx_muk]) if idx_muk is not None and idx_muk < len(row) else "",
                "lot": clean_str(row[idx_lot]) if idx_lot is not None and idx_lot < len(row) else "",
                "km_date": parse_date_from_cell(km_raw) if idx_km is not None else None,
                "ut_date": parse_date_from_cell(ut_raw) if idx_ut is not None else None,
                "belum": clean_str(row[idx_belum]) if idx_belum is not None and idx_belum < len(row) else "",
                "keputusan": clean_str(row[idx_kep]) if idx_kep is not None and idx_kep < len(row) else "",
                "induk_code": parse_induk_code(km_raw),
            }
            out.append(rec)

    try:
        wb.close()
    except Exception:
        pass

    return out


# ============================================================
# BUILD CATEGORIES
# ============================================================
def enrich_rows(rows: List[dict]) -> List[dict]:
    out = []
    for r in rows:
        rr = dict(r)
        rr["sheet_u"] = canonical_sheet_name(r["sheet"])
        rr["codes"] = extract_codes(r["fail_no_raw"], rr["sheet_u"])
        rr["serentak"] = is_serentak(rr["sheet_u"], r["fail_no_raw"])
        rr["fail_induk"] = split_fail_induk(r["fail_no_raw"])

        rr["tail"] = extract_tail_only(r["fail_no_raw"])
        rr["series_tail"] = extract_series_tail_key(r["fail_no_raw"])
        rr["osc_head_norm"] = osc_norm(extract_osc_head(r["fail_no_raw"]))

        rr["pemohon_key"] = pemohon_norm(r.get("pemohon", ""))
        rr["lot_set"] = lot_tokens(r.get("lot", ""))
        out.append(rr)
    return out


def sheet_is_ut_allowed(sheet_u: str) -> bool:
    s = canonical_sheet_name(sheet_u)
    if s in UT_ALLOWED_SHEETS:
        return True
    if "GUNA" in s and ("TKR" in s or "TUKAR" in s or s == "TG"):
        return True
    return False


def _agenda_fallback_match(row: dict, agenda: "AgendaIndex") -> bool:
    if row["sheet_u"] not in AGENDA_FILTER_SHEETS:
        return False
    if not row.get("pemohon_key"):
        return False
    if not row.get("lot_set"):
        return False

    row_codes = set(row.get("codes") or set())

    for blk in agenda.blocks:
        if blk.is_ptj:
            continue
        if not blk.pemohon_key or not blk.lot_set:
            continue

        if blk.codes and not (row_codes & blk.codes):
            continue

        if row["pemohon_key"] != blk.pemohon_key:
            continue

        inter = row["lot_set"] & blk.lot_set
        if not inter:
            continue

        if min(len(row["lot_set"]), len(blk.lot_set)) >= 2 and len(inter) < 2:
            continue

        return True

    return False


def build_categories(
    rows: List[dict],
    agenda: Optional["AgendaIndex"],
    km_start: dt.date,
    km_end: dt.date,
    ut_start: dt.date,
    ut_end: dt.date,
    ut_enabled: bool,
    agenda_enabled: bool,
) -> Tuple[List[dict], List[dict], List[dict], List[dict], List[dict]]:

    rows = [r for r in rows if keputusan_is_empty(r.get("keputusan"))]

    # Tapisan agenda hanya untuk AGENDA_FILTER_SHEETS, dan PTJ dalam agenda dikecualikan (tak ditapis)
    if agenda_enabled and agenda:
        def _keep(r: dict) -> bool:
            if r["sheet_u"] not in AGENDA_FILTER_SHEETS:
                return True

            # PRIORITY ikut cara Unit OSC: NO UNIK HUJUNG (tail)
            if r.get("tail") and r["tail"] in agenda.tails_all:
                return False

            # backup lebih ketat: series|tail
            if r.get("series_tail") and r["series_tail"] in agenda.series_tail_all:
                return False

            # backup lagi: head normalized
            if r.get("osc_head_norm") and r["osc_head_norm"] in agenda.osc_head_norm_all:
                return False

            # fallback pemohon+lot (strict)
            if _agenda_fallback_match(r, agenda):
                return False

            return True

        rows = [r for r in rows if _keep(r)]

    by_induk: Dict[str, List[dict]] = {}
    for r in rows:
        by_induk.setdefault(r["fail_induk"], []).append(r)

    def nama_simplify(x: str) -> str:
        return pemohon_norm(x)

    def make_rec(cat: int, tindakan: str, base_r: dict, jenis: str, fail_no: str, perkara: str, extra_key: str) -> dict:
        return {
            "cat": cat,
            "tindakan": tindakan,
            "jenis": jenis,
            "fail_no": fail_no,
            "pemohon": base_r["pemohon"],
            "daerah": base_r["daerah"],
            "mukim": base_r["mukim"],
            "lot": base_r["lot"],
            "perkara": perkara,
            "dedup_key": f"{cat}|{tindakan}|{osc_norm(fail_no)}|{nama_simplify(base_r['pemohon'])}|{extra_key}",
        }

    cat1, cat2, cat3, cat4, cat5 = [], [], [], [], []

    for induk, grp in by_induk.items():
        is_ser = any(g["serentak"] for g in grp)

        union_codes: Set[str] = set()
        km_dates = [g["km_date"] for g in grp if g.get("km_date")]
        for g in grp:
            union_codes |= set(g["codes"])
        km_date = min(km_dates) if km_dates else None

        codes_sorted = canon_serentak_codes(union_codes)
        codes_join = "+".join(codes_sorted)
        jenis_ser = (f"{codes_join} (Serentak)".strip() if codes_join else "(Serentak)").strip()
        fail_no_ser = f"{induk}-{codes_join}" if codes_join else induk

        # KATEGORI 1 — KM
        if is_ser and in_range(km_date, km_start, km_end):
            if union_codes & (PB_CODES - {"PS", "SB", "CT"}):
                cat1.append(make_rec(1, "Pengarah Perancang Bandar", grp[0], jenis_ser, fail_no_ser, perkara_3lines(km_date), "SER-PB"))
            if union_codes & {"BGN", "EVCB", "EV", "TELCO"}:
                cat1.append(make_rec(1, "Pengarah Bangunan", grp[0], jenis_ser, fail_no_ser, perkara_3lines(km_date), "SER-BGN"))

        if not is_ser:
            for g in grp:
                if not in_range(g.get("km_date"), km_start, km_end):
                    continue
                if g["codes"] & {"PKM", "TKR", "TKR-GUNA"}:
                    cat1.append(make_rec(1, "Pengarah Perancang Bandar", g, g["sheet_u"], g["fail_no_raw"], perkara_3lines(g.get("km_date")), "NS-PB"))
                if g["codes"] & {"BGN", "EVCB", "EV", "TELCO"}:
                    cat1.append(make_rec(1, "Pengarah Bangunan", g, g["sheet_u"], g["fail_no_raw"], perkara_3lines(g.get("km_date")), "NS-BGN"))

        # KATEGORI 2 — UT
        if ut_enabled:
            for g in grp:
                if not sheet_is_ut_allowed(g["sheet_u"]):
                    continue
                if not in_range(g.get("ut_date"), ut_start, ut_end):
                    continue
                if is_blankish_text(g.get("belum")):
                    continue

                if g["sheet_u"] == "SERENTAK":
                    if (g.get("induk_code") or "") not in SERENTAK_UT_ALLOWED_INDUK:
                        continue

                tindakan = tindakan_ut(g.get("belum", ""))
                if is_blankish_text(tindakan):
                    continue

                perkara = f"Ulasan teknikal belum dikemukakan. Tamat Tempoh {g['ut_date'].strftime('%d.%m.%Y')}."
                jenis = jenis_ser if is_ser else g["sheet_u"]
                fail_no = fail_no_ser if is_ser else g["fail_no_raw"]

                extra_key = f"{g['sheet_u']}|{g['ut_date'].isoformat()}|{(g.get('belum') or '').strip()}"
                cat2.append(make_rec(2, tindakan, g, jenis, fail_no, perkara, extra_key))

        # KATEGORI 3/4/5 — KM
        if is_ser and in_range(km_date, km_start, km_end):
            if union_codes & KEJ_CODES:
                cat3.append(make_rec(3, "Pengarah Kejuruteraan", grp[0], jenis_ser, fail_no_ser, perkara_3lines(km_date), "SER-KEJ"))
            if union_codes & JL_CODES:
                cat4.append(make_rec(4, "Pengarah Landskap", grp[0], jenis_ser, fail_no_ser, perkara_3lines(km_date), "SER-JL"))
            if union_codes & {"124A", "204D"}:
                cat5.append(make_rec(5, "Pengarah Perancang Bandar", grp[0], jenis_ser, fail_no_ser, perkara_3lines(km_date), "SER-124A204D"))

        if not is_ser:
            for g in grp:
                if in_range(g.get("km_date"), km_start, km_end) and (g["sheet_u"] in {"KTUP", "JP", "LJUP"}):
                    cat3.append(make_rec(3, "Pengarah Kejuruteraan", g, g["sheet_u"], g["fail_no_raw"], perkara_3lines(g.get("km_date")), f"NS-{g['sheet_u']}"))
                if in_range(g.get("km_date"), km_start, km_end) and (g["sheet_u"] == "PL"):
                    cat4.append(make_rec(4, "Pengarah Landskap", g, g["sheet_u"], g["fail_no_raw"], perkara_3lines(g.get("km_date")), "NS-PL"))
                if in_range(g.get("km_date"), km_start, km_end) and (g["sheet_u"] in {"PS", "SB", "CT"}):
                    cat5.append(make_rec(5, "Pengarah Perancang Bandar", g, g["sheet_u"], g["fail_no_raw"], perkara_3lines(g.get("km_date")), f"NS-{g['sheet_u']}"))

    def dedup_list(lst: List[dict]) -> List[dict]:
        seen, out = set(), []
        for r in lst:
            if r["dedup_key"] in seen:
                continue
            seen.add(r["dedup_key"])
            out.append(r)
        return out

    cat1, cat2, cat3, cat4, cat5 = map(dedup_list, [cat1, cat2, cat3, cat4, cat5])

    cat1.sort(key=lambda r: (0 if r["tindakan"].startswith("Pengarah Perancang") else 1, DAERAH_ORDER.get(r["daerah"], 9), r["fail_no"]))
    cat2.sort(key=lambda r: (DAERAH_ORDER.get(r["daerah"], 9), r["fail_no"], r["tindakan"]))
    cat3.sort(key=lambda r: (DAERAH_ORDER.get(r["daerah"], 9), r["fail_no"]))
    cat4.sort(key=lambda r: (DAERAH_ORDER.get(r["daerah"], 9), r["fail_no"]))
    cat5.sort(key=lambda r: (DAERAH_ORDER.get(r["daerah"], 9), r["fail_no"]))

    for lst in [cat1, cat2, cat3, cat4, cat5]:
        for i, r in enumerate(lst, start=1):
            r["bil"] = i

    return cat1, cat2, cat3, cat4, cat5


# ============================================================
# WORD FORMATTER
# ============================================================
COL_WIDTHS_IN = [0.50, 1.75, 1.55, 1.55, 1.45, 0.75, 0.65, 0.70, 1.99]
HEADERS = ["BIL", "TINDAKAN", "JENIS\nPERMOHONAN", "FAIL NO", "PEMAJU/PEMOHON", "DAERAH", "MUKIM", "LOT", "PERKARA"]


def _find_font_path(prefer_bold: bool = True) -> Optional[str]:
    candidates = [
        "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman_Bold.ttf",
        "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
    ]
    for p in candidates:
        if os.path.exists(p) and (not prefer_bold or "Bold" in os.path.basename(p)):
            return p
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


# NOTE: Besarkan huruf G => naikkan font_pt
def make_g_logo_png(diameter_px: int = 140, outline_px: int = 4, font_pt: int = 34) -> bytes:
    scale = 4
    D = diameter_px * scale
    img = Image.new("RGBA", (D, D), (255, 255, 255, 0))
    dr = ImageDraw.Draw(img)

    pad = (outline_px + 4) * scale
    dr.ellipse((pad, pad, D - pad, D - pad), outline=(0, 0, 0, 255), width=outline_px * scale)

    font_path = _find_font_path(prefer_bold=True)
    size_px = int(font_pt * 96 / 72) * scale
    if font_path:
        try:
            font = ImageFont.truetype(font_path, size_px)
        except Exception:
            font = ImageFont.load_default()
    else:
        font = ImageFont.load_default()

    bbox = dr.textbbox((0, 0), "G", font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    x = (D - tw) / 2 - bbox[0]
    y = (D - th) / 2 - bbox[1] - int(2 * scale)
    dr.text((x, y), "G", font=font, fill=(0, 0, 0, 255))

    img_small = img.resize((diameter_px, diameter_px), resample=Image.LANCZOS)
    buf = io.BytesIO()
    img_small.save(buf, format="PNG")
    return buf.getvalue()


def set_section_landscape(sec):
    sec.orientation = WD_ORIENTATION.LANDSCAPE
    sec.page_width = Inches(11.69)
    sec.page_height = Inches(8.27)
    sec.left_margin = Inches(0.4)
    sec.right_margin = Inches(0.4)
    sec.top_margin = Inches(1.0)
    sec.bottom_margin = Inches(1.0)


def clear_header(hdr):
    for p in list(hdr.paragraphs):
        p._element.getparent().remove(p._element)


def add_logo_first_page(sec, logo_png_bytes: bytes):
    sec.different_first_page_header_footer = True
    sec.header.is_linked_to_previous = False
    sec.first_page_header.is_linked_to_previous = False
    sec.footer.is_linked_to_previous = False
    sec.first_page_footer.is_linked_to_previous = False

    clear_header(sec.header)
    clear_header(sec.first_page_header)

    hdr = sec.first_page_header
    p = hdr.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run()
    run.add_picture(io.BytesIO(logo_png_bytes), width=Inches(0.70))


def set_paragraph_font(p, font_name: str, size_pt: float, bold: bool = False, align=None):
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1
    for r in p.runs:
        r.font.name = font_name
        r._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
        r.font.size = Pt(size_pt)
        r.font.bold = bold


def add_blank(doc: Document):
    p = doc.add_paragraph("")
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1


def add_title_line_main(doc: Document):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1

    before = "PERMOHONAN YANG "
    mid = "TELAH DAN AKAN SAMPAI"
    after = " TEMPOH & ULASAN JABATAN TEKNIKAL BELUM DIKEMUKAKAN"

    r1 = p.add_run(before)
    r2 = p.add_run(mid)
    r3 = p.add_run(after)

    for r in [r1, r2, r3]:
        r.font.name = "Trebuchet MS"
        r._element.rPr.rFonts.set(qn("w:eastAsia"), "Trebuchet MS")
        r.font.bold = True
        r.font.size = Pt(12)

    r2.font.size = Pt(13.5)


def add_center_bold(doc: Document, text: str, font: str = "Trebuchet MS", size: float = 12):
    p = doc.add_paragraph(text)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_font(p, font, size, bold=True)


def add_km_line(doc: Document, km_start: dt.date, km_end: dt.date):
    p = doc.add_paragraph(f"KERTAS MESYUARAT (TEMPOH {km_start.strftime('%d/%m/%Y')} HINGGA {km_end.strftime('%d/%m/%Y')})")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_paragraph_font(p, "Trebuchet MS", 12, bold=True)
    add_blank(doc)


def add_ut_line(doc: Document, ut_start: dt.date, ut_end: dt.date):
    p = doc.add_paragraph(f"ULASAN TEKNIKAL (TEMPOH {ut_start.strftime('%d/%m/%Y')} HINGGA {ut_end.strftime('%d/%m/%Y')})")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_paragraph_font(p, "Arial", 12, bold=True)
    add_blank(doc)


def set_cell_vcenter(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    vAlign = tcPr.find(qn("w:vAlign"))
    if vAlign is None:
        vAlign = OxmlElement("w:vAlign")
        tcPr.append(vAlign)
    vAlign.set(qn("w:val"), "center")


def set_row_as_header(row):
    trPr = row._tr.get_or_add_trPr()
    tblHeader = trPr.find(qn("w:tblHeader"))
    if tblHeader is None:
        tblHeader = OxmlElement("w:tblHeader")
        trPr.append(tblHeader)
    tblHeader.set(qn("w:val"), "true")


def set_table_borders(tbl):
    tbl_pr = tbl._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)

    def _edge(tag):
        el = borders.find(qn(f"w:{tag}"))
        if el is None:
            el = OxmlElement(f"w:{tag}")
            borders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "8")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")

    for t in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        _edge(t)


def format_table(tbl):
    tbl.autofit = False
    set_table_borders(tbl)

    for row in tbl.rows:
        for i, cell in enumerate(row.cells):
            cell.width = Inches(COL_WIDTHS_IN[i])
    for i, col in enumerate(tbl.columns):
        col.width = Inches(COL_WIDTHS_IN[i])

    hdr_row = tbl.rows[0]
    set_row_as_header(hdr_row)

    for i, cell in enumerate(hdr_row.cells):
        cell.text = HEADERS[i]
        set_cell_vcenter(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1
        for run in p.runs:
            run.font.name = "Arial"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
            run.font.size = Pt(9)
            run.font.bold = True

    for r in tbl.rows[1:]:
        for c in r.cells:
            set_cell_vcenter(c)
            for p in c.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pf = p.paragraph_format
                pf.space_before = Pt(0)
                pf.space_after = Pt(0)
                pf.line_spacing = 1
                for run in p.runs:
                    run.font.name = "Arial"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
                    run.font.size = Pt(9)
                    run.font.bold = False


def fill_table(tbl, recs: List[dict]):
    note_fields = ["bil", "tindakan", "jenis", "fail_no", "pemohon", "daerah", "mukim", "lot", "perkara"]
    for rec in recs:
        row = tbl.add_row()
        vals = [str(rec.get(k, "")) for k in note_fields]
        for i, val in enumerate(vals):
            cell = row.cells[i]
            cell.text = ""
            set_cell_vcenter(cell)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = 1
            run = p.add_run(str(val))
            run.font.name = "Arial"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
            run.font.size = Pt(9)
            run.font.bold = False


def build_word_doc(
    meeting_info: str,
    km_start: dt.date,
    km_end: dt.date,
    ut_start: dt.date,
    ut_end: dt.date,
    cat1: List[dict],
    cat2: List[dict],
    cat3: List[dict],
    cat4: List[dict],
    cat5: List[dict],
    ut_enabled: bool,
) -> bytes:
    logo_png = make_g_logo_png()

    doc = Document()
    set_section_landscape(doc.sections[0])

    def add_category_section(cat_num: int, recs: List[dict]):
        if cat_num == 1:
            sec = doc.sections[0]
        else:
            sec = doc.add_section(WD_SECTION.NEW_PAGE)
            set_section_landscape(sec)

        add_logo_first_page(sec, logo_png)

        if cat_num == 2:
            add_ut_line(doc, ut_start, ut_end)
        else:
            add_title_line_main(doc)
            if cat_num in {3, 4, 5}:
                bagi = {
                    3: "BAGI PELAN KEJURUTERAAN",
                    4: "BAGI PELAN LANDSKAP",
                    5: "BAGI PELAN PS / SB / CT",
                }[cat_num]
                add_center_bold(doc, bagi)
            add_center_bold(doc, meeting_info.strip())
            add_blank(doc)
            add_km_line(doc, km_start, km_end)

        tbl = doc.add_table(rows=1, cols=9)
        format_table(tbl)
        fill_table(tbl, recs)

    add_category_section(1, cat1)
    if ut_enabled:
        add_category_section(2, cat2)
    add_category_section(3, cat3)
    add_category_section(4, cat4)
    add_category_section(5, cat5)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ============================================================
# STREAMLIT UI
# ============================================================
bg_ok = _inject_bg_and_css("assets/bg.jpg")
if not bg_ok:
    st.warning("Background tidak dijumpai. Pastikan fail ada di folder assets/ (contoh: assets/bg.jpg).")

st.markdown("<h1 class='app-title'>LAMPIRAN G UNIT OSC</h1>", unsafe_allow_html=True)
st.markdown("<div class='hero-spacer'></div>", unsafe_allow_html=True)

with st.expander("Nota Penting: Upload Limit 1GB (Setting Server)", expanded=False):
    st.markdown(
        """
**Limit upload Streamlit tidak boleh diubah dalam Python code.**  
Untuk naikkan limit sampai **1GB**, set dalam file config server:

**.streamlit/config.toml**
```toml
[server]
maxUploadSize = 1024  # MB (1GB)
