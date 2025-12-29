import io
import math
import os
import re
import base64
import datetime as dt
from typing import Dict, List, Optional, Tuple, Set

import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont

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

# Urutan daerah untuk sort
DAERAH_ORDER = {"SPU": 0, "SPS": 1, "SPT": 2}

# Kod rujukan (untuk kenal pasti jenis)
KNOWN_CODES = [
    "PKM", "TKR-GUNA", "TKR", "124A", "204D", "PS", "SB", "CT",
    "KTUP", "LJUP", "JP", "PL",
    "BGN", "EVCB", "EV", "TELCO",
]
PB_CODES = {"PKM", "TKR-GUNA", "TKR", "124A", "204D", "PS", "SB", "CT"}
BGN_CODES = {"BGN", "BGN EVCB", "EVCB", "EV", "TELCO"}
KEJ_CODES = {"KTUP", "LJUP", "JP"}
JL_CODES = {"PL"}

# Agenda filter hanya valid untuk kategori SERENTAK/PKM/BGN sahaja (ikut arahan terbaru).
AGENDA_FILTER_SHEETS = {"SERENTAK", "PKM", "BGN", "BGN EVCB"}

# Kategori 2 (UT) hanya untuk sheet ini sahaja.
UT_ALLOWED_SHEETS = {"SERENTAK", "PKM", "BGN", "BGN EVCB", "TKR-GUNA", "PKM TUKARGUNA"}

# Untuk sheet SERENTAK dalam Kategori 2:
# hanya benarkan jika "Tempoh Untuk Proses Oleh Jabatan Induk*" menunjuk induk PB/PKM/BGN.
SERENTAK_UT_ALLOWED_INDUK = {"PB", "PKM", "BGN"}


# ============================================================
# UI HELPERS (BACKGROUND + CSS)
# ============================================================
def _inject_bg_and_css(img_path: str) -> bool:
    """Background fixed layer + CSS stabilizer."""
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

      .hero-spacer {{
        height: 22vh;
      }}

      div[data-testid="stVerticalBlockBorderWrapper"] {{
        background: rgba(0,0,0,0.44) !important;
        border: 1px solid rgba(255,255,255,0.12) !important;
        border-radius: 18px !important;
        padding: 14px 16px 12px 16px !important;
        box-shadow: 0 10px 30px rgba(0,0,0,0.25);
        backdrop-filter: blur(2px);
      }}

      h1 a, h2 a, h3 a {{
        display: none !important;
      }}

      label {{
        font-size: 0.85rem !important;
      }}
      .stTextInput input {{
        height: 2.35rem !important;
      }}

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
    """Fail/No Rujukan OSC biasanya tak perlukan whitespace. Buang semua whitespace untuk elak format pelik."""
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
    """Anggap kosong jika None / "" / '-' / '—' / '–' / N/A / NA / TIADA / NIL."""
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
    """Terima datetime/date/excel-serial/string seperti '73 Hari (27/12/2025)'."""
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
    """
    Extract kod induk dari hujung cell "Tempoh Untuk Proses Oleh Jabatan Induk*".
    Contoh: '73 Hari (05/02/2026) PB' -> 'PB'
    """
    if val is None or is_nan(val):
        return ""
    s = str(val).strip()
    if not s:
        return ""
    toks = re.findall(r"[A-Z]{2,5}", s.upper())
    if not toks:
        return ""
    return toks[-1]


def in_range(d: Optional[dt.date], start: dt.date, end: dt.date) -> bool:
    return d is not None and start <= d <= end


def osc_norm(x: str) -> str:
    s = str(x or "").lower()
    s = re.sub(r"[\s\r\n\t]+", "", s)
    s = re.sub(r"[-/\\()\[\]{}+.,:;]", "", s)
    return s


def keputusan_is_empty(v) -> bool:
    """
    Keputusan dianggap kosong jika:
    - empty / dash / tiada / nil / n/a
    Jika ada apa-apa teks lain atau tarikh -> dianggap ADA keputusan.
    """
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


def sheet_norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip()).upper()


def is_serentak(sheet_name: str, fail_no: str) -> bool:
    if sheet_norm(sheet_name) == "SERENTAK":
        return True
    s = str(fail_no or "").upper()
    return "SERENTAK" in s


def extract_codes(fail_no: str, sheet_name: str) -> Set[str]:
    s = str(fail_no or "")
    tokens = re.split(r"[\s\+\-/\\(),]+", s.upper())
    codes: Set[str] = set()
    for t in tokens:
        if t in KNOWN_CODES:
            codes.add(t)

    sn = sheet_norm(sheet_name)
    if sn == "BGN EVCB":
        codes.add("BGN")
        codes.add("EVCB")
    elif sn in {c for c in ALLOWED_SHEETS}:
        if sn in {"PKM", "TKR", "TKR-GUNA", "BGN", "EVCB", "EV", "TELCO", "PS", "SB", "CT", "PL", "KTUP", "JP", "LJUP"}:
            codes.add(sn)

    return codes


def split_fail_induk(fail_no: str) -> str:
    """Fail Induk = bahagian No Rujukan OSC sebelum '-<kod+kod...>'."""
    s = str(fail_no or "").strip()
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
    """
    Convert list jabatan dari kolum "Belum memberi ulasan" -> TINDAKAN.
    - Jika kod dalaman: tukar jadi "Pengarah ...."
    - Jika luaran: kekal ringkas (uppercase)
    """
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


# ============================================================
# AGENDA PARSER (WORD .docx)
# ============================================================
OSC_PATTERN = re.compile(r"MBSP[0-9A-Z/\-\+\s]{10,120}", flags=re.IGNORECASE)

def parse_agenda_docx(file_bytes: bytes) -> Set[str]:
    """
    Tapisan agenda ikut arahan terbaru:
    - Jangan tapis ikut 'Tetuan' (nama pemaju) sebab boleh tersalah buang (contoh: Bertam/ PTJ).
    - Hanya guna padanan No. Rujukan OSC yang betul.
    """
    doc = Document(io.BytesIO(file_bytes))
    texts: List[str] = []

    for p in doc.paragraphs:
        if p.text:
            texts.append(p.text)

    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                if cell.text:
                    texts.append(cell.text)

    full = "\n".join(texts)

    osc_set: Set[str] = set()
    for m in OSC_PATTERN.finditer(full):
        cand = m.group(0)
        cand = cand.strip(" ,.;")
        cand = re.sub(r"[\s\r\n\t]+", "", cand)
        if cand.upper().startswith("MBSP") and "/" in cand and "-" in cand:
            osc_set.add(osc_norm(cand))

    return osc_set


# ============================================================
# EXCEL READER (auto-detect header row)
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
    "fail_no": ["norujukanosc", "no rujukan osc", "rujukan osc", "fail no", "failno", "no rujukan"],
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


def find_header_row(excel_bytes: bytes, sheet: str) -> Tuple[Optional[int], int]:
    raw = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=sheet, header=None, engine="openpyxl", nrows=80)
    best_idx, best_score = None, 0
    for i in range(len(raw)):
        row = raw.iloc[i].astype(str).fillna("")
        joined = " | ".join(row.tolist())
        score = 0
        for h in HEADER_HINTS:
            if h.lower() in joined.lower():
                score += 1
        if score > best_score:
            best_score, best_idx = score, i
    return best_idx, best_score


def detect_columns(df: pd.DataFrame) -> Dict[str, str]:
    norm_map = {col: norm_basic(col) for col in df.columns}
    found: Dict[str, str] = {}
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
    out: List[dict] = []
    xl = pd.ExcelFile(io.BytesIO(excel_bytes), engine="openpyxl")

    allowed_upper = {s.upper() for s in ALLOWED_SHEETS}

    for sheet in xl.sheet_names:
        sheet_clean = (sheet or "").strip()
        if sheet_clean.upper() not in allowed_upper:
            continue

        hdr_idx, score = find_header_row(excel_bytes, sheet)
        if hdr_idx is None or score == 0:
            continue

        df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=sheet, header=hdr_idx, engine="openpyxl")
        df = df.dropna(how="all")
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
                "sheet": sheet_clean,
                "fail_no_raw": clean_fail_no(fail),
                "pemohon": clean_str(pem),
                "mukim": clean_str(row.get(cols["mukim"])) if "mukim" in cols else "",
                "lot": clean_str(row.get(cols["lot"])) if "lot" in cols else "",
                "km_date": parse_date_from_cell(km_raw) if "km" in cols else None,
                "ut_date": parse_date_from_cell(row.get(cols["ut"])) if "ut" in cols else None,
                "belum": clean_str(row.get(cols["belum"])) if "belum" in cols else "",
                "keputusan": clean_str(row.get(cols["keputusan"])) if "keputusan" in cols else "",
                "induk_code": parse_induk_code(km_raw),
            }
            out.append(rec)

    return out


# ============================================================
# BUILD CATEGORIES
# ============================================================
def enrich_rows(rows: List[dict]) -> List[dict]:
    out = []
    for r in rows:
        codes = extract_codes(r["fail_no_raw"], r["sheet"])
        rr = dict(r)
        rr["codes"] = codes
        rr["serentak"] = is_serentak(r["sheet"], r["fail_no_raw"])
        rr["fail_induk"] = split_fail_induk(r["fail_no_raw"])
        rr["osc_norm"] = osc_norm(r["fail_no_raw"])
        rr["sheet_u"] = sheet_norm(r["sheet"])
        out.append(rr)
    return out


def sheet_is_ut_allowed(sheet_u: str) -> bool:
    s = sheet_u
    if s in UT_ALLOWED_SHEETS:
        return True
    # handle variasi tukar guna
    if "GUNA" in s and ("TKR" in s or "TUKAR" in s or s == "TG"):
        return True
    return False


def build_categories(
    rows: List[dict],
    agenda_osc_set: Set[str],
    km_start: dt.date,
    km_end: dt.date,
    ut_start: dt.date,
    ut_end: dt.date,
    ut_enabled: bool,
    agenda_enabled: bool,
) -> Tuple[List[dict], List[dict], List[dict], List[dict], List[dict]]:

    # 1) Buang yang ada keputusan
    rows = [r for r in rows if keputusan_is_empty(r.get("keputusan"))]

    # 2) Tapisan agenda: hanya untuk sheet SERENTAK/PKM/BGN sahaja (ikut arahan terbaru)
    if agenda_enabled and agenda_osc_set:
        rows = [
            r for r in rows
            if not (r["sheet_u"] in AGENDA_FILTER_SHEETS and r["osc_norm"] in agenda_osc_set)
        ]

    # group by fail_induk (handle serentak & dedup)
    by_induk: Dict[str, List[dict]] = {}
    for r in rows:
        by_induk.setdefault(r["fail_induk"], []).append(r)

    def nama_simplify(x: str) -> str:
        # untuk dedup sahaja (bukan untuk tapisan agenda)
        s = str(x or "").lower()
        s = re.sub(r"\b(tetuan|tuan|puan)\b", "", s)
        s = re.sub(r"\b(sdn\.?\s*bhd\.?|sdn\s*bhd|bhd|berhad|enterprise|enterprises|plc|llp|ltd)\b", "", s)
        s = re.sub(r"[^a-z0-9]+", "", s)
        return s

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

        # ----------------------------
        # KATEGORI 1 — KM (PB/BGN)
        # ----------------------------
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

        # ----------------------------
        # KATEGORI 2 — UT (TERHAD)
        # Syarat:
        # - sheet mesti: SERENTAK / PKM / BGN / TUKAR GUNA
        # - ut_date dalam range UT
        # - kolum "Belum memberi..." mesti ADA isi (bukan kosong/dash)
        # - keputusan mesti kosong (dah ditapis awal)
        # - untuk SERENTAK: induk_code (KM) mesti PB/PKM/BGN
        # ----------------------------
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

        # ----------------------------
        # KATEGORI 3/4/5 — KM
        # ----------------------------
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
# Lebar kolum dibetulkan supaya header "BIL" dan "DAERAH" tak pecah baris.
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
        if os.path.exists(p):
            if prefer_bold and "Bold" in os.path.basename(p):
                return p
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


def make_g_logo_png(diameter_px: int = 140, outline_px: int = 4, font_pt: int = 26) -> bytes:
    """
    Logo G: Times New Roman Bold, huruf G dibesarkan sedikit (ikut request terbaru).
    Kalau nak lagi besar: naikkan font_pt (contoh 28/30).
    """
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

    text = "G"
    bbox = dr.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    x = (D - tw) / 2 - bbox[0]
    y = (D - th) / 2 - bbox[1] - int(2 * scale)
    dr.text((x, y), text, font=font, fill=(0, 0, 0, 255))

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
    for rec in recs:
        row = tbl.add_row()
        vals = [
            str(rec.get("bil", "")),
            rec.get("tindakan", ""),
            rec.get("jenis", ""),
            rec.get("fail_no", ""),
            rec.get("pemohon", ""),
            rec.get("daerah", ""),
            rec.get("mukim", ""),
            rec.get("lot", ""),
            rec.get("perkara", ""),
        ]
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

if "running" not in st.session_state:
    st.session_state.running = False

left_col, right_col = st.columns([1.15, 0.85], gap="large")

with left_col:
    with st.container(border=True):
        st.markdown("### Maklumat Mesyuarat")
        st.markdown("**Maklumat mesyuarat**")
        meeting_info = st.text_input("", value="", key="meeting_info", label_visibility="collapsed")

        st.markdown("### Tempoh Kertas Mesyuarat (KM)")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**KM Mula (dd/mm/yyyy)**")
            km_mula_str = st.text_input("", value="", key="km_mula", label_visibility="collapsed")
        with c2:
            st.markdown("**KM Akhir (dd/mm/yyyy)**")
            km_akhir_str = st.text_input("", value="", key="km_akhir", label_visibility="collapsed")

        ut_enabled = st.checkbox("Aktifkan Ulasan Teknikal (UT)", value=True)

        ut_mula_str, ut_akhir_str = "", ""
        if ut_enabled:
            st.markdown("### Tempoh Ulasan Teknikal (UT)")
            u1, u2 = st.columns(2)
            with u1:
                st.markdown("**UT Mula (dd/mm/yyyy)**")
                ut_mula_str = st.text_input("", value="", key="ut_mula", label_visibility="collapsed")
            with u2:
                st.markdown("**UT Akhir (dd/mm/yyyy)**")
                ut_akhir_str = st.text_input("", value="", key="ut_akhir", label_visibility="collapsed")

with right_col:
    with st.container(border=True):
        st.markdown("### Muat Naik Fail")

        st.markdown("**Agenda JK OSC (.docx)**")
        agenda_file = st.file_uploader("", type=["docx"], key="agenda_docx", label_visibility="collapsed")

        proceed_without_agenda = st.checkbox(
            "Teruskan tanpa Agenda",
            value=False,
            help="Tick jika agenda belum diterima. Jika tick, sistem jana tanpa tapisan agenda.",
        )

        st.markdown("**Kertas Maklumat (Excel) — SPU/SPS/SPT (boleh upload 1 atau 2 fail setiap daerah)**")
        st.caption("Nota: Hujung/awal tahun boleh jadi 2 fail (tahun lama + tahun baru). Sistem akan gabungkan.")

        st.markdown("**SPU (maks 2 fail)**")
        spu_files = st.file_uploader("", type=["xlsx", "xlsm"], key="spu_multi", label_visibility="collapsed", accept_multiple_files=True)

        st.markdown("**SPS (maks 2 fail)**")
        sps_files = st.file_uploader("", type=["xlsx", "xlsm"], key="sps_multi", label_visibility="collapsed", accept_multiple_files=True)

        st.markdown("**SPT (maks 2 fail)**")
        spt_files = st.file_uploader("", type=["xlsx", "xlsm"], key="spt_multi", label_visibility="collapsed", accept_multiple_files=True)

mid = st.columns([1, 0.55, 1])[1]
with mid:
    gen = st.button("JANA LAMPIRAN G", type="primary", disabled=st.session_state.running)


# ============================================================
# ACTION (GENERATE)
# ============================================================
if gen:
    st.session_state.running = True
    try:
        with st.spinner("Sedang jana Lampiran G... Sila tunggu."):
            km_start = _parse_ddmmyyyy(km_mula_str)
            km_end = _parse_ddmmyyyy(km_akhir_str)

            if ut_enabled:
                ut_start = _parse_ddmmyyyy(ut_mula_str)
                ut_end = _parse_ddmmyyyy(ut_akhir_str)
            else:
                ut_start = None
                ut_end = None

            # Validasi Agenda
            agenda_enabled = True
            if proceed_without_agenda:
                agenda_enabled = False
            else:
                if not agenda_file:
                    st.error("Sila upload Agenda JK OSC (.docx) atau tick 'Teruskan tanpa Agenda'.")
                    st.stop()

            # Validasi kertas maklumat (multi upload)
            spu_files = spu_files or []
            sps_files = sps_files or []
            spt_files = spt_files or []

            if len(spu_files) == 0 or len(sps_files) == 0 or len(spt_files) == 0:
                st.error("Sila upload sekurang-kurangnya 1 fail untuk setiap SPU, SPS dan SPT.")
                st.stop()

            if len(spu_files) > 2 or len(sps_files) > 2 or len(spt_files) > 2:
                st.error("Maksimum 2 fail dibenarkan bagi setiap daerah (SPU/SPS/SPT).")
                st.stop()

            # Validasi tarikh KM/UT
            if km_start is None or km_end is None:
                st.error("Sila isi tarikh KM Mula dan KM Akhir dalam format dd/mm/yyyy.")
                st.stop()
            if km_start > km_end:
                st.error("KM Mula tidak boleh lebih besar daripada KM Akhir.")
                st.stop()

            if ut_enabled:
                if ut_start is None or ut_end is None:
                    st.error("Sila isi tarikh UT Mula dan UT Akhir dalam format dd/mm/yyyy.")
                    st.stop()
                if ut_start > ut_end:
                    st.error("UT Mula tidak boleh lebih besar daripada UT Akhir.")
                    st.stop()

            # Read agenda (jika digunakan)
            if agenda_enabled:
                agenda_bytes = agenda_file.read()
                agenda_osc_set = parse_agenda_docx(agenda_bytes)
            else:
                agenda_osc_set = set()

            # Read all excel files
            rows: List[dict] = []

            for f in spu_files:
                rows += read_kertas_excel(f.read(), "SPU")

            for f in sps_files:
                rows += read_kertas_excel(f.read(), "SPS")

            for f in spt_files:
                rows += read_kertas_excel(f.read(), "SPT")

            rows = enrich_rows(rows)

            cat1, cat2, cat3, cat4, cat5 = build_categories(
                rows=rows,
                agenda_osc_set=agenda_osc_set,
                km_start=km_start,
                km_end=km_end,
                ut_start=ut_start if ut_enabled else km_start,
                ut_end=ut_end if ut_enabled else km_end,
                ut_enabled=ut_enabled,
                agenda_enabled=agenda_enabled,
            )

            doc_bytes = build_word_doc(
                meeting_info=meeting_info.strip() if meeting_info.strip() else "JK OSC",
                km_start=km_start,
                km_end=km_end,
                ut_start=ut_start if ut_enabled else km_start,
                ut_end=ut_end if ut_enabled else km_end,
                cat1=cat1,
                cat2=cat2,
                cat3=cat3,
                cat4=cat4,
                cat5=cat5,
                ut_enabled=ut_enabled,
            )

            st.success("Lampiran G berjaya dijana.")
            st.download_button(
                "Muat turun Lampiran G (Word)",
                data=doc_bytes,
                file_name="Lampiran_G.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

            with st.expander("Ringkasan (untuk semakan cepat)"):
                st.write({
                    "Kategori 1": len(cat1),
                    "Kategori 2": len(cat2) if ut_enabled else 0,
                    "Kategori 3": len(cat3),
                    "Kategori 4": len(cat4),
                    "Kategori 5": len(cat5),
                    "Agenda digunakan?": "YA" if agenda_enabled else "TIDAK (Teruskan tanpa Agenda)",
                })
    finally:
        st.session_state.running = False
