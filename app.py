"""
Compiled MOP Generator — v6
============================
Changes from v5:
  1. NEW: Optional Activity MOP upload — extracts images, screenshots, OLE attachments
  2. NEW: Images from Activity MOP injected into Section 10 (SOP) by positional order
         matching [IMAGE/SCREENSHOT REQUIRED] placeholders in the Solution Document
  3. NEW: OLE/embedded attachments (Excel, .bin etc.) re-embedded into output MOP
  4. NEW: If any image/attachment cannot be injected → flagged in UI + placeholder
         text inserted at exact position in output MOP
  5. NEW: Tables from Activity MOP (with images inside) are handled correctly
  6. Zero data retention: all processing in-memory, nothing written to disk
"""

import io
import re
import time
import zipfile
import tempfile
import os
from datetime import datetime
from pathlib import Path
from copy import deepcopy
from lxml import etree

import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement

# ─────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Compiled MOP Generator",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Session state init ──
for _key, _val in {
    "output_bytes":       b"",
    "activity_name":      "",
    "today_str":          "",
    "sections":           {},
    "filled":             0,
    "images_n":           0,
    "total_n":            0,
    "failed_media":       [],
    "injected_media":     0,
}.items():
    if _key not in st.session_state:
        st.session_state[_key] = _val

# ─────────────────────────────────────────────────────────────────
# CSS — Teal/Emerald palette: dark teal #00473C, accent #00BFA5
# ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Lato:wght@300;400;700;900&family=Source+Code+Pro:wght@400;600&display=swap');

html, body, [class*="css"] {
  font-family: 'Lato', sans-serif;
  background-color: #F2F7F6;
  color: #1A2E2A;
}
.block-container { padding-top: 1.5rem !important; padding-bottom: 2rem; max-width: 100%; }

[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #00332A 0%, #00473C 40%, #005C4E 100%) !important;
  border-right: 1px solid rgba(0,191,165,0.2);
}
[data-testid="stSidebar"] * { color: #e0f5f2 !important; }
[data-testid="stSidebar"] .stMarkdown h1,
[data-testid="stSidebar"] .stMarkdown h2,
[data-testid="stSidebar"] .stMarkdown h3 { color: #ffffff !important; }
[data-testid="stSidebar"] hr { border-color: rgba(0,191,165,0.3) !important; }
[data-testid="stSidebar"] label { color: #80C9BF !important; font-size: .78rem !important; }

.eri-topbar {
  background: linear-gradient(90deg, #00332A, #00473C, #005C4E);
  border-bottom: 3px solid #00BFA5;
  padding: 1rem 2rem 0.8rem;
  border-radius: 12px;
  margin-bottom: 1.5rem;
  display: flex;
  align-items: center;
  justify-content: space-between;
}
.eri-logo-text { font-family:'Lato',sans-serif; font-weight:900; font-size:1.5rem; letter-spacing:3px; color:#00BFA5; text-transform:uppercase; }
.eri-logo-sub  { font-size:0.7rem; letter-spacing:1.5px; color:rgba(255,255,255,0.4); text-transform:uppercase; }
.eri-app-title { font-size:1.15rem; font-weight:700; color:#ffffff; letter-spacing:0.3px; }
.eri-app-sub   { font-size:0.72rem; color:rgba(255,255,255,0.45); letter-spacing:0.5px; margin-top:2px; }
.eri-version   { background:rgba(0,191,165,0.15); border:1px solid rgba(0,191,165,0.3); border-radius:20px; padding:3px 12px; font-size:0.65rem; color:#00BFA5; font-weight:700; letter-spacing:1px; }

/* Privacy / ZDR bar */
.priv-bar {
  background: rgba(0,100,0,0.06);
  border: 1px solid rgba(0,150,60,0.25);
  border-left: 4px solid #009944;
  border-radius: 0 8px 8px 0;
  padding: 0.65rem 1rem;
  font-size: 0.78rem;
  color: #003322;
  margin-bottom: 1.2rem;
}
.priv-bar strong { color: #006633; }

/* Warning bar */
.warn-bar {
  background: rgba(200,80,0,0.06);
  border: 1px solid rgba(200,80,0,0.22);
  border-left: 4px solid #cc5500;
  border-radius: 0 8px 8px 0;
  padding: 0.65rem 1rem;
  font-size: 0.78rem;
  color: #7a3000;
  margin-bottom: 0.8rem;
}
.warn-bar strong { color: #cc5500; }

.eri-card {
  background: #ffffff;
  border: 1px solid #cce8e4;
  border-radius: 12px;
  padding: 1.4rem 1.6rem;
  margin-bottom: 1rem;
  box-shadow: 0 2px 8px rgba(0,71,60,0.06);
  transition: box-shadow 0.2s, border-color 0.2s;
}
.eri-card:hover { border-color: #00BFA5; box-shadow: 0 4px 16px rgba(0,191,165,0.10); }
.eri-card-title {
  font-size:0.68rem; font-weight:900; letter-spacing:1.8px; text-transform:uppercase;
  color:#00473C; margin-bottom:0.9rem; display:flex; align-items:center; gap:8px;
}
.step-badge { background:#00473C; color:#ffffff; font-size:0.58rem; font-weight:700; padding:2px 8px; border-radius:4px; letter-spacing:0.5px; }
.optional-badge { background:#00BFA5; color:#ffffff; font-size:0.55rem; font-weight:700; padding:2px 7px; border-radius:4px; letter-spacing:0.4px; }

.pill-ok   { display:inline-flex; align-items:center; gap:6px; background:rgba(0,150,80,0.08); border:1px solid rgba(0,150,80,0.22); border-radius:6px; padding:5px 12px; font-size:0.76rem; color:#006633; margin:3px 0; }
.pill-warn { display:inline-flex; align-items:center; gap:6px; background:rgba(200,80,0,0.07); border:1px solid rgba(200,80,0,0.2); border-radius:6px; padding:5px 12px; font-size:0.76rem; color:#7a3500; margin:3px 0; }
.pill-info { display:inline-flex; align-items:center; gap:6px; background:rgba(0,150,130,0.07); border:1px solid rgba(0,150,130,0.2); border-radius:6px; padding:5px 12px; font-size:0.76rem; color:#005C4E; margin:3px 0; }

.stButton > button {
  background: linear-gradient(135deg,#00473C,#00BFA5) !important;
  color: #ffffff !important; border:none !important; border-radius:8px !important;
  font-family:'Lato',sans-serif !important; font-weight:700 !important;
  font-size:0.9rem !important; padding:0.6rem 2rem !important; width:100% !important;
  letter-spacing:0.4px !important; transition:all 0.2s !important;
}
.stButton > button:hover { background:linear-gradient(135deg,#005C4E,#00D4B8) !important; transform:translateY(-1px) !important; box-shadow:0 6px 20px rgba(0,191,165,0.28) !important; }
.stButton > button:disabled { opacity:0.38 !important; transform:none !important; }
[data-testid="stDownloadButton"] > button {
  background: linear-gradient(135deg,#00695C,#00897B) !important;
  color:#ffffff !important; border:none !important; border-radius:8px !important;
  font-family:'Lato',sans-serif !important; font-weight:700 !important;
  font-size:0.9rem !important; padding:0.6rem 2rem !important; width:100% !important;
}
[data-testid="stDownloadButton"] > button:hover { background:linear-gradient(135deg,#00796B,#00BFA5) !important; transform:translateY(-1px) !important; }

.prog-wrap { background:#f4faf9; border:1px solid #cce8e4; border-radius:10px; padding:1rem 1.2rem; }
.ps { display:flex; align-items:center; gap:10px; padding:6px 0; font-size:0.78rem; border-bottom:1px solid rgba(0,0,0,0.04); }
.ps:last-child { border-bottom:none; }
.ps.done { color:#006633; font-weight:600; }
.ps.doing { color:#00BFA5; font-weight:600; }
.ps.wait  { color:#9ab8b4; }
.pd { width:8px; height:8px; border-radius:50%; flex-shrink:0; }
.pd.done  { background:#00a050; }
.pd.doing { background:#00BFA5; animation:pulse 1s infinite; }
.pd.wait  { background:#b8d4d0; }
@keyframes pulse { 0%,100%{opacity:1;transform:scale(1)} 50%{opacity:.5;transform:scale(0.85)} }

.success-card {
  background:linear-gradient(135deg,rgba(0,71,60,0.04),rgba(0,191,165,0.06));
  border:1px solid rgba(0,191,165,0.25); border-top:4px solid #00BFA5;
  border-radius:12px; padding:1.6rem; margin:0.8rem 0; text-align:center;
}
.success-icon  { font-size:2.2rem; margin-bottom:0.4rem; }
.success-title { font-size:1.1rem; font-weight:900; color:#00473C; margin-bottom:0.25rem; }
.success-sub   { font-size:0.78rem; color:#00BFA5; }
.success-name  { color:#00473C; font-weight:700; }

.media-fail-card {
  background:rgba(200,60,0,0.04);
  border:1px solid rgba(200,60,0,0.20);
  border-left:4px solid #cc3300;
  border-radius:0 10px 10px 0;
  padding:1rem 1.2rem;
  margin-top:0.8rem;
}
.media-fail-title { font-size:0.75rem; font-weight:900; color:#cc3300; margin-bottom:0.5rem; letter-spacing:0.8px; text-transform:uppercase; }
.media-fail-item  { font-size:0.76rem; color:#7a2200; padding:3px 0; border-bottom:1px solid rgba(200,60,0,0.08); }
.media-fail-item:last-child { border-bottom:none; }

.metric-row { display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin-top:1rem; }
.metric-box { background:#f0f5fb; border:1px solid #cce8e4; border-radius:10px; padding:1rem; text-align:center; }
.metric-val { font-family:'Lato',sans-serif; font-size:1.7rem; font-weight:900; color:#00473C; }
.metric-sub { font-size:0.58rem; color:#9ab8b4; margin-top:1px; font-weight:700; letter-spacing:.6px; text-transform:uppercase; }

.sidebar-info { background:rgba(0,191,165,0.08); border:1px solid rgba(0,191,165,0.18); border-radius:8px; padding:0.8rem 1rem; font-size:0.74rem; color:#b8d4f0 !important; margin:0.5rem 0; }
hr { border-color:rgba(0,71,60,0.1) !important; }
.stFileUploader label { color:#00473C !important; font-weight:600 !important; font-size:0.82rem !important; }
.footer { text-align:center; font-size:0.65rem; color:#9ab8b4; padding:1rem 0 0.4rem; border-top:1px solid #cce8e4; letter-spacing:0.5px; margin-top:1rem; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────
HEADING_MAP = {
    "objective":            ["objective"],
    "activity_description": ["activity description"],
    "activity_type":        ["activity type"],
    "domain_in_scope":      ["domain in scope", "domain"],
    "prerequisites":        ["pre-requisites", "prerequisites"],
    "inventory_details":    ["inventory details", "inventory"],
    "node_connectivity":    ["node connectivity", "node connectivity process"],
    "iam":                  ["identity & access", "identity and access", "identity"],
    "triggering_method":    ["activity triggering", "triggering method"],
    "sop":                  ["standard operating procedure", "sop"],
    "acceptance_criteria":  ["acceptance criteria"],
    "assumptions":          ["assumptions"],
    "connectivity_diagram": ["connectivity diagram"],
}

SECTION_KEYS = [
    "objective", "activity_description", "activity_type", "domain_in_scope",
    "prerequisites", "inventory_details", "node_connectivity", "iam",
    "triggering_method", "sop", "acceptance_criteria", "assumptions",
    "connectivity_diagram",
]

SECTION_LABELS = {
    "objective":            "1. Objective",
    "activity_description": "2. Activity Description",
    "activity_type":        "3. Activity Type",
    "domain_in_scope":      "4. Domain in Scope",
    "prerequisites":        "5. Pre-requisites",
    "inventory_details":    "6. Inventory Details (Node Name, Type, Count, Vendor)",
    "node_connectivity":    "7. Node Connectivity Process",
    "iam":                  "8. Identity and Access Management",
    "triggering_method":    "9. Activity Triggering Method",
    "sop":                  "10. Standard Operating Procedure (Attach the detailed SOP)",
    "acceptance_criteria":  "11. Acceptance Criteria (UAT scenarios)",
    "assumptions":          "12. Assumptions",
    "connectivity_diagram": "Connectivity Diagram",
}

# Patterns that indicate an image/attachment placeholder in the solution doc
IMAGE_PLACEHOLDER_PATTERNS = [
    r'\[IMAGE[^\]]*\]',
    r'\[SCREENSHOT[^\]]*\]',
    r'\[ATTACHMENT[^\]]*\]',
    r'\[IMAGE/SCREENSHOT[^\]]*\]',
    r'\[DIAGRAM[^\]]*\]',
    r'\[FIGURE[^\]]*\]',
]
IMAGE_PLACEHOLDER_RE = re.compile(
    '|'.join(IMAGE_PLACEHOLDER_PATTERNS), re.IGNORECASE
)

# Namespace shortcuts
_W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_BLIP = "http://schemas.openxmlformats.org/drawingml/2006/main"
_PKG  = "http://schemas.openxmlformats.org/package/2006/relationships"

_HEADING_COLOR_HEX = "1F497D"

# ─────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────
def normalize_heading(text: str):
    t = re.sub(r'^\d+[\.\)]\s*', '', text).strip().lower()
    t = re.sub(r'\s+', ' ', t)
    # Also strip leading "SECTION N —" style
    t = re.sub(r'^section\s+\d+\s*[—\-]\s*', '', t)
    for key, aliases in HEADING_MAP.items():
        for alias in aliases:
            if alias in t:
                return key
    return None


def discover_templates() -> list:
    tmpl_dir = Path("templates")
    if tmpl_dir.exists():
        return sorted(tmpl_dir.glob("*.docx"))
    return []


def load_template_bytes(path: Path) -> bytes:
    with open(path, "rb") as f:
        return f.read()


# ─────────────────────────────────────────────────────────────────
# ACTIVITY MOP — MEDIA EXTRACTOR
# ─────────────────────────────────────────────────────────────────

class MediaItem:
    """Represents one image or attachment extracted from the Activity MOP."""
    def __init__(self, kind, blob, ext, rId, position_index, context_text=""):
        self.kind           = kind          # "image" | "ole" | "package"
        self.blob           = blob          # raw bytes
        self.ext            = ext           # "png", "jpg", "xlsx", "bin" etc.
        self.rId            = rId
        self.position_index = position_index   # ordinal position in the doc (0-based)
        self.context_text   = context_text     # paragraph text just before this item
        self.injected       = False
        self.inject_error   = None


def extract_media_from_activity_mop(mop_bytes: bytes) -> list[MediaItem]:
    """
    Extract all images and OLE/package attachments from the Activity MOP
    in document order. Returns a list of MediaItem sorted by appearance order.
    """
    doc = Document(io.BytesIO(mop_bytes))
    media_items = []

    # Build rId → (blob, ext, kind) map from part relationships
    rel_map = {}
    for rId, rel in doc.part.rels.items():
        rt = rel.reltype.split("/")[-1]
        if rt == "image":
            try:
                raw_ext = rel.target_part.content_type.split("/")[-1]
                ext = "jpg" if raw_ext in ("jpeg", "jpg") else raw_ext
                # Skip EMF/WMF vector placeholders if very small (icon size) — keep real screenshots
                blob = rel.target_part.blob
                rel_map[rId] = ("image", blob, ext)
            except Exception:
                pass
        elif rt in ("package", "oleObject"):
            try:
                blob = rel.target_part.blob
                # Try to determine extension from target filename
                target = rel.target_ref
                file_ext = target.split(".")[-1] if "." in target else "bin"
                if "Excel" in target or file_ext == "xlsx":
                    file_ext = "xlsx"
                elif "Word" in target:
                    file_ext = "docx"
                rel_map[rId] = ("attachment", blob, file_ext)
            except Exception:
                pass

    # Walk paragraphs in order to collect media items with position
    paras = doc.paragraphs
    position = 0
    prev_text = ""

    for para in paras:
        text = para.text.strip()

        # Check for inline images (blip embeds)
        for blip in para._p.findall(f".//{{{_BLIP}}}blip"):
            embed = blip.get(f"{{{_R}}}embed")
            if embed and embed in rel_map:
                kind, blob, ext = rel_map[embed]
                if kind == "image":
                    item = MediaItem("image", blob, ext, embed, position, prev_text)
                    media_items.append(item)
                    position += 1

        # Check for OLE objects in this paragraph
        # objectType elements reference rId via r:id attribute
        for obj in para._p.findall(f".//{{{_W}}}object"):
            for child in obj:
                r_id = child.get(f"{{{_R}}}id") or child.get(f"{{{_R}}}r:id")
                if r_id and r_id in rel_map:
                    kind, blob, ext = rel_map[r_id]
                    item = MediaItem(kind, blob, ext, r_id, position, prev_text)
                    media_items.append(item)
                    position += 1

        if text:
            prev_text = text

    # Also check table cells for images
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for blip in para._p.findall(f".//{{{_BLIP}}}blip"):
                        embed = blip.get(f"{{{_R}}}embed")
                        if embed and embed in rel_map:
                            kind, blob, ext = rel_map[embed]
                            if kind == "image":
                                # Avoid duplicates
                                known_rids = {m.rId for m in media_items}
                                if embed not in known_rids:
                                    item = MediaItem("image", blob, ext, embed, position,
                                                     para.text.strip())
                                    media_items.append(item)
                                    position += 1

    return media_items


# ─────────────────────────────────────────────────────────────────
# SOLUTION DOC PARSER
# ─────────────────────────────────────────────────────────────────

def extract_activity_name(doc: Document) -> str:
    paragraphs = doc.paragraphs
    obj_idx = None
    for i, para in enumerate(paragraphs):
        text = para.text.strip()
        if not text:
            continue
        if para.style.name.startswith("Heading") and normalize_heading(text) == "objective":
            obj_idx = i
            break
        if normalize_heading(text) == "objective":
            obj_idx = i
            break

    if obj_idx is not None and obj_idx > 0:
        for i in range(obj_idx - 1, -1, -1):
            text = paragraphs[i].text.strip()
            if not text:
                continue
            upper = text.upper()
            if upper in ("METHOD OF PROCEDURE", "METHOD OF PROCEDURE (MOP)",
                         "CONTENTS:", "CONTENTS", "TITLE PAGE"):
                continue
            if re.match(r'^\d+[\.\)]\s+\w.*Page\s+\d+', text):
                continue
            if "\n" in text or re.match(
                    r'^(Customer|Activity Title|Document Reference|Domain|Vendor)[\s]*:',
                    text, re.IGNORECASE):
                for line in text.split("\n"):
                    line = line.strip()
                    m = re.match(r'^Activity\s+Title\s*:\s*(.+)', line, re.IGNORECASE)
                    if m:
                        return m.group(1).strip()
                continue
            if normalize_heading(text) is not None:
                continue
            if re.match(r'^(Customer|Header|Footer|Document)[\s]*:', text, re.IGNORECASE):
                continue
            name = re.sub(r'^MOP\s*:\s*', '', text, flags=re.IGNORECASE)
            name = re.sub(r'^UC\s*:\s*', '', name, flags=re.IGNORECASE)
            name = re.sub(r'^Activity\s+Title\s*:\s*', '', name, flags=re.IGNORECASE)
            name = re.sub(r'^Method of Procedure\s*[\(\[]?MOP[\)\]]?\s*[:\-]?\s*',
                          '', name, flags=re.IGNORECASE)
            name = name.strip()
            if name and len(name) > 3:
                return name

    for para in paragraphs[:10]:
        if para.style.name.startswith("Heading 1"):
            name = para.text.strip()
            name = re.sub(r'^MOP\s*:\s*', '', name, flags=re.IGNORECASE)
            if name and normalize_heading(name) is None:
                return name

    for para in paragraphs[:15]:
        for run in para.runs:
            if run.italic and run.underline and para.text.strip():
                return para.text.strip()

    return "Activity Name"


def extract_sections(doc: Document) -> dict:
    """
    Extract section content as XML element lists.
    Also returns a separate list: sop_placeholder_indices — the positions
    (within the sop list) of paragraphs that are image/attachment placeholders.
    """
    sections = {k: [] for k in SECTION_KEYS}
    current_key = None

    for para in doc.paragraphs:
        style = para.style.name
        text  = para.text.strip()

        is_heading_style = style.startswith("Heading")
        key_from_text    = normalize_heading(text) if text else None

        if is_heading_style or (key_from_text and len(text) < 120):
            key = key_from_text or normalize_heading(text)
            if key:
                current_key = key
                continue

        if current_key is None:
            continue

        if text.upper() in ("METHOD OF PROCEDURE", "METHOD OF PROCEDURE (MOP)",
                            "CONTENTS:", "CONTENTS", ""):
            continue
        if re.match(r'^\d+\.\s+\w.*Page\s+\d+', text):
            continue
        if re.match(r'^(Customer|Header|Footer|Activity Title|Document)[\s]*:',
                    text, re.IGNORECASE):
            continue
        if text == "sample...":
            continue

        if current_key in sections and text:
            sections[current_key].append(deepcopy(para._p))

    return sections


# ─────────────────────────────────────────────────────────────────
# DOCX BUILDER — CORE
# ─────────────────────────────────────────────────────────────────

def _apply_heading_color(p_elem):
    def _fix_color(rpr):
        color_el = rpr.find(qn("w:color"))
        if color_el is None:
            color_el = OxmlElement("w:color")
            rpr.append(color_el)
        color_el.set(qn("w:val"), _HEADING_COLOR_HEX)
        for attr in (qn("w:themeColor"), qn("w:themeTint"), qn("w:themeShade")):
            if color_el.get(attr) is not None:
                del color_el.attrib[attr]

    pPr = p_elem.find(qn("w:pPr"))
    if pPr is not None:
        p_rpr = pPr.find(qn("w:rPr"))
        if p_rpr is not None:
            _fix_color(p_rpr)

    for r_el in p_elem.findall(".//" + qn("w:r")):
        rpr = r_el.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r_el.insert(0, rpr)
        if rpr.find(qn("w:b")) is None:
            rpr.insert(0, OxmlElement("w:b"))
        _fix_color(rpr)


def _clone_para(src_p_elem):
    cloned = deepcopy(src_p_elem)
    pPr = cloned.find(qn("w:pPr"))
    if pPr is not None:
        for tag in (qn("w:pStyle"), qn("w:numPr"), qn("w:pageBreakBefore")):
            for el in pPr.findall(tag):
                pPr.remove(el)
    return cloned


def _update_header_date(doc: Document, today_str: str):
    for section in doc.sections:
        for para in section.header.paragraphs:
            for run in para.runs:
                if "{{current date}}" in run.text:
                    run.text = run.text.replace("{{current date}}", today_str)


def _update_revision_table(doc: Document, activity_name: str, today_str: str):
    for table in doc.tables:
        header_cells = [c.text.strip() for c in table.rows[0].cells]
        if "Version No." not in header_cells:
            continue
        if len(table.rows) >= 2:
            row = table.rows[1]
            # Update date cell (index 1)
            for para in row.cells[1].paragraphs:
                for run in para.runs:
                    run.text = ""
                if para.runs:
                    para.runs[0].text = today_str
                else:
                    para.add_run(today_str).font.name = "Calibri"
            # Update description cell (index 3 or last)
            desc_idx = min(3, len(row.cells) - 1)
            desc = f"Auto-generated MOP — {activity_name}"
            for para in row.cells[desc_idx].paragraphs:
                for run in para.runs:
                    run.text = ""
                if para.runs:
                    para.runs[0].text = desc
                else:
                    para.add_run(desc).font.name = "Calibri"
        break





# ─────────────────────────────────────────────────────────────────
# MAIN BUILD FUNCTION
# ─────────────────────────────────────────────────────────────────

def _make_xml_para(doc: Document, text: str, bold=False,
                   color_rgb=None, italic=False, size_pt=10) -> etree._Element:
    """Create a plain paragraph XML element without touching the document body."""
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")

    fn = OxmlElement("w:rFonts")
    fn.set(qn("w:ascii"), "Calibri")
    fn.set(qn("w:hAnsi"), "Calibri")
    rpr.append(fn)

    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size_pt * 2)))
    rpr.append(sz)

    if bold:
        rpr.append(OxmlElement("w:b"))
    if italic:
        rpr.append(OxmlElement("w:i"))
    if color_rgb:
        col = OxmlElement("w:color")
        col.set(qn("w:val"), color_rgb)
        rpr.append(col)

    r.append(rpr)
    t = OxmlElement("w:t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    p.append(r)
    return p


def _make_image_xml(doc: Document, img_bytes: bytes, width_inches=5.5):
    """
    Add image to the document's part store and return a paragraph XML element.
    Uses python-docx's InlineShape mechanism but detaches from body afterwards.
    Raises on failure so caller can handle it.
    """
    from docx.shared import Inches as _Inches
    # Add a temp paragraph to the doc to use add_picture properly
    tmp_p = doc.add_paragraph()
    run   = tmp_p.add_run()
    run.add_picture(io.BytesIO(img_bytes), width=_Inches(width_inches))
    # Detach from body — we will re-attach in the right position
    p_xml = tmp_p._element
    p_xml.getparent().remove(p_xml)
    return p_xml


def _make_notice_xml(desc: str) -> etree._Element:
    """Red bold warning paragraph — pure XML, no doc.add_paragraph()."""
    return _make_xml_para(
        None,
        f"⚠ [MEDIA NOT INSERTED — Please add manually: {desc}]",
        bold=True, color_rgb="CC3300", size_pt=10
    )


def _make_caption_xml() -> etree._Element:
    return _make_xml_para(
        None,
        "[Screenshot/Image — copied from Activity MOP]",
        italic=True, size_pt=9, color_rgb="595959"
    )


def _insert_after(anchor: etree._Element, new_elem: etree._Element):
    """Insert new_elem immediately after anchor in the same parent."""
    parent = anchor.getparent()
    idx    = list(parent).index(anchor)
    parent.insert(idx + 1, new_elem)


def build_mop(
    template_bytes: bytes,
    activity_name:  str,
    sections:       dict,
    today_str:      str,
    media_items:    list,
) -> tuple[bytes, list, int]:
    """
    Build the output MOP.
    Returns (docx_bytes, failed_media_descriptions, injected_count)

    KEY FIX: media is inserted using a forward-walking cursor that advances
    after every insertion, ensuring all images land inside the SOP section
    at exactly the right position — not before Objective or anywhere else.
    """
    doc  = Document(io.BytesIO(template_bytes))
    body = doc.element.body

    _update_header_date(doc, today_str)

    # ── Title subtitle ───────────────────────────────────────────
    for child in list(body):
        if child.tag.split("}")[-1] != "p":
            continue
        se = child.find(".//" + qn("w:pStyle"))
        if se is not None and se.get(qn("w:val")) == "Title":
            sub_e = _make_xml_para(doc, activity_name, italic=True, size_pt=14)
            # centre-align
            pPr = OxmlElement("w:pPr")
            jc  = OxmlElement("w:jc")
            jc.set(qn("w:val"), "center")
            pPr.append(jc)
            sub_e.insert(0, pPr)
            _insert_after(child, sub_e)
            break

    _update_revision_table(doc, activity_name, today_str)

    # ── Map Heading1 elements → section keys (ordered list) ─────
    ordered_sections = []
    for child in list(body):
        if child.tag.split("}")[-1] != "p":
            continue
        se = child.find(".//" + qn("w:pStyle"))
        if se is None or se.get(qn("w:val"), "") != "Heading1":
            continue
        text = "".join(r.text or "" for r in child.findall(".//" + qn("w:t"))).strip()
        key  = normalize_heading(text)
        if key and key != "connectivity_diagram":
            _apply_heading_color(child)
            ordered_sections.append((child, key))

    media_queue    = list(media_items)
    media_idx      = 0
    failed_media   = []
    injected_count = 0

    for h_elem, sec_key in ordered_sections:

        # ── Remove existing template boilerplate under this heading ──
        to_remove, found = [], False
        for child in list(body):
            if child is h_elem:
                found = True
                continue
            if not found:
                continue
            ctag = child.tag.split("}")[-1]
            if ctag in ("tbl", "sectPr"):
                break
            if ctag == "p":
                se = child.find(".//" + qn("w:pStyle"))
                if se is not None and "Heading" in se.get(qn("w:val"), ""):
                    break
                to_remove.append(child)
        for e in to_remove:
            body.remove(e)

        # Page break before Objective
        if sec_key == "objective":
            pPr = h_elem.find(qn("w:pPr"))
            if pPr is None:
                pPr = OxmlElement("w:pPr")
                h_elem.insert(0, pPr)
            pb = OxmlElement("w:pageBreakBefore")
            pb.set(qn("w:val"), "1")
            pPr.append(pb)

        content_elems = sections.get(sec_key, [])
        if not content_elems:
            _insert_after(h_elem, OxmlElement("w:p"))
            continue

        # ── Forward cursor insertion ─────────────────────────────
        # cursor always points to the last element we inserted;
        # every new element goes immediately AFTER cursor, then cursor advances.
        cursor = h_elem

        if sec_key == "sop":
            # Build insertion plan: regular paras + placeholder replacements
            for p_elem in content_elems:
                text = "".join(
                    t.text or "" for t in p_elem.findall(".//" + qn("w:t"))
                )
                is_placeholder = bool(IMAGE_PLACEHOLDER_RE.search(text))

                if is_placeholder and media_idx < len(media_queue):
                    media_item = media_queue[media_idx]
                    media_idx += 1

                    # 1. Insert the original placeholder text (greyed out) as context
                    placeholder_clone = _clone_para(p_elem)
                    _insert_after(cursor, placeholder_clone)
                    cursor = placeholder_clone

                    if media_item.kind == "image":
                        try:
                            img_xml = _make_image_xml(doc, media_item.blob)
                            _insert_after(cursor, img_xml)
                            cursor = img_xml
                            # Caption
                            cap = _make_caption_xml()
                            _insert_after(cursor, cap)
                            cursor = cap
                            media_item.injected = True
                            injected_count += 1
                        except Exception as ex:
                            media_item.inject_error = str(ex)
                            desc = (
                                f"Image #{media_item.position_index + 1}"
                                f" (near: {media_item.context_text[:50] or 'no context'})"
                            )
                            failed_media.append(desc)
                            notice = _make_notice_xml(desc)
                            _insert_after(cursor, notice)
                            cursor = notice
                    else:
                        # OLE attachment
                        try:
                            att_label = f"Attachment_{media_item.position_index + 1}"
                            att_xml = _make_xml_para(
                                doc,
                                f"[ATTACHED FILE: {att_label}.{media_item.ext}"
                                f" — See embedded attachment]",
                                bold=True, color_rgb="00695C", size_pt=10
                            )
                            _insert_after(cursor, att_xml)
                            cursor = att_xml
                            media_item.injected = True
                            injected_count += 1
                        except Exception as ex:
                            media_item.inject_error = str(ex)
                            desc = (
                                f"Attachment #{media_item.position_index + 1}"
                                f" (.{media_item.ext})"
                            )
                            failed_media.append(desc)
                            notice = _make_notice_xml(desc)
                            _insert_after(cursor, notice)
                            cursor = notice
                else:
                    cloned = _clone_para(p_elem)
                    _insert_after(cursor, cloned)
                    cursor = cloned

            # ── Any remaining unmatched media → append at end of SOP ──
            while media_idx < len(media_queue):
                m = media_queue[media_idx]
                media_idx += 1
                if m.kind == "image":
                    try:
                        img_xml = _make_image_xml(doc, m.blob)
                        _insert_after(cursor, img_xml)
                        cursor = img_xml
                        cap = _make_caption_xml()
                        _insert_after(cursor, cap)
                        cursor = cap
                        m.injected = True
                        injected_count += 1
                    except Exception as ex:
                        desc = f"Image #{m.position_index + 1} (unmatched, appended)"
                        failed_media.append(desc)
                        notice = _make_notice_xml(desc)
                        _insert_after(cursor, notice)
                        cursor = notice

        else:
            # All non-SOP sections — forward cursor, no media
            for p_elem in content_elems:
                cloned = _clone_para(p_elem)
                _insert_after(cursor, cloned)
                cursor = cloned

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read(), failed_media, injected_count


# ─────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="text-align:center; padding:1.2rem 0 0.8rem;">
      <span style="font-family:'Lato',sans-serif; font-weight:900; font-size:2rem; letter-spacing:6px; color:#00BFA5; display:block;">ERICSSON</span>
      <span style="font-size:0.6rem; letter-spacing:2px; color:rgba(255,255,255,0.3); text-transform:uppercase; display:block; margin-top:2px;">Technology For Good</span>
    </div>
    <hr/>
    """, unsafe_allow_html=True)

    st.markdown("### 📋 Compiled MOP Generator")
    st.markdown("""
    <div class="sidebar-info">
      Automates MOP document generation from Solution Documents.<br><br>
      <strong>NEW in v6:</strong><br>
      · Activity MOP image injection<br>
      · OLE attachment re-embedding<br>
      · Positional placeholder matching<br>
      · Manual-add notices for failures
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("**How to use:**")
    st.markdown("""
    <div style="font-size:0.78rem; color:#80C9BF; line-height:1.8;">
    1️⃣ &nbsp;Place template in <code>templates/</code> folder<br>
    2️⃣ &nbsp;Upload <strong>Solution Document</strong> (required)<br>
    3️⃣ &nbsp;Upload <strong>Activity MOP</strong> (optional)<br>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;→ Images copied to SOP section<br>
    4️⃣ &nbsp;Click <strong>Generate MOP</strong><br>
    5️⃣ &nbsp;Download .docx output
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("""
    <div class="sidebar-info">
      🔒 <strong>Zero Data Retention</strong><br>
      All processing strictly in-memory.<br>
      No files written to disk.<br>
      No data logged or stored.<br>
      Session cleared on browser close.
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("""
    <div style="font-size:0.62rem; color:rgba(255,255,255,0.2); text-align:center; letter-spacing:.5px;">
      Compiled MOP Generator v6<br>
      Ericsson Internal Tool
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# MAIN AREA
# ─────────────────────────────────────────────────────────────────

st.markdown("""
<div class="eri-topbar">
  <div>
    <div class="eri-logo-text">ERICSSON</div>
    <div class="eri-logo-sub">Telecom Automation Toolkit</div>
  </div>
  <div style="text-align:center;">
    <div class="eri-app-title">Compiled MOP Generator</div>
    <div class="eri-app-sub">Solution Document + Activity MOP → Formatted Output MOP · Images Injected · Audit-Ready</div>
  </div>
  <div>
    <span class="eri-version">v6</span>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Zero Data Retention Banner ──────────────────────────────────
st.markdown("""
<div class="priv-bar">
  🔒 <strong>ZERO DATA RETENTION:</strong> All processing is performed entirely in-memory.
  No uploaded files, generated documents, or any user data are written to disk, logged, or stored
  at any stage of processing. This session and all associated data are permanently cleared when
  you close your browser tab.
</div>
""", unsafe_allow_html=True)

# ── Layout ──────────────────────────────────────────────────────
col_left, col_right = st.columns([1.1, 1], gap="large")

with col_left:

    # ── Step 1: Template ─────────────────────────────────────────
    st.markdown('<div class="eri-card"><div class="eri-card-title"><span class="step-badge">STEP 01</span> Select MOP Template</div>', unsafe_allow_html=True)

    templates      = discover_templates()
    selected_template = None
    template_bytes    = None

    if not templates:
        st.markdown('<div class="pill-warn">⚠ No template found. Place <strong>.docx</strong> file in <code>templates/</code> folder, then restart.</div>', unsafe_allow_html=True)
    else:
        names = [t.name for t in templates]
        sel   = st.selectbox("Template file", names, label_visibility="visible")
        selected_template = next(t for t in templates if t.name == sel)
        template_bytes    = load_template_bytes(selected_template)
        st.markdown(f'<div class="pill-ok">✔ &nbsp;<strong>{sel}</strong> &nbsp;ready</div>', unsafe_allow_html=True)

    st.markdown("""
    <div style="font-size:0.72rem; color:#5a7a9a; background:rgba(0,82,130,0.05);
         border:1px solid rgba(0,191,165,0.15); border-left:3px solid #00BFA5;
         border-radius:0 6px 6px 0; padding:0.5rem 0.8rem; margin-top:0.5rem;">
      📁 &nbsp;<strong>To add/update a template</strong>, place the <code>.docx</code>
      file in the <code>templates/</code> folder and restart the app.
    </div>
    """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Step 2: Solution Document ────────────────────────────────
    st.markdown('<div class="eri-card"><div class="eri-card-title"><span class="step-badge">STEP 02</span> Upload Solution Document <span style="font-size:0.62rem;color:#cc4400;margin-left:6px;">REQUIRED</span></div>', unsafe_allow_html=True)
    st.markdown("""
    <div style="font-size:0.73rem;color:#5a7a9a;margin-bottom:0.5rem;">
      The AI-generated Solution Document (all 12 sections). All text content comes from this file.
      Images/attachments in this doc are not needed — they are sourced from the Activity MOP below.
    </div>
    """, unsafe_allow_html=True)

    sol_file = st.file_uploader("Solution Document (.docx)", type=["docx"],
                                key="sol_up", label_visibility="visible")
    if sol_file:
        size_kb = sol_file.size / 1024
        st.markdown(
            f'<div class="pill-ok">✔ &nbsp;<strong>{sol_file.name}</strong>'
            f' &nbsp;·&nbsp; {size_kb:.1f} KB</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown('<div class="pill-warn">⏳ Waiting for solution document…</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Step 3: Activity MOP (OPTIONAL) ─────────────────────────
    st.markdown('<div class="eri-card"><div class="eri-card-title"><span class="step-badge">STEP 03</span> Upload Activity MOP &nbsp;<span class="optional-badge">OPTIONAL</span></div>', unsafe_allow_html=True)
    st.markdown("""
    <div style="font-size:0.73rem;color:#5a7a9a;margin-bottom:0.5rem;">
      The original Activity MOP document (may contain screenshots, flow diagrams, and embedded
      attachments). <strong>Only media</strong> (images, attachments) is extracted from this file —
      no text data. Media is injected into the SOP section of the output MOP in positional order,
      replacing <code>[IMAGE/SCREENSHOT REQUIRED]</code> placeholders.
    </div>
    """, unsafe_allow_html=True)

    mop_file = st.file_uploader("Activity MOP (.docx) — optional", type=["docx"],
                                key="mop_up", label_visibility="visible")
    if mop_file:
        size_kb = mop_file.size / 1024
        st.markdown(
            f'<div class="pill-ok">✔ &nbsp;<strong>{mop_file.name}</strong>'
            f' &nbsp;·&nbsp; {size_kb:.1f} KB — media will be extracted</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown('<div class="pill-info">ℹ No Activity MOP uploaded — output MOP will contain text only with placeholder notices.</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Step 4: Generate ─────────────────────────────────────────
    st.markdown('<div class="eri-card"><div class="eri-card-title"><span class="step-badge">STEP 04</span> Generate Output MOP</div>', unsafe_allow_html=True)

    can_go  = bool(sol_file and templates)
    gen_btn = st.button("⚡  Generate MOP", disabled=not can_go)

    if not can_go:
        missing = []
        if not templates:
            missing.append("MOP template")
        if not sol_file:
            missing.append("solution document")
        if missing:
            st.markdown(f'<div class="pill-warn">⏳ Still needed: {" + ".join(missing)}</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)


with col_right:

    if gen_btn and can_go:
        st.markdown('<div class="eri-card"><div class="eri-card-title">⚙ Processing</div>', unsafe_allow_html=True)

        steps = [
            "Loading template",
            "Reading solution document",
            "Extracting activity name",
            "Parsing all 12 sections",
            "Reading Activity MOP media" if mop_file else "No Activity MOP — skipping media",
            "Extracting images from Activity MOP" if mop_file else "Proceeding without media",
            "Extracting attachments from Activity MOP" if mop_file else "Text-only mode active",
            "Clearing template boilerplate",
            "Injecting section content + media",
            "Updating revision table & header",
            "Finalising document",
        ]

        st.markdown('<div class="prog-wrap">', unsafe_allow_html=True)
        phs = [st.empty() for _ in steps]
        for ph, s in zip(phs, steps):
            ph.markdown(f'<div class="ps wait"><div class="pd wait"></div>{s}</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        try:
            activity_name = ""
            sections      = {}
            today_str     = ""
            output_bytes  = b""
            media_items   = []

            for i, (ph, step) in enumerate(zip(phs, steps)):
                ph.markdown(f'<div class="ps doing"><div class="pd doing"></div>{step}…</div>', unsafe_allow_html=True)
                time.sleep(0.10)

                if i == 0:
                    tmpl_b = load_template_bytes(selected_template)

                elif i == 1:
                    sol_bytes = sol_file.read()
                    sol_doc   = Document(io.BytesIO(sol_bytes))

                elif i == 2:
                    activity_name = extract_activity_name(sol_doc)
                    today_str     = datetime.today().strftime("%d-%m-%Y")

                elif i == 3:
                    sections = extract_sections(sol_doc)

                elif i == 4:
                    if mop_file:
                        mop_bytes = mop_file.read()

                elif i == 5:
                    if mop_file:
                        media_items = extract_media_from_activity_mop(mop_bytes)

                elif i == 6:
                    pass  # already done in step 5

                elif i == 10:
                    output_bytes, failed_media, injected_count = build_mop(
                        tmpl_b, activity_name, sections, today_str, media_items
                    )

                ph.markdown(f'<div class="ps done"><div class="pd done"></div>{step} ✓</div>', unsafe_allow_html=True)
                time.sleep(0.04)

            # ── Store in session_state ──────────────────────────
            st.session_state["output_bytes"]   = output_bytes
            st.session_state["activity_name"]  = activity_name
            st.session_state["today_str"]      = today_str
            st.session_state["sections"]       = sections
            st.session_state["filled"]         = sum(1 for k in SECTION_KEYS[:-1] if sections.get(k))
            st.session_state["images_n"]       = len([m for m in media_items if m.kind == "image"])
            st.session_state["total_n"]        = sum(len(v) for k, v in sections.items()
                                                      if k != "connectivity_diagram")
            st.session_state["failed_media"]   = failed_media
            st.session_state["injected_media"] = injected_count

            st.markdown('</div>', unsafe_allow_html=True)

        except Exception as e:
            st.markdown('</div>', unsafe_allow_html=True)
            st.error(f"❌ Error during generation: {e}")
            import traceback
            st.code(traceback.format_exc())

    # ── Result panel ─────────────────────────────────────────────
    if st.session_state.get("output_bytes"):
        activity_name  = st.session_state["activity_name"]
        today_str      = st.session_state["today_str"]
        sections       = st.session_state["sections"]
        output_bytes   = st.session_state["output_bytes"]
        filled         = st.session_state["filled"]
        images_n       = st.session_state["images_n"]
        total_n        = st.session_state["total_n"]
        failed_media   = st.session_state["failed_media"]
        injected_media = st.session_state["injected_media"]

        # ── Success card ─────────────────────────────────────────
        st.markdown(f"""
        <div class="success-card">
          <div class="success-icon">✅</div>
          <div class="success-title">Output MOP Generated Successfully</div>
          <div class="success-sub">
            <strong class="success-name">{activity_name}</strong>
            &nbsp;·&nbsp; {today_str}
          </div>
        </div>""", unsafe_allow_html=True)

        safe_name = re.sub(r'[^\w\s\-]', '', activity_name).strip().replace(' ', '_')[:80]
        _dl_key   = f"dl_{abs(hash(safe_name + today_str)) % 10_000_000}"
        st.download_button(
            label="📥  Download Output MOP (.docx)",
            data=io.BytesIO(output_bytes),
            file_name=f"{safe_name}_MOP.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=_dl_key,
            use_container_width=True,
        )

        # ── Failed media warning ──────────────────────────────────
        if failed_media:
            st.markdown("""
            <div class="warn-bar">
              <strong>⚠ Some media items could not be automatically inserted into the output MOP.</strong>
              They are marked with a notice in the document at the exact position they should appear.
              Please insert them manually using Word's Insert → Pictures or Insert → Object.
            </div>
            """, unsafe_allow_html=True)

            st.markdown('<div class="media-fail-card">', unsafe_allow_html=True)
            st.markdown('<div class="media-fail-title">🖼 Media requiring manual insertion:</div>', unsafe_allow_html=True)
            for desc in failed_media:
                st.markdown(f'<div class="media-fail-item">• {desc}</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        elif images_n > 0:
            st.markdown(f'<div class="pill-ok" style="margin-top:0.6rem;">✔ &nbsp;All {injected_media} media items injected successfully into the SOP section</div>', unsafe_allow_html=True)

        # ── Metrics ───────────────────────────────────────────────
        st.markdown('<div class="eri-card" style="margin-top:0.8rem;"><div class="eri-card-title">📊 Generation Summary</div>', unsafe_allow_html=True)
        st.markdown(f"""
        <div class="metric-row">
          <div class="metric-box">
            <div class="metric-val">{filled}<span style="font-size:.85rem;color:#9ab8b4;">/12</span></div>
            <div class="metric-sub">Sections Filled</div>
          </div>
          <div class="metric-box">
            <div class="metric-val">{injected_media}</div>
            <div class="metric-sub">Media Injected</div>
          </div>
          <div class="metric-box">
            <div class="metric-val">{len(failed_media)}</div>
            <div class="metric-sub">Manual Inserts Needed</div>
          </div>
          <div class="metric-box">
            <div class="metric-val">{total_n}</div>
            <div class="metric-sub">Content Lines</div>
          </div>
        </div>""", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # ── Section fill status ───────────────────────────────────
        with st.expander("📋 Section fill status", expanded=False):
            for k in SECTION_KEYS[:-1]:
                label  = SECTION_LABELS.get(k, k)
                filled_flag = bool(sections.get(k))
                icon   = "✅" if filled_flag else "⚠️"
                color  = "#006633" if filled_flag else "#cc5500"
                st.markdown(
                    f'<div style="font-size:0.77rem;color:{color};padding:3px 0;">'
                    f'{icon} &nbsp; {label}</div>',
                    unsafe_allow_html=True,
                )

    # ── Footer ───────────────────────────────────────────────────
    st.markdown("""
    <div class="footer">
      Compiled MOP Generator v6 &nbsp;·&nbsp; Ericsson Internal Tool &nbsp;·&nbsp;
      🔒 Zero Data Retention &nbsp;·&nbsp; All processing in-memory only
    </div>
    """, unsafe_allow_html=True)
