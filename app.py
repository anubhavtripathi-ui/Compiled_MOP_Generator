import streamlit as st
import io
import os
import re
import zipfile
import copy
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

st.set_page_config(page_title="MOP Generator", page_icon="📄", layout="centered")

st.markdown("""
<style>
    .stButton>button {
        background-color: #0066cc; color: white;
        border-radius: 8px; padding: 0.5em 2em;
        font-size: 16px; font-weight: bold; width: 100%;
    }
    .success-box {
        background-color: #d4edda; border-left: 5px solid #28a745;
        padding: 1em; border-radius: 5px; margin-top: 1em;
    }
</style>
""", unsafe_allow_html=True)

TEMPLATE_HEADINGS = [
    "Objective", "Activity Description", "Activity Type", "Domain in Scope",
    "Pre-requisites", "Inventory Details", "Node Connectivity Process",
    "Identity and Access Management", "Activity Triggering Method",
    "Standard Operating Procedure", "Acceptance Criteria", "Assumptions"
]

HEADING_ALIASES = {
    "Objective": ["objective", "scope", "purpose", "goal"],
    "Activity Description": ["activity description", "task overview", "process description",
                              "work summary", "activity overview", "execution details"],
    "Activity Type": ["activity type", "task category", "process type",
                      "operation type", "activity classification"],
    "Domain in Scope": ["domain in scope", "applicable domain", "functional scope", "domain"],
    "Pre-requisites": ["pre-requisites", "prerequisites", "pre-conditions",
                       "initial requirements", "mandatory conditions", "preconditions"],
    "Inventory Details": ["inventory details", "infrastructure details",
                          "system inventory", "inventory"],
    "Node Connectivity Process": ["node connectivity process", "connectivity workflow",
                                   "integration process", "network configuration steps",
                                   "connection procedure", "node connectivity"],
    "Identity and Access Management": ["identity and access management", "access control details",
                                        "authentication process", "authorization matrix",
                                        "iam", "identity & access management"],
    "Activity Triggering Method": ["activity triggering method", "trigger mechanism",
                                    "initiation method", "activation process",
                                    "execution trigger", "event trigger", "triggering method"],
    "Standard Operating Procedure": ["standard operating procedure", "operational guidelines",
                                      "process manual", "execution procedure", "work instructions",
                                      "step-by-step guide", "sop",
                                      "standard operating procedure (attach the detailed sop)"],
    "Acceptance Criteria": ["acceptance criteria", "validation criteria", "test scenarios",
                             "approval conditions", "success parameters", "uat checklist",
                             "uat criteria", "uat scenarios",
                             "acceptance criteria (uat scenarios)"],
    "Assumptions": ["assumptions", "presumptions", "considerations", "operating assumptions"]
}

DEFAULT_CONTENT = {
    "Objective": "The objective of this activity is to perform {activity} as part of the operational process for {vendor}. This procedure ensures that all necessary steps are followed systematically and accurately. The goal is to achieve the desired outcome with minimal risk and downtime.",
    "Activity Description": "This activity involves the execution of {activity} by {vendor} team. The process covers all relevant steps from initiation to completion. All actions will be performed in accordance with standard operational guidelines.",
    "Activity Type": "This is a planned operational activity of type: {activity}. It is categorized under standard change management procedures for {vendor}. The activity classification is based on its impact and execution scope.",
    "Domain in Scope": "The domain in scope for this activity includes the systems and components managed by {vendor}. All functional areas relevant to {activity} are included within this scope. Out-of-scope items will be documented separately if applicable.",
    "Pre-requisites": "Prior to executing {activity}, all prerequisite conditions must be verified. Access credentials, system availability, and approvals from {vendor} must be confirmed. All stakeholders should be notified and change window should be secured.",
    "Inventory Details": "The inventory for {activity} includes relevant nodes, systems, and components managed by {vendor}. A detailed inventory list should be prepared prior to execution. Node names, types, counts, and vendor details must be documented and verified.",
    "Node Connectivity Process": "The node connectivity process for {activity} involves verifying all network paths and connections. {vendor} team will ensure proper integration and connectivity between all nodes. Connectivity tests will be performed before and after the activity.",
    "Identity and Access Management": "Access management for {activity} will follow the standard IAM process defined by {vendor}. All user accounts and roles must be verified prior to execution. Access logs will be maintained for audit and compliance purposes.",
    "Activity Triggering Method": "The activity {activity} will be triggered based on the approved change request from {vendor}. Initiation will follow the standard trigger mechanism defined in the change management process. The execution will commence only after receiving explicit approval.",
    "Standard Operating Procedure": "",
    "Acceptance Criteria": "The acceptance criteria for {activity} will be validated by {vendor} team post-execution. All UAT scenarios must pass before the activity is marked as complete. Any deviations from expected results must be documented and escalated immediately.",
    "Assumptions": "It is assumed that all required systems are available and accessible during {activity}. {vendor} team will have necessary access and permissions throughout the execution window. Any changes in assumptions will be communicated to all stakeholders prior to execution."
}

def get_para_text(para_element):
    return "".join(t.text or "" for t in para_element.iter(qn("w:t"))).strip()

def is_heading_formatted(para_element, doc):
    """
    Check if paragraph looks like a heading by:
    1. Word Heading style (Heading 1/2/3 etc.)
    2. Bold text
    3. Underlined text
    4. Ends with : or - or :-
    5. Font size larger than 11pt (body text)
    """
    from docx.text.paragraph import Paragraph
    try:
        para = Paragraph(para_element, doc)
        style_name = para.style.name.lower()

        # Rule 1: Word heading style
        if "heading" in style_name or style_name in ["title", "subtitle"]:
            return True

        text = para.text.strip()

        # Rule 2: Ends with : or - or :-
        if text.endswith(":") or text.endswith("-") or text.endswith(":-"):
            return True

        # Check runs for formatting
        runs = para.runs
        if not runs:
            return False

        for run in runs:
            # Rule 3: Bold
            if run.bold:
                return True
            # Rule 4: Underline
            if run.underline:
                return True
            # Rule 5: Font size > 11pt
            if run.font.size and run.font.size.pt and run.font.size.pt > 11:
                return True

        # Also check paragraph-level bold via pPr/rPr
        pPr = para_element.find(qn("w:pPr"))
        if pPr is not None:
            rPr = pPr.find(qn("w:rPr"))
            if rPr is not None:
                if rPr.find(qn("w:b")) is not None:
                    return True
                if rPr.find(qn("w:u")) is not None:
                    return True

    except Exception:
        pass
    return False

def match_alias(text, seen_headings):
    """Match cleaned text against aliases. Returns canonical name or None."""
    clean = re.sub(r'[^a-z0-9 &()]', '', text.strip().lower()).strip()
    if not clean or len(clean) < 3:
        return None
    for canonical, aliases in HEADING_ALIASES.items():
        for alias in aliases:
            if alias == clean or alias in clean or clean in alias:
                if canonical not in seen_headings:
                    return canonical
                else:
                    return "__DUPLICATE__"
    return None

def extract_sections_xml(file_bytes):
    """
    Extract heading sections from docx using raw XML.
    A heading is detected only if BOTH conditions are true:
      1. Paragraph has heading-like formatting (bold/underline/heading style/ends with :/-)
      2. Text matches one of our known heading aliases
    Content between headings copied as raw XML blocks (preserves OLE, tables, images).
    """
    doc = Document(io.BytesIO(file_bytes))
    body = doc.element.body
    all_blocks = list(body)

    sections = {}
    seen = set()
    current_heading = None
    current_blocks = []

    for block in all_blocks:
        tag = block.tag.split('}')[-1]

        if tag == 'p':
            block_text = get_para_text(block)
            if block_text:
                # Check formatting first, then name match
                looks_like_heading = is_heading_formatted(block, doc)
                matched = match_alias(block_text, seen) if looks_like_heading else None

                if matched == "__DUPLICATE__":
                    if current_heading and current_blocks:
                        sections[current_heading] = current_blocks
                    current_heading, current_blocks = None, []
                    continue

                elif matched:
                    if current_heading and current_blocks:
                        sections[current_heading] = current_blocks
                    current_heading = matched
                    current_blocks = []
                    seen.add(matched)
                    continue

        # Collect everything (paragraphs, tables, OLE objects) as raw XML
        if current_heading is not None:
            current_blocks.append(copy.deepcopy(block))

    if current_heading and current_blocks:
        sections[current_heading] = current_blocks

    return sections, doc

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def inject_xml_blocks(doc, xml_blocks):
    """Inject raw XML blocks into doc body — preserves OLE/attachments/tables."""
    body = doc.element.body
    for block in xml_blocks:
        body.append(copy.deepcopy(block))

def generate_mop(activity_name, vendor_name, extracted_xml, input_doc):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(1.5)
        sec.bottom_margin = Cm(1.5)
        sec.left_margin = Cm(2)
        sec.right_margin = Cm(2)

    today = datetime.today().strftime("%d-%m-%Y")

    # Header Table
    ht = doc.add_table(rows=6, cols=4)
    ht.style = "Table Grid"
    rows_data = [
        ["Confidentiality Class", "External Confidentiality Label", "Document Type", "Sheet"],
        ["Ericsson Confidential", "", "Requirement Specification", "1"],
        ["Prepared By (Subject Responsible)", "", "Approved By (Document Responsible)", "Checked"],
        ["Automation SME", "", "", ""],
        ["Document Number", "Revision", "Date", "Reference"],
        ["", "", today, ""],
    ]
    label_rows = {0, 2, 4}
    for ri, row_data in enumerate(rows_data):
        for ci, text in enumerate(row_data):
            cell = ht.rows[ri].cells[ci]
            cell.text = text
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(8 if ri in label_rows else 9)
                    if ri in label_rows:
                        run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
                    if ri == 3:
                        run.bold = True

    doc.add_paragraph()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    tr = title.add_run("METHOD OF PROCEDURE")
    tr.font.size = Pt(18)
    tr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    doc.add_paragraph()

    rh = doc.add_paragraph()
    rhr = rh.add_run("Revision History (To be updated by Line/SME/AA)")
    rhr.bold = True
    rhr.font.size = Pt(12)

    rht = doc.add_table(rows=4, cols=4)
    rht.style = "Table Grid"
    col_headers = ["Version No.", "Revision Date", "Edited By", "Description of Change"]
    for i, cell in enumerate(rht.rows[0].cells):
        set_cell_bg(cell, "00BFFF")
        cell.text = col_headers[i]
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.bold = True
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(255, 255, 255)

    for ci, val in enumerate(["v1.0", today, "Automation SME", "Initial draft"]):
        cell = rht.rows[1].cells[ci]
        cell.text = val
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in cell.paragraphs[0].runs:
            run.font.size = Pt(10)

    for row in rht.rows[2:]:
        for cell in row.cells:
            cell.text = ""

    doc.add_paragraph()

    # 12 Headings
    for heading in TEMPLATE_HEADINGS:
        hp = doc.add_paragraph()
        hr = hp.add_run(heading)
        hr.bold = True
        hr.font.size = Pt(13)
        hr.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

        if heading == "Standard Operating Procedure":
            sp = doc.add_paragraph()
            sr = sp.add_run("Standard Operating Procedure (Attach the detailed SOP)")
            sr.bold = True
            sr.font.size = Pt(11)
            # Inject FULL input doc as-is
            for block in list(input_doc.element.body):
                tag = block.tag.split('}')[-1]
                if tag in ['p', 'tbl', 'sdt', 'bookmarkStart', 'bookmarkEnd']:
                    doc.element.body.append(copy.deepcopy(block))

        elif heading in extracted_xml and extracted_xml[heading]:
            inject_xml_blocks(doc, extracted_xml[heading])

        else:
            default = DEFAULT_CONTENT.get(heading, "")
            if default:
                filled = default.format(activity=activity_name, vendor=vendor_name)
                cp = doc.add_paragraph()
                cp.add_run(filled).font.size = Pt(11)

        doc.add_paragraph()

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

def create_zip(mop_bytes, sop_bytes, activity, vendor, sop_ext):
    mop_name = f"{activity}_{vendor}_MOP.docx"
    sop_name = f"SOP_{activity}{sop_ext}"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(mop_name, mop_bytes)
        zf.writestr(sop_name, sop_bytes)
    buf.seek(0)
    return buf.read()

# UI
st.title("📄 MOP Generator")
st.markdown("Generate a structured **Method of Procedure** document instantly.")
st.markdown("---")

col1, col2 = st.columns(2)
with col1:
    activity_name = st.text_input("🏷️ Activity Name",
                                   placeholder="e.g., Barring Unbarring Automation")
with col2:
    vendor_name = st.text_input("🏢 Vendor Name", placeholder="e.g., Ericsson")

uploaded_file = st.file_uploader(
    "📁 Upload Input MOP File (Max 30MB — .docx, .doc, .pdf, .txt)",
    type=["docx", "doc", "pdf", "txt"]
)

if uploaded_file:
    size_mb = len(uploaded_file.getvalue()) / (1024 * 1024)
    if size_mb > 30:
        st.error(f"❌ File too large ({size_mb:.1f}MB). Max: 30MB.")
        st.stop()
    else:
        st.success(f"✅ Uploaded: {uploaded_file.name} ({size_mb:.2f}MB)")

template = st.radio("🔘 Template", ["Template 1", "Template 2 (Coming Soon)"],
                    horizontal=True)
if template == "Template 2 (Coming Soon)":
    st.info("Template 2 is not yet available.")

st.markdown("---")

if st.button("⚡ Generate MOP"):
    if not activity_name.strip():
        st.error("❌ Please enter the Activity Name.")
    elif not vendor_name.strip():
        st.error("❌ Please enter the Vendor Name.")
    elif not uploaded_file:
        st.error("❌ Please upload an input MOP file.")
    elif template == "Template 2 (Coming Soon)":
        st.warning("⚠️ Template 2 not available yet.")
    else:
        with st.spinner("🔄 Generating your MOP... Please wait."):
            try:
                file_bytes = uploaded_file.getvalue()
                ext = os.path.splitext(uploaded_file.name)[1].lower()

                if ext in [".docx", ".doc"]:
                    extracted_xml, input_doc = extract_sections_xml(file_bytes)
                else:
                    extracted_xml = {}
                    input_doc = Document()

                act = activity_name.strip().replace(" ", "_")
                ven = vendor_name.strip().replace(" ", "_")

                mop_bytes = generate_mop(activity_name.strip(), vendor_name.strip(),
                                         extracted_xml, input_doc)
                zip_bytes = create_zip(mop_bytes, file_bytes, act, ven, ext or ".docx")

                st.markdown("""
                <div class='success-box'>
                    ✅ <strong>MOP Generated Successfully!</strong><br><br>
                    Your ZIP contains:<br>
                    &nbsp;&nbsp;📄 <strong>ActivityName_VendorName_MOP.docx</strong> — Formatted Output MOP<br>
                    &nbsp;&nbsp;📎 <strong>SOP_ActivityName.[ext]</strong> — Attach manually in SOP section
                </div>
                """, unsafe_allow_html=True)

                st.download_button(
                    label="📥 Download ZIP",
                    data=zip_bytes,
                    file_name=f"{act}_{ven}_MOP_Package.zip",
                    mime="application/zip"
                )

                found = [h for h in TEMPLATE_HEADINGS
                         if h in extracted_xml
                         and h != "Standard Operating Procedure"
                         and extracted_xml[h]]
                not_found = [h for h in TEMPLATE_HEADINGS
                             if h not in extracted_xml
                             and h != "Standard Operating Procedure"]

                if found:
                    st.info(f"✅ Headings found in input: {', '.join(found)}")
                if not_found:
                    st.warning(f"📝 Filled with defaults: {', '.join(not_found)}")

            except Exception as e:
                st.error(f"❌ Error: {str(e)}")

st.markdown("---")
st.markdown(
    "<small>🔒 No data stored. All processing in memory. Free to use.</small>",
    unsafe_allow_html=True
)
```

---

**Key upgrades in this final version:**

✅ **Smart heading detection** — BOTH conditions must be true:
- Formatting: bold / underline / heading style / ends with `:` or `-` or `:-` / font size > 11pt
- Name: matches alias list

✅ **Raw XML injection** — OLE objects, tables, images — sab as-is copy

✅ **SOP** — poora input doc as-is inject

✅ **Zero storage** — sab BytesIO mein

**requirements.txt same rahega:**
```
streamlit
python-docx
pdfplumber
lxml
