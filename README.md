# Smart MOP Generator — v6
### Ericsson Internal Automation Tool

---

## Overview

Smart MOP Generator is a Streamlit-based web application that automates the generation of formatted **Method of Procedure (MOP)** documents. It takes an AI-generated Solution Document (produced via the Copilot/Claude prompt) and an optional Activity MOP, then merges them into a professionally formatted output MOP using a Word template.

**Key capability in v6:** Images, screenshots, flow diagrams, and embedded attachments from the Activity MOP are automatically extracted and injected into the correct positions in the SOP section (Section 10) of the output MOP, matching `[IMAGE/SCREENSHOT REQUIRED]` placeholders in positional order.

---

## 🔒 Zero Data Retention

> **All processing is performed entirely in-memory. No uploaded files, generated documents, or any user data are written to disk, logged, or stored at any stage. All data is permanently cleared when the browser session ends.**

This is enforced at both the application and infrastructure level and is prominently displayed in the UI.

---

## Folder Structure

```
your-project/
│
├── app.py                  ← Main application (this file)
├── requirements.txt        ← Python dependencies
├── README.md               ← This file
│
└── templates/
    └── MOP_Template.docx   ← Your Word MOP template (place here once, reuse always)
```

> The `templates/` folder must exist and contain at least one `.docx` template before launching the app.

---

## Installation

### 1. Prerequisites
- Python 3.10 or higher
- pip

### 2. Clone or download the project files
Place `app.py`, `requirements.txt`, and `README.md` in a folder.

### 3. Create the templates folder and add your template
```bash
mkdir templates
# Copy your MOP Word template into the templates/ folder
cp /path/to/your/MOP_Template_Updated.docx templates/
```

### 4. Install dependencies
```bash
pip install -r requirements.txt
```

### 5. Run the app
```bash
streamlit run app.py
```

The app will open automatically in your browser at `http://localhost:8501`.

---

## How to Use

### Step-by-step

| Step | Action | Required? |
|------|--------|-----------|
| 1 | Select MOP template from the dropdown (auto-detected from `templates/` folder) | ✅ Required |
| 2 | Upload **Solution Document** — the AI-generated document with all 12 sections (output from Copilot/Claude prompt) | ✅ Required |
| 3 | Upload **Activity MOP** — the original technical activity document containing screenshots, images, flow diagrams, or embedded attachments | ⬜ Optional |
| 4 | Click **Generate MOP** | ✅ |
| 5 | Download the output `.docx` file | ✅ |

---

## How Image/Attachment Injection Works

1. When an **Activity MOP** is uploaded, the app scans it for:
   - Embedded images and screenshots (`.png`, `.jpg`, `.gif`, `.bmp`, etc.)
   - OLE embedded files (`.xlsx`, `.docx`, `.bin`, etc.)
   - Images inside tables

2. The **Solution Document** (Section 10 — SOP) is scanned for placeholder text such as:
   - `[IMAGE/SCREENSHOT REQUIRED — ...]`
   - `[ATTACHMENT REQUIRED — ...]`
   - `[IMAGE: ...]`
   - `[SCREENSHOT: ...]`
   - `[DIAGRAM: ...]`

3. Images from the Activity MOP are injected into these placeholders **in positional order** — the 1st placeholder gets the 1st image, the 2nd placeholder gets the 2nd image, and so on.

4. If any image or attachment **cannot be inserted** (e.g. unsupported format):
   - A red warning notice is placed at the exact position in the output MOP: `⚠ [MEDIA NOT INSERTED — Please add manually: ...]`
   - The UI lists all failed items with a clear instruction to insert manually using Word's Insert → Pictures or Insert → Object.

---

## What Each Section Covers

| # | Section | Source |
|---|---------|--------|
| 1 | Objective | AI-generated (Solution Document) |
| 2 | Activity Description | AI-generated (Solution Document) |
| 3 | Activity Type | AI-generated (Solution Document) |
| 4 | Domain in Scope | AI-generated (Solution Document) |
| 5 | Pre-requisites | AI-generated (Solution Document) |
| 6 | Inventory Details | AI-generated (Solution Document) |
| 7 | Node Connectivity Process | AI-generated + image placeholders filled from Activity MOP |
| 8 | Identity and Access Management | AI-generated (Solution Document) |
| 9 | Activity Triggering Method | AI-generated (Solution Document) |
| 10 | **Standard Operating Procedure** | **Verbatim copy from Solution Document + images injected from Activity MOP** |
| 11 | Acceptance Criteria (UAT) | AI-generated (embedded UAT checklist + activity-specific additions) |
| 12 | Assumptions | AI-generated (Solution Document) |

---

## Generating the Solution Document (Copilot/Claude)

The Solution Document input for this app is generated using the **Smart MOP Prompt**. To generate it:

1. Open **Microsoft Copilot** (or Claude)
2. Paste the full Smart MOP Prompt
3. Upload your **Activity MOP** document
4. The AI will generate a complete 12-section Solution Document
5. Download/copy the output as a `.docx` file
6. Upload that `.docx` as the **Solution Document** in this app

> The prompt ensures that all images and attachments in the Activity MOP are referenced as `[IMAGE/SCREENSHOT REQUIRED — ...]` placeholders in Section 10. This app then replaces those placeholders with the actual media from the Activity MOP.

---

## Updating the Template

The MOP template is a **one-time setup**. To update it:

1. Replace the `.docx` file in the `templates/` folder
2. Restart the app (`Ctrl+C` then `streamlit run app.py`)
3. The new template will appear in the dropdown automatically

The template must use `Heading1` style for all 12 section headings (matching the section names in the app). The app clears boilerplate content under each heading and replaces it with the Solution Document content.

---

## Troubleshooting

| Issue | Cause | Fix |
|-------|-------|-----|
| "No template found" error | `templates/` folder missing or empty | Create folder and add `.docx` template |
| Section appears empty in output | Section heading in Solution Doc doesn't match expected names | Check that headings use standard names (e.g. "Standard Operating Procedure", "Pre-requisites") |
| Images not injected | No `[IMAGE/SCREENSHOT REQUIRED]` placeholders in Section 10 of Solution Doc | Re-generate Solution Document using the Smart MOP Prompt — it creates these placeholders automatically |
| Image listed as "manual insert needed" | Image format not supported by python-docx (e.g. EMF/WMF vector) | Insert manually in Word using Insert → Pictures |
| App crashes on large files | Memory limit | Reduce file size or increase server memory |

---

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| `streamlit` | ≥ 1.32.0 | Web UI framework |
| `python-docx` | ≥ 1.1.0 | Word document read/write |
| `lxml` | ≥ 5.1.0 | XML manipulation for docx internals |
| `Pillow` | ≥ 10.2.0 | Image format handling |

---

## Version History

| Version | Changes |
|---------|---------|
| v6 | Activity MOP image/attachment injection; positional placeholder matching; manual-insert notices; table image support; OLE attachment handling |
| v5 | Solution Document parser; 12-section template population; Ericsson UI theme |
| v4 | Multi-template support; section heading auto-detection |
| v3 | Zero data retention enforcement; in-memory processing |

---

## Contact / Support

For issues, feature requests, or template updates, raise a request with your local Ericsson automation team.

---

*Smart MOP Generator · Ericsson Internal Tool · Zero Data Retention · All processing in-memory*
