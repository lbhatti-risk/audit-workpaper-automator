# ITGC Audit Workpaper Automator

A CLI tool that uses Claude AI to automate the documentation of IT General Controls (ITGC) design and implementation reviews, generating a formatted Excel workpaper.

## What it does

For each ITGC control you review, the tool will:

1. **Analyse evidence** — uploads screenshots, Excel files, Word docs, and PDFs to Claude, which describes each piece of evidence in professional audit workpaper language
2. **Perform gap analysis** — Claude compares the client's stated process against the evidence for each tailored procedure, flagging gaps
3. **Identify control deficiencies** — if evidence doesn't support the process, Claude creates a CD/W with a description and severity rating
4. **Generate the workpaper** — produces a formatted `.xlsx` file with one tab per control, covering:
   - Section A: Tailored Procedures (with responses and gap flags)
   - Section B: Process Description
   - Section C: Evidence Analysis (chronological, with descriptions and alignment)
   - Section D: Control Deficiencies (CD/W references)
   - Section E: Conclusion

## Supported ITGC Controls

| # | Control |
|---|---------|
| 1 | Access Provisioning (New Starter) |
| 2 | Access Deprovisioning (Leaver) |
| 3 | Privileged Access Management |
| 4 | Authentication Controls |
| 5 | Database Access Controls |
| 6 | User Access Recertification (UAR) |
| 7 | Change Management — Code Changes |
| 8 | Change Management — Configuration Changes |

## Supported Evidence Formats

- **Images**: PNG, JPG, JPEG, GIF, BMP, WebP (analysed via Claude's vision)
- **Spreadsheets**: XLSX, XLS, CSV
- **Documents**: DOCX, DOC
- **PDF**: PDF
- **Plain text**: TXT and other text files

## Setup

```bash
# Install dependencies
pip install -r requirements.txt

# Set your API key
cp .env.example .env
# Edit .env and add your ANTHROPIC_API_KEY
```

## Usage

```bash
python main.py
```

The tool runs as an interactive CLI wizard:

1. Enter client name, application, and audit period
2. Select which ITGC controls to document
3. For each control, paste the client's end-to-end process description
4. Add evidence file paths one by one (with optional description hints)
5. The tool analyses everything with Claude and generates the workpaper

### Example session

```
====================================================
  ITGC Audit Workpaper Automator  |  Powered by Claude
====================================================

  Client name: Acme Corporation
  Application name (e.g. Oracle EBS, SAP): Oracle EBS
  Audit period (e.g. FY2025, Q1 2025): FY2025

  Available ITGC Controls:
     1.  Access Provisioning (New Starter)
     2.  Access Deprovisioning (Leaver)
    ...

  Select controls (comma-separated numbers): 1,2

  ────────────────────────────────────────────────────
  Control 1 of 2: Access Provisioning (New Starter)
  ────────────────────────────────────────────────────

  Paste the end-to-end process description:
  (Press Enter twice on a blank line to finish)

  > New starter requests are submitted via ServiceNow by the line manager...
  > Access is provisioned by IT within 2 business days of approval...
  >
  >

  Evidence files for 'Access Provisioning (New Starter)'
  File 1 path: /path/to/servicenow_request.png
  Brief description: ServiceNow access request ticket
  File 2 path: done

  Analysing: Access Provisioning (New Starter)
  Evidence analysis:
    ✓ servicenow_request.png
  Gap analysis ... complete  (1 deficiency identified)

====================================================
  ✓ Workpaper saved to: Acme_Corporation_Oracle_EBS_ITGC_Workpaper_FY2025.xlsx

  Controls documented:      2
  Evidence files analysed:  3
  Control deficiencies:     1
====================================================
```

### Options

```
python main.py --output my_workpaper.xlsx   # Custom output filename
python main.py --api-key sk-ant-...         # Pass API key directly
```
