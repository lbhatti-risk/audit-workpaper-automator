import base64
import os
from pathlib import Path
from typing import Optional

import anthropic

from .models import EvidenceItem, ITGCControl, ProcedureResult, ControlDeficiency
from .controls import TAILORED_PROCEDURES

EVIDENCE_TOOL = {
    "name": "record_evidence_analysis",
    "description": "Record the structured analysis of an ITGC audit evidence file.",
    "input_schema": {
        "type": "object",
        "properties": {
            "description": {
                "type": "string",
                "description": "1-3 sentence professional audit workpaper description of what the evidence shows (third-person, factual).",
            },
            "key_findings": {
                "type": "array",
                "items": {"type": "string"},
                "description": "List of key audit-relevant findings from this evidence.",
            },
            "aligns_with_process": {
                "type": "string",
                "enum": ["Yes", "No", "Partial"],
                "description": "Whether the evidence supports the stated control process.",
            },
            "alignment_explanation": {
                "type": "string",
                "description": "Brief explanation of the alignment assessment.",
            },
            "concerns": {
                "type": "string",
                "description": "Any gaps, anomalies, or missing elements observed. State 'None noted' if none.",
            },
        },
        "required": ["description", "key_findings", "aligns_with_process", "alignment_explanation", "concerns"],
    },
}

GAP_ANALYSIS_TOOL = {
    "name": "record_gap_analysis",
    "description": "Record the structured gap analysis results for an ITGC control review.",
    "input_schema": {
        "type": "object",
        "properties": {
            "procedure_results": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "number": {"type": "integer"},
                        "response": {
                            "type": "string",
                            "description": "Professional audit workpaper response for this procedure (1-3 sentences).",
                        },
                        "evidence_references": {
                            "type": "string",
                            "description": "Which evidence items support this procedure, e.g. 'Evidence 1, Evidence 3' or 'Not evidenced'.",
                        },
                        "has_gap": {
                            "type": "boolean",
                            "description": "True if the evidence does not sufficiently support this procedure.",
                        },
                    },
                    "required": ["number", "response", "evidence_references", "has_gap"],
                },
            },
            "deficiencies": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "procedure_number": {"type": "integer"},
                        "description": {
                            "type": "string",
                            "description": "Clear description of the deficiency: what was expected per the process vs. what the evidence showed.",
                        },
                        "severity": {
                            "type": "string",
                            "enum": ["Control Deficiency", "Significant Deficiency", "Material Weakness"],
                        },
                    },
                    "required": ["procedure_number", "description", "severity"],
                },
            },
            "conclusion": {
                "type": "string",
                "description": "2-3 sentence conclusion on the overall design and implementation effectiveness of the control.",
            },
        },
        "required": ["procedure_results", "deficiencies", "conclusion"],
    },
}


def _detect_file_type(file_path: str) -> str:
    ext = Path(file_path).suffix.lower()
    if ext in {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tiff", ".tif"}:
        return "image"
    if ext in {".xlsx", ".xls", ".csv"}:
        return "excel"
    if ext in {".docx", ".doc"}:
        return "word"
    if ext == ".pdf":
        return "pdf"
    return "text"


def _image_to_message_content(file_path: str) -> dict:
    ext = Path(file_path).suffix.lower()
    media_map = {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".gif": "image/gif",
        ".bmp": "image/bmp",
        ".webp": "image/webp",
    }
    media_type = media_map.get(ext, "image/png")
    with open(file_path, "rb") as f:
        data = base64.standard_b64encode(f.read()).decode("utf-8")
    return {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": data}}


def _read_as_text(file_path: str, file_type: str) -> str:
    if file_type == "excel":
        try:
            import pandas as pd
            df = pd.read_csv(file_path) if file_path.endswith(".csv") else pd.read_excel(file_path)
            return df.to_string(index=False, max_rows=100)
        except Exception as e:
            return f"[Could not parse Excel file: {e}]"

    if file_type == "word":
        try:
            from docx import Document
            doc = Document(file_path)
            return "\n".join(p.text for p in doc.paragraphs if p.text.strip())[:5000]
        except Exception as e:
            return f"[Could not parse Word file: {e}]"

    if file_type == "pdf":
        try:
            from pypdf import PdfReader
            reader = PdfReader(file_path)
            return "\n".join(page.extract_text() or "" for page in reader.pages)[:5000]
        except Exception as e:
            return f"[Could not parse PDF file: {e}]"

    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()[:5000]
    except Exception as e:
        return f"[Could not read file: {e}]"


class ITGCAnalyzer:
    def __init__(self, api_key: Optional[str] = None):
        self.client = anthropic.Anthropic(api_key=api_key or os.environ.get("ANTHROPIC_API_KEY"))
        self.model = "claude-sonnet-4-6"

    def analyze_evidence(self, control: ITGCControl, evidence: EvidenceItem) -> EvidenceItem:
        procedures_text = "\n".join(
            f"{i + 1}. {p}" for i, p in enumerate(TAILORED_PROCEDURES.get(control.control_type, []))
        )
        hint_line = f"\nAuditor's note about this evidence: {evidence.hint}" if evidence.hint else ""

        base_text = (
            f"You are analyzing evidence for an ITGC audit review.\n\n"
            f"Control: {control.name}\n"
            f"Application: {control.application}\n"
            f"Client: {control.client_name}\n"
            f"Evidence file: {evidence.filename}{hint_line}\n\n"
            f"Control Process Description:\n{control.process_description}\n\n"
            f"Tailored Procedures Being Tested:\n{procedures_text}\n\n"
            f"Analyze this evidence and record your findings using the provided tool."
        )

        if evidence.file_type == "image":
            content = [_image_to_message_content(evidence.file_path), {"type": "text", "text": base_text}]
        else:
            text_content = _read_as_text(evidence.file_path, evidence.file_type)
            content = [{"type": "text", "text": base_text + f"\n\nEvidence Content:\n{text_content}"}]

        response = self.client.messages.create(
            model=self.model,
            max_tokens=1024,
            system=(
                "You are an experienced ITGC auditor. "
                "Write professional, factual, third-person audit workpaper content."
            ),
            tools=[EVIDENCE_TOOL],
            tool_choice={"type": "any"},
            messages=[{"role": "user", "content": content}],
        )

        for block in response.content:
            if block.type == "tool_use":
                result = block.input
                evidence.description = result.get("description", "")
                evidence.key_findings = result.get("key_findings", [])
                evidence.aligns_with_process = result.get("aligns_with_process", "N/A")
                evidence.alignment_explanation = result.get("alignment_explanation", "")
                evidence.concerns = result.get("concerns", "None noted")
                break

        return evidence

    def perform_gap_analysis(self, control: ITGCControl) -> ITGCControl:
        procedures = TAILORED_PROCEDURES.get(control.control_type, [])
        procedures_text = "\n".join(f"{i + 1}. {p}" for i, p in enumerate(procedures))

        evidence_summary = "\n\n".join(
            f"Evidence {i + 1} ({ev.filename}):\n"
            f"  Description: {ev.description}\n"
            f"  Key Findings: {'; '.join(ev.key_findings)}\n"
            f"  Aligns with Process: {ev.aligns_with_process} — {ev.alignment_explanation}\n"
            f"  Concerns: {ev.concerns}"
            for i, ev in enumerate(control.evidence_items)
        ) or "No evidence provided."

        prompt = (
            f"Complete a gap analysis for this ITGC Design & Implementation review.\n\n"
            f"Client: {control.client_name}\n"
            f"Application: {control.application}\n"
            f"Control: {control.name}\n"
            f"Audit Period: {control.audit_period}\n\n"
            f"STATED CONTROL PROCESS:\n{control.process_description}\n\n"
            f"TAILORED PROCEDURES ({len(procedures)} total):\n{procedures_text}\n\n"
            f"EVIDENCE REVIEWED AND ANALYSED:\n{evidence_summary}\n\n"
            "For each tailored procedure, assess whether the evidence is sufficient and aligns with "
            "the stated process. Identify any gaps where evidence does not support a procedure and "
            "create control deficiencies accordingly. Record your findings using the provided tool."
        )

        response = self.client.messages.create(
            model=self.model,
            max_tokens=3000,
            system=(
                "You are a senior ITGC auditor completing a formal workpaper. "
                "Write professional, concise, factual audit content."
            ),
            tools=[GAP_ANALYSIS_TOOL],
            tool_choice={"type": "any"},
            messages=[{"role": "user", "content": prompt}],
        )

        for block in response.content:
            if block.type == "tool_use":
                result = block.input

                control.procedure_results = []
                for i, proc_text in enumerate(procedures):
                    n = i + 1
                    matched = next((r for r in result.get("procedure_results", []) if r.get("number") == n), None)
                    if matched:
                        control.procedure_results.append(
                            ProcedureResult(
                                number=n,
                                procedure_text=proc_text,
                                response=matched.get("response", ""),
                                evidence_references=matched.get("evidence_references", ""),
                                has_gap=matched.get("has_gap", False),
                            )
                        )
                    else:
                        control.procedure_results.append(
                            ProcedureResult(number=n, procedure_text=proc_text)
                        )

                control.deficiencies = []
                for cd_num, deficiency in enumerate(result.get("deficiencies", []), start=1):
                    control.deficiencies.append(
                        ControlDeficiency(
                            cd_number=cd_num,
                            procedure_number=deficiency.get("procedure_number", 0),
                            description=deficiency.get("description", ""),
                            severity=deficiency.get("severity", "Control Deficiency"),
                        )
                    )

                control.conclusion = result.get("conclusion", "")
                break

        return control
