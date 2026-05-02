#!/usr/bin/env python3
"""ITGC Audit Workpaper Automator — powered by Claude."""

import os
import sys
from pathlib import Path

import click
from dotenv import load_dotenv

load_dotenv()

from src.models import ControlType, CONTROL_NAMES, EvidenceItem, ITGCControl
from src.analyzer import ITGCAnalyzer, _detect_file_type
from src.workpaper import generate_workpaper

CONTROL_MENU = list(ControlType)


def _print_banner():
    click.echo()
    click.echo("=" * 60)
    click.echo("  ITGC Audit Workpaper Automator  |  Powered by Claude")
    click.echo("=" * 60)
    click.echo()


def _print_step(msg: str):
    click.echo(f"\n  → {msg}")


def _read_multiline(prompt: str) -> str:
    click.echo(f"\n{prompt}")
    click.echo("  (Press Enter twice on a blank line to finish)\n")
    lines = []
    while True:
        line = input("  > ")
        if line == "" and lines and lines[-1] == "":
            lines.pop()
            break
        lines.append(line)
    return "\n".join(lines).strip()


def _collect_evidence(control_name: str) -> list[EvidenceItem]:
    click.echo(f"\n  Evidence files for '{control_name}'")
    click.echo("  (Enter file paths one at a time. Type 'done' when finished.)\n")
    items = []
    order = 1
    while True:
        path_input = input(f"  File {order} path (or 'done'): ").strip()
        if path_input.lower() == "done":
            break
        if not os.path.isfile(path_input):
            click.echo(f"  ! File not found: {path_input}. Try again.")
            continue
        hint = input("  Brief description (optional, press Enter to skip): ").strip()
        file_type = _detect_file_type(path_input)
        items.append(
            EvidenceItem(
                filename=Path(path_input).name,
                file_path=path_input,
                file_type=file_type,
                order=order,
                hint=hint,
            )
        )
        click.echo(f"  ✓ Added ({file_type}): {Path(path_input).name}")
        order += 1
    return items


@click.command()
@click.option("--output", "-o", default=None, help="Output Excel file path.")
@click.option("--api-key", envvar="ANTHROPIC_API_KEY", default=None, help="Anthropic API key.")
def main(output: str, api_key: str):
    """Generate an ITGC audit workpaper with AI-assisted evidence analysis."""
    _print_banner()

    if not api_key:
        click.echo("  ERROR: ANTHROPIC_API_KEY is not set.")
        click.echo("  Set it in a .env file or pass --api-key.")
        sys.exit(1)

    # ── Engagement details ──────────────────────────────────────────────────
    client_name = click.prompt("  Client name").strip()
    application = click.prompt("  Application name (e.g. Oracle EBS, SAP)").strip()
    audit_period = click.prompt("  Audit period (e.g. FY2025, Q1 2025)").strip()

    # ── Select controls ─────────────────────────────────────────────────────
    click.echo("\n  Available ITGC Controls:")
    for i, ct in enumerate(CONTROL_MENU, start=1):
        click.echo(f"    {i:2}.  {CONTROL_NAMES[ct]}")

    raw = click.prompt("\n  Select controls (comma-separated numbers, e.g. 1,3,6)").strip()
    selected_indices = []
    for part in raw.split(","):
        try:
            idx = int(part.strip()) - 1
            if 0 <= idx < len(CONTROL_MENU):
                selected_indices.append(idx)
            else:
                click.echo(f"  ! Ignored out-of-range value: {part.strip()}")
        except ValueError:
            click.echo(f"  ! Ignored non-numeric value: {part.strip()}")

    if not selected_indices:
        click.echo("  No valid controls selected. Exiting.")
        sys.exit(1)

    # ── Collect process descriptions and evidence ───────────────────────────
    controls: list[ITGCControl] = []
    for step, idx in enumerate(selected_indices, start=1):
        ct = CONTROL_MENU[idx]
        control_name = CONTROL_NAMES[ct]
        click.echo(f"\n{'─' * 60}")
        click.echo(f"  Control {step} of {len(selected_indices)}: {control_name}")
        click.echo(f"{'─' * 60}")

        process = _read_multiline("  Paste the end-to-end process description:")
        evidence_items = _collect_evidence(control_name)

        controls.append(
            ITGCControl(
                control_type=ct,
                application=application,
                client_name=client_name,
                audit_period=audit_period,
                process_description=process,
                evidence_items=evidence_items,
            )
        )

    # ── Analyse with Claude ─────────────────────────────────────────────────
    analyzer = ITGCAnalyzer(api_key=api_key)
    total_cds = 0

    for ctrl in controls:
        click.echo(f"\n  Analysing: {ctrl.name}")

        if ctrl.evidence_items:
            click.echo("  Evidence analysis:")
            for ev in ctrl.evidence_items:
                try:
                    analyzer.analyze_evidence(ctrl, ev)
                    click.echo(f"    ✓ {ev.filename}")
                except Exception as e:
                    click.echo(f"    ! Failed to analyse {ev.filename}: {e}")
                    ev.description = f"[Analysis failed: {e}]"

        click.echo("  Gap analysis ... ", nl=False)
        try:
            analyzer.perform_gap_analysis(ctrl)
            cd_count = len(ctrl.deficiencies)
            total_cds += cd_count
            if cd_count:
                click.echo(f"complete  ({cd_count} deficien{'cy' if cd_count == 1 else 'cies'} identified)")
            else:
                click.echo("complete  (no deficiencies identified)")
        except Exception as e:
            click.echo(f"failed: {e}")

    # ── Generate workpaper ─────────────────────────────────────────────────
    if not output:
        safe_client = client_name.replace(" ", "_").replace("/", "-")
        safe_app = application.replace(" ", "_").replace("/", "-")
        output = f"{safe_client}_{safe_app}_ITGC_Workpaper_{audit_period}.xlsx"

    click.echo(f"\n  Generating Excel workpaper ...")
    generate_workpaper(controls, output)

    # ── Summary ─────────────────────────────────────────────────────────────
    click.echo()
    click.echo("=" * 60)
    click.echo(f"  ✓ Workpaper saved to: {output}")
    click.echo()
    click.echo(f"  Controls documented:      {len(controls)}")
    click.echo(f"  Evidence files analysed:  {sum(len(c.evidence_items) for c in controls)}")
    click.echo(f"  Control deficiencies:     {total_cds}")
    click.echo()
    click.echo("  ⚠  AUDITOR REVIEW REQUIRED")
    click.echo("  All AI-generated content is a draft. Before sign-off,")
    click.echo("  verify each evidence description, gap finding, and CD/W")
    click.echo("  against the source files and apply professional judgement.")
    click.echo("=" * 60)
    click.echo()


if __name__ == "__main__":
    main()
