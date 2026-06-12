"""
formatter.py
------------
Takes a subtask dict (from database) and formats it
exactly as Solumina expects — ready to be pasted.

Output format based on screenshots of actual Solumina entries:

    REFERENCE
    TASK          : CIR 72-41-31-100-802
    SUBTASK       : 72-41-31-110-066-001    Remove the Grease and Loose Dirt
    REVISION DATE : 15 JUL 2017

    ADDITIONAL DOCUMENT
    DATA CARD     : CLNDT-DC-OP101

    PROCEDURE
    NOTE: This SUBTASK is an alternative to SUBTASK 72-41-31-110-066-002.
    A. Clean the Stage 4 Disc by chemical cleaning.
    B. Refer to OP TASK 70-00-00-100-101 (Non-Aqueous Vapor and Liquid Degreasing).
    ...

    PART ACCOUNTABILITY
    Select and Record:
    • Part Functional Location
    • Part Number
    • Serial Number (if applicable)
"""


def format_subtask(subtask: dict) -> str:
    """
    Format a single subtask dict into Solumina-ready text.
    Cross-references are intentionally excluded from the output.

    Args:
        subtask: dict from database.get_subtask() or get_subtask_with_refs()

    Returns:
        Formatted string ready to copy-paste into Solumina
    """
    lines = []

    # ── REFERENCE block ────────────────────────────────────────
    lines.append("REFERENCE")

    task_id = subtask.get("task_id", "")
    if task_id:
        lines.append(f"TASK          : {task_id}")

    subtask_id = subtask.get("subtask_id", "")
    title      = subtask.get("title", "")
    if subtask_id:
        lines.append(f"SUBTASK       : {subtask_id:<35} {title}")

    rev_date = subtask.get("revision_date", "")
    if rev_date:
        lines.append(f"REVISION DATE : {rev_date}")

    lines.append("")  # blank line

    # ── ADDITIONAL DOCUMENT / DATA CARD ────────────────────────
    data_cards = subtask.get("data_cards", [])
    rtv_refs   = subtask.get("rtv_refs", [])

    if data_cards or rtv_refs:
        lines.append("ADDITIONAL DOCUMENT")
        for dc in data_cards:
            lines.append(f"DATA CARD     : {dc}")
        for rtv in rtv_refs:
            lines.append(f"RTV           : {rtv}")
        lines.append("")

    # ── PROCEDURE ──────────────────────────────────────────────
    lines.append("PROCEDURE")

    # Notes first (before steps, as seen in screenshots)
    for note in subtask.get("notes", []):
        lines.append(note)

    # Cautions
    for caution in subtask.get("cautions", []):
        lines.append(caution)

    # Warnings
    for warning in subtask.get("warnings", []):
        lines.append(warning)

    # Steps
    for step in subtask.get("procedure_steps", []):
        lines.append(step)

    lines.append("")

    # ── FIGURE references ──────────────────────────────────────
    figure_refs = subtask.get("figure_refs", [])
    if figure_refs:
        lines.append("FIGURES")
        for fig in figure_refs:
            lines.append(f"  Refer to Fig {fig}")
        lines.append("")

    # ── ACCOUNTABILITY sections ────────────────────────────────
    accountability = subtask.get("accountability", {})

    if "PART" in accountability:
        lines.append("PART ACCOUNTABILITY")
        lines.append("Select and Record:")
        lines.append("  \u2022 Part Functional Location")
        lines.append("  \u2022 Part Number")
        lines.append("  \u2022 Serial Number (if applicable)")
        lines.append("")

    if "OMAT" in accountability:
        lines.append("OMAT ACCOUNTABILITY")
        lines.append("Select and Record:")
        lines.append("  \u2022 OMat Name used")
        lines.append("  \u2022 OMat Batch number")
        lines.append("")

    if "TOOLING" in accountability:
        lines.append("TOOLING ACCOUNTABILITY")
        lines.append("Select and Record:")
        lines.append("  \u2022 Tooling used")
        lines.append("  \u2022 Tooling Serial Number")
        lines.append("")

    # NOTE: Cross-references are intentionally excluded from all output.

    return "\n".join(lines)


def format_for_clipboard(subtask: dict) -> str:
    """
    Same as format_subtask but trims leading/trailing whitespace
    and ensures it ends with a single newline — ideal for clipboard.
    """
    return format_subtask(subtask).strip() + "\n"


def preview_subtask(subtask: dict) -> str:
    """
    Short one-line summary of a subtask — used in the web app
    search results list.
    """
    sid   = subtask.get("subtask_id", "?")
    tid   = subtask.get("task_id", "?")
    title = subtask.get("title", "")[:50]
    steps = len(subtask.get("procedure_steps", []))
    figs  = len(subtask.get("figure_refs", []))
    return f"[{sid}] {tid} | {title} | {steps} steps | {figs} figs"
