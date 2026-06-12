"""
database.py — SQLite layer for SAESL subtask data.
content_items (ordered text+image sequence) is stored as JSON.
"""

import json
import sqlite3
import uuid as _uuid
import re as _re
from pathlib import Path
from typing import Optional


class Database:
    def __init__(self, db_path: str = "output/saesl.db"):
        Path(db_path).parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(db_path)
        self.conn.row_factory = sqlite3.Row
        self._create_tables()

    def _create_tables(self):
        self.conn.executescript("""
            CREATE TABLE IF NOT EXISTS subtasks (
                id              TEXT PRIMARY KEY,
                source_pdf      TEXT,
                task_id         TEXT,
                subtask_id      TEXT UNIQUE,
                title           TEXT,
                revision_date   TEXT,
                content_items   TEXT,
                procedure_steps TEXT,
                notes           TEXT,
                cautions        TEXT,
                warnings        TEXT,
                figure_refs     TEXT,
                data_cards      TEXT,
                rtv_refs        TEXT,
                cross_refs      TEXT,
                accountability  TEXT,
                raw_text        TEXT,
                created_at      DATETIME DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS figures (
                id          TEXT PRIMARY KEY,
                subtask_id  TEXT,
                label       TEXT,
                image_path  TEXT,
                FOREIGN KEY (subtask_id) REFERENCES subtasks(subtask_id)
            );

            CREATE TABLE IF NOT EXISTS figure_index (
                figure_id   TEXT PRIMARY KEY,
                image_path  TEXT NOT NULL,
                caption     TEXT,
                page_num    INTEGER,
                source_pdf  TEXT
            );

            CREATE INDEX IF NOT EXISTS idx_subtask_id  ON subtasks(subtask_id);
            CREATE INDEX IF NOT EXISTS idx_task_id     ON subtasks(task_id);
            CREATE INDEX IF NOT EXISTS idx_fig_subtask ON figures(subtask_id);
            CREATE INDEX IF NOT EXISTS idx_fig_index   ON figure_index(figure_id);
        """)
        self.conn.commit()

    # ──────────────────────────────────────────
    # Write
    # ──────────────────────────────────────────

    def save_subtask(self, rec) -> None:
        self.conn.execute("""
            INSERT OR REPLACE INTO subtasks
                (id, source_pdf, task_id, subtask_id, title, revision_date,
                 content_items, procedure_steps, notes, cautions, warnings,
                 figure_refs, data_cards, rtv_refs, cross_refs, accountability, raw_text)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            rec.id, rec.source_pdf, rec.task_id, rec.subtask_id,
            rec.title, rec.revision_date,
            json.dumps(rec.content_items),
            json.dumps(rec.procedure_steps),
            json.dumps(rec.notes),
            json.dumps(rec.cautions),
            json.dumps(rec.warnings),
            json.dumps(rec.figure_refs),
            json.dumps(rec.data_cards),
            json.dumps(rec.rtv_refs),
            json.dumps(rec.cross_refs),
            json.dumps(rec.accountability),
            rec.raw_text,
        ))

        self.conn.execute("DELETE FROM figures WHERE subtask_id = ?", (rec.subtask_id,))
        for fig in rec.figure_images:
            self.conn.execute(
                "INSERT INTO figures (id, subtask_id, label, image_path) VALUES (?,?,?,?)",
                (str(_uuid.uuid4()), rec.subtask_id, fig["label"], fig["path"])
            )

        self.conn.commit()

    def save_figure_index(self, entries: list[dict]) -> None:
        """
        entries: list of {figure_id, image_path, caption, page_num, source_pdf}
        Upserts into figure_index table.
        """
        for e in entries:
            self.conn.execute("""
                INSERT OR REPLACE INTO figure_index
                    (figure_id, image_path, caption, page_num, source_pdf)
                VALUES (?, ?, ?, ?, ?)
            """, (
                e["figure_id"],
                e["image_path"],
                e.get("caption", ""),
                e.get("page_num", 0),
                e.get("source_pdf", ""),
            ))
        self.conn.commit()

    def get_figure_image(self, figure_id: str) -> Optional[dict]:
        """Look up a figure by its ID. Returns {figure_id, image_path, caption} or None."""
        # Try exact match first
        row = self.conn.execute(
            "SELECT * FROM figure_index WHERE figure_id = ?", (figure_id,)
        ).fetchone()
        if row:
            return dict(row)
        # Fallback: normalise dashes/underscores and try LIKE
        norm = figure_id.replace("-", "%").replace("_", "%")
        row = self.conn.execute(
            "SELECT * FROM figure_index WHERE figure_id LIKE ?", (norm,)
        ).fetchone()
        return dict(row) if row else None

    def list_figure_index(self) -> list[dict]:
        rows = self.conn.execute(
            "SELECT figure_id, image_path, caption, page_num FROM figure_index ORDER BY figure_id"
        ).fetchall()
        return [dict(r) for r in rows]

    # ──────────────────────────────────────────
    # Read
    # ──────────────────────────────────────────

    def get_subtask(self, subtask_id: str) -> Optional[dict]:
        row = self.conn.execute(
            "SELECT * FROM subtasks WHERE subtask_id = ?", (subtask_id,)
        ).fetchone()
        return self._deserialize(dict(row)) if row else None

    def get_subtask_with_refs(self, subtask_id: str) -> Optional[dict]:
        main = self.get_subtask(subtask_id)
        if not main:
            return None
        # Inject live image-existence check into content_items
        main["content_items"] = self._validate_image_items(main.get("content_items", []))
        # Figures list (for stats)
        main["figures"] = [
            ci for ci in main["content_items"] if ci.get("type") == "image"
        ]
        # Resolved cross-refs
        resolved = []
        for ref_id in main.get("cross_refs", []):
            ref_data = self.get_subtask(ref_id)
            if ref_data:
                ref_data["content_items"] = self._validate_image_items(
                    ref_data.get("content_items", [])
                )
                resolved.append(ref_data)
        main["resolved_refs"] = resolved
        return main

    def _validate_image_items(self, items: list) -> list:
        """Remove image items whose file no longer exists on disk."""
        validated = []
        for ci in items:
            if ci.get("type") == "image":
                if ci.get("path") and Path(ci["path"]).exists():
                    validated.append(ci)
                # else: skip — file missing
            else:
                validated.append(ci)
        return validated

    def search_subtasks(self, query: str) -> list[dict]:
        rows = self.conn.execute("""
            SELECT subtask_id, task_id, title, revision_date
            FROM subtasks WHERE subtask_id LIKE ? OR title LIKE ?
            ORDER BY subtask_id LIMIT 50
        """, (f"%{query}%", f"%{query}%")).fetchall()
        return [dict(r) for r in rows]

    def list_all(self) -> list[dict]:
        rows = self.conn.execute(
            "SELECT subtask_id, task_id, title, revision_date, source_pdf FROM subtasks ORDER BY subtask_id"
        ).fetchall()
        return [dict(r) for r in rows]

    def stats(self) -> dict:
        n_sub = self.conn.execute("SELECT COUNT(*) FROM subtasks").fetchone()[0]
        n_fig = self.conn.execute("SELECT COUNT(*) FROM figures").fetchone()[0]
        return {"total_subtasks": n_sub, "total_figures": n_fig}

    # ──────────────────────────────────────────
    # Helpers
    # ──────────────────────────────────────────

    def _deserialize(self, row: dict) -> dict:
        json_fields = [
            "content_items", "procedure_steps", "notes", "cautions", "warnings",
            "figure_refs", "data_cards", "rtv_refs", "cross_refs", "accountability",
        ]
        for f in json_fields:
            if row.get(f):
                try:
                    row[f] = json.loads(row[f])
                except (json.JSONDecodeError, TypeError):
                    row[f] = []
        return row

    def close(self):
        self.conn.close()
