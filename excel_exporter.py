"""Hazırlanan ders çalışma programını orijinal Excel şablonlarına aktaran modül."""

from __future__ import annotations

import io
import math
import re
from copy import copy
from dataclasses import dataclass
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import Cell
from openpyxl.styles import Border

from curriculum import DAYS_OF_WEEK, MATRIX_SUBJECTS, tr_upper

TEMPLATES_DIR = Path(__file__).resolve().parent / "templates"


@dataclass
class ScheduleTask:
    ders: str
    konu: str
    study_type: str = "Soru Çözümü"
    target_questions: int | None = None
    note: str = ""

    def formatted_topic(self) -> str:
        parts = [self.konu.strip()]
        if self.study_type and self.study_type != "Standart":
            parts.append(f"[{self.study_type}]")
        if self.note.strip():
            parts.append(f"({self.note.strip()})")
        return " ".join(p for p in parts if p)


def _copy_cell_style(src: Cell, dst: Cell) -> None:
    if src.has_style:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format


def _sanitize_sheet_title(title: str, fallback: str) -> str:
    cleaned = re.sub(r"[:\\/*?\[\]]", "-", title).strip()
    return cleaned[:31] if cleaned else fallback


def _detect_day_blocks(ws) -> list[tuple[str, int, int]]:
    """Şablonun A sütunundan gün bloklarını (gün_adı, başlangıç_satırı, satır_kapasitesi) tespit eder."""
    day_set = set(DAYS_OF_WEEK)
    col_a_merges: dict[int, int] = {}
    for m in ws.merged_cells.ranges:
        if m.min_col == 1 and m.max_col == 1:
            col_a_merges[m.min_row] = m.max_row - m.min_row + 1

    blocks: list[tuple[str, int, int]] = []
    for r in range(1, min(ws.max_row + 1, 80)):
        val = ws.cell(r, 1).value
        if isinstance(val, str):
            norm = tr_upper(val.strip())
            if norm in day_set:
                cap = col_a_merges.get(r, 1)
                blocks.append((norm, r, cap))
    return blocks


def _expand_day_block(
    ws,
    insert_at: int,
    extra: int,
    max_col: int,
    extra_row_merges: list[tuple[int, int]] | None = None,
) -> None:
    """Belirli bir satırın üzerine `extra` adet stil-korumalı satır ekler ve birleşik hücreleri günceller."""
    if extra <= 0:
        return

    saved_merges: list[tuple[int, int, int, int]] = [
        (m.min_row, m.min_col, m.max_row, m.max_col)
        for m in list(ws.merged_cells.ranges)
    ]
    for min_r, min_c, max_r, max_c in saved_merges:
        ws.unmerge_cells(
            start_row=min_r, start_column=min_c, end_row=max_r, end_column=max_c
        )

    template_row = insert_at - 1
    template_height = ws.row_dimensions[template_row].height

    old_heights = {
        r: ws.row_dimensions[r].height
        for r in range(insert_at, ws.max_row + 1)
        if ws.row_dimensions[r].height is not None
    }

    ws.insert_rows(insert_at, amount=extra)

    for offset in range(extra):
        r = insert_at + offset
        if template_height is not None:
            ws.row_dimensions[r].height = template_height
        for col in range(1, max_col + 1):
            _copy_cell_style(ws.cell(template_row, col), ws.cell(r, col))

    for old_r, h in old_heights.items():
        ws.row_dimensions[old_r + extra].height = h

    updated_merges: list[tuple[int, int, int, int]] = []
    for min_r, min_c, max_r, max_c in saved_merges:
        if max_r < insert_at:
            updated_merges.append((min_r, min_c, max_r, max_c))
        elif min_r >= insert_at:
            updated_merges.append((min_r + extra, min_c, max_r + extra, max_c))
        else:
            updated_merges.append((min_r, min_c, max_r + extra, max_c))

    if extra_row_merges:
        for offset in range(extra):
            r = insert_at + offset
            for c1, c2 in extra_row_merges:
                updated_merges.append((r, c1, r, c2))

    for min_r, min_c, max_r, max_c in updated_merges:
        ws.merge_cells(
            start_row=min_r, start_column=min_c, end_row=max_r, end_column=max_c
        )


def _shrink_day_block(
    ws,
    start_row: int,
    base_cap: int,
    target_rows: int,
    max_col: int,
) -> int:
    """Gün blokundaki kullanılmayan alt satırları siler, alt kenarlığı ve birleşik hücreleri korur."""
    if target_rows >= base_cap or target_rows < 1:
        return 0

    remove_count = base_cap - target_rows
    base_end = start_row + base_cap - 1
    new_end = start_row + target_rows - 1
    delete_from = new_end + 1

    # Silinecek son satırın alt kenarlığını yeni son satıra aktar
    for col in range(1, max_col + 1):
        src_border = ws.cell(base_end, col).border
        dst_cell = ws.cell(new_end, col)
        dst_border = copy(dst_cell.border) if dst_cell.has_style else Border()
        dst_border.bottom = (
            copy(src_border.bottom) if src_border and src_border.bottom else None
        )
        dst_cell.border = dst_border

    saved_merges: list[tuple[int, int, int, int]] = [
        (m.min_row, m.min_col, m.max_row, m.max_col)
        for m in list(ws.merged_cells.ranges)
    ]
    for min_r, min_c, max_r, max_c in saved_merges:
        ws.unmerge_cells(
            start_row=min_r, start_column=min_c, end_row=max_r, end_column=max_c
        )

    old_heights = {
        r: ws.row_dimensions[r].height
        for r in range(base_end + 1, ws.max_row + 1)
        if ws.row_dimensions[r].height is not None
    }
    for r in range(delete_from, max(old_heights.keys(), default=base_end) + 1):
        if r in ws.row_dimensions:
            ws.row_dimensions[r].height = None

    ws.delete_rows(delete_from, amount=remove_count)

    for old_r, h in old_heights.items():
        ws.row_dimensions[old_r - remove_count].height = h

    for min_r, min_c, max_r, max_c in saved_merges:
        if max_r < delete_from:
            new_min_r, new_max_r = min_r, max_r
        elif min_r > base_end:
            new_min_r, new_max_r = min_r - remove_count, max_r - remove_count
        elif min_r >= delete_from and max_r <= base_end:
            # Silinen aralıktaki satır içi birleşmeler (örn: C5:E5)
            continue
        else:
            # Gün blokunu dikey kapsayan birleşme (örn: A3:A8, J3:J8)
            new_min_r, new_max_r = min_r, max(min_r, max_r - remove_count)

        if new_max_r > new_min_r or max_c > min_c:
            ws.merge_cells(
                start_row=new_min_r,
                start_column=min_c,
                end_row=new_max_r,
                end_column=max_c,
            )

    return remove_count


def _ensure_day_block_height(ws, start_row: int, total_rows: int, day: str) -> None:
    """Dikey gün adının kırpılmaması için daraltılmış blok yüksekliğini dengeler."""
    if total_rows <= 0:
        return
    min_block_height = len(day) * 5.6 + 4.0
    current_heights = [
        ws.row_dimensions[r].height or 19.5
        for r in range(start_row, start_row + total_rows)
    ]
    if sum(current_heights) < min_block_height:
        per_row = round(min_block_height / total_rows, 2)
        for idx, r in enumerate(range(start_row, start_row + total_rows)):
            ws.row_dimensions[r].height = max(current_heights[idx], per_row)


def export_single_week_excel(
    week_plan: dict[str, list[ScheduleTask]],
    student_name: str = "",
    week_title: str = "",
    template_path: Path = TEMPLATES_DIR / "haftalik_bos.xlsx",
    shrink_empty: bool = False,
) -> bytes:
    """HAFTALIK BOŞ DERS PROGRAMI.xlsx şablonunu doldurur ve bytes olarak döner."""
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    clean_week_title = week_title.strip()
    is_default_week = tr_upper(clean_week_title) in ("", "1. HAFTA")
    header_parts = [
        p
        for p in (student_name.strip(), "" if is_default_week else clean_week_title)
        if p
    ]
    if header_parts:
        ws["A1"].value = f"{tr_upper(' - '.join(header_parts))} - HAFTALIK DERS ÇALIŞMA PROGRAMI"

    target_sheet = "Haftalık Program" if is_default_week else clean_week_title
    target_title = _sanitize_sheet_title(target_sheet, ws.title)
    if target_title != ws.title:
        ws.title = "__temp_sheet__"
        ws.title = target_title

    base_blocks = _detect_day_blocks(ws)

    shift = 0
    for day, base_start, base_cap in base_blocks:
        start_row = base_start + shift
        tasks = week_plan.get(day, [])
        needed = max(1, len(tasks))
        extra = max(0, needed - base_cap)
        if extra > 0:
            insert_at = start_row + min(2, base_cap)
            _expand_day_block(
                ws,
                insert_at=insert_at,
                extra=extra,
                max_col=10,
                extra_row_merges=[(3, 5)],
            )
            shift += extra
            total_rows = base_cap + extra
        elif shrink_empty and needed < base_cap:
            removed = _shrink_day_block(
                ws,
                start_row=start_row,
                base_cap=base_cap,
                target_rows=needed,
                max_col=10,
            )
            shift -= removed
            total_rows = needed
            _ensure_day_block_height(ws, start_row, total_rows, day)
        else:
            total_rows = base_cap

        for idx in range(total_rows):
            r = start_row + idx
            if idx < len(tasks):
                t = tasks[idx]
                ws.cell(r, 2).value = t.ders
                ws.cell(r, 3).value = t.formatted_topic()
                ws.cell(r, 6).value = (
                    int(t.target_questions)
                    if t.target_questions is not None and t.target_questions > 0
                    else None
                )
            else:
                ws.cell(r, 2).value = None
                ws.cell(r, 3).value = None
                ws.cell(r, 6).value = None

        end_row = start_row + total_rows - 1
        ws.cell(start_row, 10).value = (
            f'=IF(SUM(G{start_row}:G{end_row})>0,SUM(G{start_row}:G{end_row}),"")'
        )

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def export_compact_four_weeks_excel(
    weeks: list[dict[str, list[ScheduleTask]]],
    student_name: str = "",
    week_labels: list[str] | None = None,
    template_path: Path = TEMPLATES_DIR / "kompakt_sablon.xlsx",
    shrink_empty: bool = False,
) -> bytes:
    """Ders_Calisma_Programi_Kompakt_Sablon.xlsx (4 hafta) şablonunu doldurur."""
    wb = openpyxl.load_workbook(template_path)

    padded_weeks = list(weeks[:4]) + [{} for _ in range(max(0, 4 - len(weeks)))]
    labels = list(week_labels or []) + [""] * 4

    sheet_pairs = [
        ("1-2. Hafta Programı", 0, 1, "1. HAFTA", "2. HAFTA"),
        ("3-4. Hafta Programı", 2, 3, "3. HAFTA", "4. HAFTA"),
    ]

    for sheet_name, left_idx, right_idx, default_left_lbl, default_right_lbl in sheet_pairs:
        ws = wb[sheet_name]

        if student_name.strip():
            ws["A1"].value = f"{tr_upper(student_name.strip())} - DERS ÇALIŞMA PROGRAMI"

        left_lbl = labels[left_idx].strip()
        right_lbl = labels[right_idx].strip()
        ws["A3"].value = (
            f"{default_left_lbl} ({left_lbl})"
            if left_lbl and tr_upper(left_lbl) != default_left_lbl
            else default_left_lbl
        )
        ws["J3"].value = (
            f"{default_right_lbl} ({right_lbl})"
            if right_lbl and tr_upper(right_lbl) != default_right_lbl
            else default_right_lbl
        )

        left_plan = padded_weeks[left_idx]
        right_plan = padded_weeks[right_idx]

        base_blocks = _detect_day_blocks(ws)
        base_last_day_row = max(start + cap - 1 for _, start, cap in base_blocks)

        shift = 0
        for day, base_start, base_cap in base_blocks:
            start_row = base_start + shift
            l_tasks = left_plan.get(day, [])
            r_tasks = right_plan.get(day, [])
            needed = max(1, len(l_tasks), len(r_tasks))
            extra = max(0, needed - base_cap)
            if extra > 0:
                _expand_day_block(
                    ws,
                    insert_at=start_row + min(2, base_cap),
                    extra=extra,
                    max_col=17,
                    extra_row_merges=None,
                )
                shift += extra
                total_rows = base_cap + extra
            elif shrink_empty and needed < base_cap:
                removed = _shrink_day_block(
                    ws,
                    start_row=start_row,
                    base_cap=base_cap,
                    target_rows=needed,
                    max_col=17,
                )
                shift -= removed
                total_rows = needed
                _ensure_day_block_height(ws, start_row, total_rows, day)
            else:
                total_rows = base_cap

            for idx in range(total_rows):
                r = start_row + idx
                if idx < len(l_tasks):
                    lt = l_tasks[idx]
                    ws.cell(r, 2).value = lt.ders
                    ws.cell(r, 3).value = lt.formatted_topic()
                    ws.cell(r, 4).value = (
                        int(lt.target_questions)
                        if lt.target_questions is not None and lt.target_questions > 0
                        else None
                    )
                else:
                    ws.cell(r, 2).value = None
                    ws.cell(r, 3).value = None
                    ws.cell(r, 4).value = None

                if idx < len(r_tasks):
                    rt = r_tasks[idx]
                    ws.cell(r, 11).value = rt.ders
                    ws.cell(r, 12).value = rt.formatted_topic()
                    ws.cell(r, 13).value = (
                        int(rt.target_questions)
                        if rt.target_questions is not None and rt.target_questions > 0
                        else None
                    )
                else:
                    ws.cell(r, 11).value = None
                    ws.cell(r, 12).value = None
                    ws.cell(r, 13).value = None

            end_row = start_row + total_rows - 1
            ws.cell(start_row, 8).value = (
                f'=IF(SUM(E{start_row}:E{end_row})>0,SUM(E{start_row}:E{end_row}),"")'
            )
            ws.cell(start_row, 17).value = (
                f'=IF(SUM(N{start_row}:N{end_row})>0,SUM(N{start_row}:N{end_row}),"")'
            )

        last_data_row = base_last_day_row + shift
        total_row = last_data_row + 1
        ws.cell(total_row, 4).value = f"=SUM(D5:D{last_data_row})"
        ws.cell(total_row, 13).value = f"=SUM(M5:M{last_data_row})"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def export_monthly_matrix_excel(
    matrix_weeks: list[dict[str, dict]],
    student_name: str = "",
    week_labels: list[str] | None = None,
    template_path: Path = TEMPLATES_DIR / "aylik_matris.xlsx",
) -> bytes:
    """AYLIK KAZANIM VE GELİŞİM MATRİSİ.xlsx şablonunu doldurur."""
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    if student_name.strip():
        ws["A1"].value = (
            f"{tr_upper(student_name.strip())} - 4 HAFTALIK DERS ÇALIŞMA ÇİZELGESİ"
        )

    labels = list(week_labels or []) + ["", "", "", ""]
    week_header_cols = [(2, "1. HAFTA"), (4, "2. HAFTA"), (6, "3. HAFTA"), (8, "4. HAFTA")]
    for idx, (col, default_hdr) in enumerate(week_header_cols):
        lbl = labels[idx].strip()
        if lbl and tr_upper(lbl) != default_hdr:
            ws.cell(2, col).value = f"{default_hdr} ({lbl})"

    # Şablondaki ders satırlarını tespit et (R03..R24)
    subject_rows: dict[str, int] = {}
    for r in range(3, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if isinstance(val, str):
            norm = tr_upper(val.strip())
            if norm in MATRIX_SUBJECTS:
                subject_rows[norm] = r

    padded_weeks = list(matrix_weeks[:4]) + [{} for _ in range(max(0, 4 - len(matrix_weeks)))]
    week_cols = [(2, 3), (4, 5), (6, 7), (8, 9)]  # (konu_col, soru_col)

    for subj, r in subject_rows.items():
        max_lines = 1
        for w_idx, (topic_col, q_col) in enumerate(week_cols):
            entry = padded_weeks[w_idx].get(subj, {})
            topics = [str(t).strip() for t in entry.get("topics", []) if str(t).strip()]
            text = "\n".join(topics) if topics else None
            ws.cell(r, topic_col).value = text
            if text:
                lines = sum(max(1, math.ceil(len(line) / 30)) for line in text.split("\n"))
                max_lines = max(max_lines, lines)

            q_val = entry.get("target_questions")
            ws.cell(r, q_col).value = int(q_val) if q_val and int(q_val) > 0 else None

        # Satır toplam hedef formülü (J sütunu)
        ws.cell(r, 10).value = f'=IF(SUM(C{r},E{r},G{r},I{r})>0,SUM(C{r},E{r},G{r},I{r}),"")'
        if max_lines > 1:
            ws.row_dimensions[r].height = max(19.5, 13 * max_lines + 6)

    # R25 alt toplam formülleri (eğer 25. satır toplam satırıysa)
    if subject_rows:
        first_r = min(subject_rows.values())
        last_r = max(subject_rows.values())
        total_r = last_r + 1
        for col_letter, col_idx in [("C", 3), ("E", 5), ("G", 7), ("I", 9), ("J", 10)]:
            ws.cell(total_r, col_idx).value = (
                f'=IF(SUM({col_letter}{first_r}:{col_letter}{last_r})>0,'
                f'SUM({col_letter}{first_r}:{col_letter}{last_r}),"")'
            )

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
