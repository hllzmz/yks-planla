"""Hazırlanan ders çalışma programını orijinal Excel şablonlarına aktaran modül."""

from __future__ import annotations

import io
import re
from copy import copy
from dataclasses import dataclass
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import Cell
from openpyxl.styles import Alignment, Font

from curriculum import DAYS_OF_WEEK, tr_upper

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
    return (cleaned[:31] if cleaned else fallback)


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

    # 1. Mevcut birleşik hücre koordinatlarını kaydet ve çöz
    saved_merges: list[tuple[int, int, int, int]] = [
        (m.min_row, m.min_col, m.max_row, m.max_col)
        for m in list(ws.merged_cells.ranges)
    ]
    for min_r, min_c, max_r, max_c in saved_merges:
        ws.unmerge_cells(
            start_row=min_r, start_column=min_c, end_row=max_r, end_column=max_c
        )

    # 2. Satırları ekle (insert_at satırından önce araya eklenir)
    template_row = insert_at - 1
    template_height = ws.row_dimensions[template_row].height

    # Aşağıdaki satır yüksekliklerini kaydet
    old_heights = {
        r: ws.row_dimensions[r].height
        for r in range(insert_at, ws.max_row + 1)
        if ws.row_dimensions[r].height is not None
    }

    ws.insert_rows(insert_at, amount=extra)

    # Eklenen satırlara stil ve yükseklik kopyala, kayan satırların yüksekliklerini güncelle
    for offset in range(extra):
        r = insert_at + offset
        if template_height is not None:
            ws.row_dimensions[r].height = template_height
        for col in range(1, max_col + 1):
            _copy_cell_style(ws.cell(template_row, col), ws.cell(r, col))

    for old_r, h in old_heights.items():
        ws.row_dimensions[old_r + extra].height = h

    # 3. Birleşik hücreleri güncelle
    updated_merges: list[tuple[int, int, int, int]] = []
    for min_r, min_c, max_r, max_c in saved_merges:
        if max_r < insert_at:
            updated_merges.append((min_r, min_c, max_r, max_c))
        elif min_r >= insert_at:
            updated_merges.append((min_r + extra, min_c, max_r + extra, max_c))
        else:
            # insert_at bu birleşik bloğun tam içinde (örn. PAZARTESİ A5:A11)
            updated_merges.append((min_r, min_c, max_r + extra, max_c))

    # Yeni eklenen satırlardaki satır içi birleştirmeler (ör. haftalik_bos C:E)
    if extra_row_merges:
        for offset in range(extra):
            r = insert_at + offset
            for c1, c2 in extra_row_merges:
                updated_merges.append((r, c1, r, c2))

    for min_r, min_c, max_r, max_c in updated_merges:
        ws.merge_cells(
            start_row=min_r, start_column=min_c, end_row=max_r, end_column=max_c
        )


def export_single_week_excel(
    week_plan: dict[str, list[ScheduleTask]],
    student_name: str = "",
    week_title: str = "",
    template_path: Path = TEMPLATES_DIR / "haftalik_bos.xlsx",
) -> bytes:
    """HAFTALIK BOŞ DERS PROGRAMI.xlsx şablonunu doldurur ve bytes olarak döner."""
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # Başlık (A1:J3 birleşik alanında logoların ortasında görünür)
    header_parts = [p.strip() for p in (student_name, week_title) if p and p.strip()]
    if header_parts:
        ws["A1"].value = tr_upper(" - ".join(header_parts)) + "\nHAFTALIK DERS ÇALIŞMA PROGRAMI"
        ws["A1"].font = Font(name="Calibri", size=13, bold=True)
        ws["A1"].alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    if week_title.strip():
        target_title = _sanitize_sheet_title(week_title, ws.title)
        if target_title != ws.title:
            ws.title = "__temp_sheet__"
            ws.title = target_title

    # Orijinal gün blokları: (gün_adı, başlangıç_satırı, satır_kapasitesi)
    base_blocks = [
        ("PAZARTESİ", 5, 7),
        ("SALI", 12, 6),
        ("ÇARŞAMBA", 18, 6),
        ("PERŞEMBE", 24, 6),
        ("CUMA", 30, 6),
        ("CUMARTESİ", 36, 7),
        ("PAZAR", 43, 7),
    ]

    shift = 0
    for day, base_start, base_cap in base_blocks:
        start_row = base_start + shift
        tasks = week_plan.get(day, [])
        extra = max(0, len(tasks) - base_cap)
        if extra > 0:
            # Blok içindeki 2. satırdan sonra araya ekle (üst/alt kalın kenarlıkları korumak için)
            insert_at = start_row + 2
            _expand_day_block(
                ws,
                insert_at=insert_at,
                extra=extra,
                max_col=10,
                extra_row_merges=[(3, 5)],  # C:E birleştirmesi
            )
            shift += extra

        total_rows = base_cap + extra
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

        # Sağdaki TOPLAM ÇÖZÜLEN SORU S. (J sütunu) için formül
        end_row = start_row + total_rows - 1
        ws.cell(start_row, 10).value = f"=IF(SUM(G{start_row}:G{end_row})>0,SUM(G{start_row}:G{end_row}),\"\")"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def export_compact_four_weeks_excel(
    weeks: list[dict[str, list[ScheduleTask]]],
    student_name: str = "",
    week_labels: list[str] | None = None,
    template_path: Path = TEMPLATES_DIR / "kompakt_sablon.xlsx",
) -> bytes:
    """Ders_Calisma_Programi_Kompakt_Sablon.xlsx (4 hafta) şablonunu doldurur."""
    wb = openpyxl.load_workbook(template_path)

    # En az 4 hafta olacak şekilde tamamla
    padded_weeks = list(weeks[:4]) + [{} for _ in range(max(0, 4 - len(weeks)))]
    labels = list(week_labels or []) + [""] * 4

    sheet_pairs = [
        ("1-2. Hafta Programı", 0, 1, "1. HAFTA", "2. HAFTA"),
        ("3-4. Hafta Programı", 2, 3, "3. HAFTA", "4. HAFTA"),
    ]

    base_blocks = [
        ("PAZARTESİ", 5, 7),
        ("SALI", 12, 6),
        ("ÇARŞAMBA", 18, 6),
        ("PERŞEMBE", 24, 6),
        ("CUMA", 30, 6),
        ("CUMARTESİ", 36, 7),
        ("PAZAR", 43, 6),
    ]

    for sheet_name, left_idx, right_idx, default_left_lbl, default_right_lbl in sheet_pairs:
        ws = wb[sheet_name]

        if student_name.strip():
            ws["C1"].value = f"{tr_upper(student_name.strip())} - DERS ÇALIŞMA PROGRAMI"

        left_lbl = labels[left_idx].strip()
        right_lbl = labels[right_idx].strip()
        ws["A3"].value = (
            f"{default_left_lbl} ({left_lbl})" if left_lbl else default_left_lbl
        )
        ws["J3"].value = (
            f"{default_right_lbl} ({right_lbl})" if right_lbl else default_right_lbl
        )

        left_plan = padded_weeks[left_idx]
        right_plan = padded_weeks[right_idx]

        shift = 0
        for day, base_start, base_cap in base_blocks:
            start_row = base_start + shift
            l_tasks = left_plan.get(day, [])
            r_tasks = right_plan.get(day, [])
            needed = max(len(l_tasks), len(r_tasks))
            extra = max(0, needed - base_cap)
            if extra > 0:
                _expand_day_block(
                    ws,
                    insert_at=start_row + 2,
                    extra=extra,
                    max_col=17,
                    extra_row_merges=None,
                )
                shift += extra

            total_rows = base_cap + extra
            for idx in range(total_rows):
                r = start_row + idx
                # Sol hafta: B=Ders, C=Konu, D=Hedef Soru
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

                # Sağ hafta: K=Ders, L=Konu, M=Hedef Soru
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
            ws.cell(start_row, 8).value = f"=IF(SUM(E{start_row}:E{end_row})>0,SUM(E{start_row}:E{end_row}),\"\")"
            ws.cell(start_row, 17).value = f"=IF(SUM(N{start_row}:N{end_row})>0,SUM(N{start_row}:N{end_row}),\"\")"

        # Alt toplam satırı (49 + shift)
        total_row = 49 + shift
        last_data_row = total_row - 1
        ws.cell(total_row, 4).value = f"=SUM(D5:D{last_data_row})"
        ws.cell(total_row, 13).value = f"=SUM(M5:M{last_data_row})"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
