"""YKS Çalışma Programı Planlayıcı (FastAPI + openpyxl - Saf Python Arka Uç)."""

from __future__ import annotations

import os
import re
from pathlib import Path
from urllib.parse import quote

import uvicorn
from fastapi import FastAPI, Query
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from curriculum import DAYS_OF_WEEK, STUDY_TYPES, CurriculumStore
from excel_exporter import (
    ScheduleTask,
    export_compact_four_weeks_excel,
    export_single_week_excel,
)

BASE_DIR = Path(__file__).resolve().parent
WEB_DIR = BASE_DIR / "web"

STORE = CurriculumStore()

app = FastAPI(title="YKS Program Planlayıcı")
app.mount("/data", StaticFiles(directory=BASE_DIR / "data"), name="data")


# ------------------------------------------------------------------ #
# Veri Modelleri
# ------------------------------------------------------------------ #
class TaskPayload(BaseModel):
    ders: str
    konu: str
    study_type: str = "Soru Çözümü"
    target_questions: int | None = None
    note: str = ""


class ExportPayload(BaseModel):
    student_name: str = ""
    template_mode: str = "single"  # "single" | "compact"
    week_labels: list[str] = Field(
        default_factory=lambda: ["1. Hafta", "2. Hafta", "3. Hafta", "4. Hafta"]
    )
    weeks: list[dict[str, list[TaskPayload]]] = Field(default_factory=list)


def safe_filename(name: str) -> str:
    cleaned = re.sub(r"[^a-zA-Z0-9çğıöşüÇĞİÖŞÜ _-]", "", name).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned or "Ogrenci"


def _to_schedule_week(raw_week: dict[str, list[TaskPayload]]) -> dict[str, list[ScheduleTask]]:
    result: dict[str, list[ScheduleTask]] = {day: [] for day in DAYS_OF_WEEK}
    for day in DAYS_OF_WEEK:
        for t in raw_week.get(day, []):
            result[day].append(
                ScheduleTask(
                    ders=t.ders,
                    konu=t.konu,
                    study_type=t.study_type,
                    target_questions=t.target_questions,
                    note=t.note,
                )
            )
    return result


# ------------------------------------------------------------------ #
# Sayfa ve API Rotaları
# ------------------------------------------------------------------ #
@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    html = (WEB_DIR / "index.html").read_text(encoding="utf-8")
    return HTMLResponse(content=html)


@app.get("/api/meta")
def get_meta() -> dict:
    return {
        "days": DAYS_OF_WEEK,
        "study_types": STUDY_TYPES,
        "years": [2027, 2028],
    }


@app.get("/api/subjects")
def get_subjects(
    year: int = Query(2027),
    session: str = Query("TYT"),
    puan_turu: str = Query("Tümü"),
) -> dict:
    topics = STORE.filter_topics(
        year=year, session=session, puan_turu=puan_turu, ders="Tümü", query=""
    )
    counts: dict[str, int] = {}
    for t in topics:
        counts[t.ders] = counts.get(t.ders, 0) + 1
    return {
        "subjects": [{"ders": k, "count": v} for k, v in counts.items()],
        "total": len(topics),
    }


@app.get("/api/curriculum")
def get_curriculum(
    year: int = Query(2027),
    session: str = Query("TYT"),
    puan_turu: str = Query("Tümü"),
    ders: str = Query("Tümü"),
    q: str = Query(""),
) -> dict:
    topics = STORE.filter_topics(
        year=year, session=session, puan_turu=puan_turu, ders=ders, query=q
    )
    grouped = STORE.group_by_subject_and_unit(topics)
    sections = []
    for subject_key, units in grouped.items():
        unit_list = []
        for unit_name, unit_topics in units.items():
            unit_list.append(
                {
                    "unite": unit_name,
                    "topics": [
                        {
                            "id": t.id,
                            "session": t.session,
                            "ders": t.ders,
                            "full_ders": t.full_ders_label,
                            "sinif": t.sinif,
                            "puan_turu": t.puan_turu,
                            "unite": t.unite,
                            "konu": t.konu,
                            "kazanim": t.kazanim,
                            "agirlik": t.agirlik,
                        }
                        for t in unit_topics
                    ],
                }
            )
        sections.append({"subject": subject_key, "units": unit_list})
    return {"total": len(topics), "sections": sections}


@app.post("/api/export")
def export_excel(payload: ExportPayload) -> Response:
    base = safe_filename(payload.student_name)
    labels = list(payload.week_labels) + ["1. Hafta", "2. Hafta", "3. Hafta", "4. Hafta"]

    if payload.template_mode == "single":
        week0 = _to_schedule_week(payload.weeks[0] if payload.weeks else {})
        data = export_single_week_excel(
            week_plan=week0,
            student_name=payload.student_name,
            week_title=labels[0],
        )
        w_part = safe_filename(labels[0])
        filename = f"{base}_{w_part}_Program.xlsx"
    else:
        weeks_conv = [
            _to_schedule_week(payload.weeks[i] if i < len(payload.weeks) else {})
            for i in range(4)
        ]
        data = export_compact_four_weeks_excel(
            weeks=weeks_conv,
            student_name=payload.student_name,
            week_labels=labels[:4],
        )
        filename = f"{base}_4_Haftalik_Kompakt_Program.xlsx"

    encoded_name = quote(filename)
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_name}"
        },
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")
