"""2027 ve 2028 YKS müfredat verilerini yükleyen ve filtreleyen modül."""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

DATA_DIR = Path(__file__).resolve().parent / "data"

DAYS_OF_WEEK = [
    "PAZARTESİ",
    "SALI",
    "ÇARŞAMBA",
    "PERŞEMBE",
    "CUMA",
    "CUMARTESİ",
    "PAZAR",
]

STUDY_TYPES = [
    "Konu",
    "Soru Çözümü",
    "Tekrar",
    "Deneme",
]

MATRIX_SUBJECTS = [
    "TÜRKÇE",
    "TEMEL MATEMATİK",
    "TYT GEOMETRİ",
    "TYT FİZİK",
    "TYT KİMYA",
    "TYT BİYOLOJİ",
    "TYT TARİH",
    "TYT COĞRAFYA",
    "TYT FELSEFE",
    "TYT DİN KÜLTÜRÜ",
    "AYT MATEMATİK",
    "AYT GEOMETRİ",
    "AYT FİZİK",
    "AYT KİMYA",
    "AYT BİYOLOJİ",
    "TÜRK DİLİ VE EDEBİYATI",
    "AYT TARİH 1",
    "AYT TARİH 2",
    "AYT COĞRAFYA 1",
    "AYT COĞRAFYA 2",
    "AYT FELSEFE GRUBU",
    "AYT DİN KÜLTÜRÜ",
]


def map_to_matrix_subject(
    session: str, ders: str, preferred_subject: str | None = None
) -> str | None:
    """TYT/AYT ders adını Aylık Gelişim Matrisi satırına eşler."""
    s = session.strip().upper()
    d = ders.strip()
    if s == "TYT":
        tyt_map = {
            "Türkçe": "TÜRKÇE",
            "Matematik": "TEMEL MATEMATİK",
            "Geometri": "TYT GEOMETRİ",
            "Fizik": "TYT FİZİK",
            "Kimya": "TYT KİMYA",
            "Biyoloji": "TYT BİYOLOJİ",
            "Tarih": "TYT TARİH",
            "Coğrafya": "TYT COĞRAFYA",
            "Felsefe": "TYT FELSEFE",
            "Din Kültürü": "TYT DİN KÜLTÜRÜ",
        }
        return tyt_map.get(d)
    if s == "AYT":
        if d == "Tarih":
            return (
                preferred_subject
                if preferred_subject in ("AYT TARİH 1", "AYT TARİH 2")
                else "AYT TARİH 1"
            )
        if d == "Coğrafya":
            return (
                preferred_subject
                if preferred_subject in ("AYT COĞRAFYA 1", "AYT COĞRAFYA 2")
                else "AYT COĞRAFYA 1"
            )
        ayt_map = {
            "Matematik": "AYT MATEMATİK",
            "Geometri": "AYT GEOMETRİ",
            "Fizik": "AYT FİZİK",
            "Kimya": "AYT KİMYA",
            "Biyoloji": "AYT BİYOLOJİ",
            "Türk Dili ve Edebiyatı": "TÜRK DİLİ VE EDEBİYATI",
            "Felsefe": "AYT FELSEFE GRUBU",
            "Mantık": "AYT FELSEFE GRUBU",
            "Psikoloji": "AYT FELSEFE GRUBU",
            "Sosyoloji": "AYT FELSEFE GRUBU",
            "Din Kültürü": "AYT DİN KÜLTÜRÜ",
        }
        return ayt_map.get(d)
    return preferred_subject if preferred_subject in MATRIX_SUBJECTS else None


def tr_lower(text: str) -> str:
    """Türkçe karakter duyarlı küçük harf dönüşümü."""
    return (
        text.replace("İ", "i")
        .replace("I", "ı")
        .replace("Ş", "ş")
        .replace("Ğ", "ğ")
        .replace("Ü", "ü")
        .replace("Ö", "ö")
        .replace("Ç", "ç")
        .lower()
    )


def tr_upper(text: str) -> str:
    """Türkçe karakter duyarlı büyük harf dönüşümü."""
    return (
        text.replace("i", "İ")
        .replace("ı", "I")
        .replace("ş", "Ş")
        .replace("ğ", "Ğ")
        .replace("ü", "Ü")
        .replace("ö", "Ö")
        .replace("ç", "Ç")
        .upper()
    )


@dataclass(frozen=True)
class Topic:
    id: str
    year: int
    sira: int
    session: str
    puan_turu: str
    test_grubu: str
    ders: str
    sinif: str
    unite: str
    konu: str
    kazanim: str
    agirlik: str = ""

    @property
    def full_ders_label(self) -> str:
        """Excel DERS sütunu için kısa ve anlaşılır etiket (örn: TYT Matematik)."""
        return f"{self.session} {self.ders}"


class CurriculumStore:
    def __init__(self, data_dir: Path = DATA_DIR) -> None:
        self._topics_by_year: dict[int, list[Topic]] = {}
        for year in (2027, 2028):
            path = data_dir / f"mufredat_{year}.json"
            raw = json.loads(path.read_text(encoding="utf-8"))
            self._topics_by_year[year] = [Topic(**item) for item in raw["topics"]]

    def get_topics(self, year: int) -> list[Topic]:
        return self._topics_by_year.get(year, [])

    def get_subjects(
        self,
        year: int,
        session: str = "Tümü",
        puan_turu: str = "Tümü",
    ) -> list[str]:
        topics = self.filter_topics(
            year=year,
            session=session,
            puan_turu=puan_turu,
            ders="Tümü",
            query="",
        )
        seen: list[str] = []
        for t in topics:
            if t.ders not in seen:
                seen.append(t.ders)
        return seen

    def filter_topics(
        self,
        year: int,
        session: str = "Tümü",
        puan_turu: str = "Tümü",
        ders: str = "Tümü",
        query: str = "",
    ) -> list[Topic]:
        results = self.get_topics(year)
        if session and session != "Tümü":
            results = [t for t in results if t.session == session]

        if puan_turu and puan_turu != "Tümü":
            if puan_turu == "DİL":
                results = [
                    t
                    for t in results
                    if t.session == "TYT" or "DİL" in t.puan_turu
                ]
            else:
                results = [
                    t
                    for t in results
                    if t.session == "TYT" or puan_turu in t.puan_turu
                ]

        if ders and ders != "Tümü":
            results = [t for t in results if t.ders == ders]

        q = tr_lower(query.strip())
        if q:
            results = [
                t
                for t in results
                if (
                    q in tr_lower(t.konu)
                    or q in tr_lower(t.unite)
                    or q in tr_lower(t.ders)
                    or q in tr_lower(t.kazanim)
                )
            ]
        return results

    def group_by_subject_and_unit(
        self, topics: list[Topic]
    ) -> dict[str, dict[str, list[Topic]]]:
        """Konuları (Oturum + Ders) -> Ünite -> [Konu] şeklinde gruplar."""
        grouped: dict[str, dict[str, list[Topic]]] = {}
        for t in topics:
            subject_key = f"{t.session} · {t.ders}"
            unit_dict = grouped.setdefault(subject_key, {})
            unit_dict.setdefault(t.unite, []).append(t)
        return grouped
