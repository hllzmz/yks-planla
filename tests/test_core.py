"""Müfredat ve Excel aktarım modülleri için temel testler."""

import io
import unittest
import zipfile
import openpyxl

from curriculum import DAYS_OF_WEEK, MATRIX_SUBJECTS, CurriculumStore, map_to_matrix_subject
from excel_exporter import (
    ScheduleTask,
    export_compact_four_weeks_excel,
    export_monthly_matrix_excel,
    export_single_week_excel,
)


class TestCurriculumAndExporter(unittest.TestCase):
    def setUp(self) -> None:
        self.store = CurriculumStore()

    def test_2027_and_2028_loaded(self) -> None:
        self.assertGreater(len(self.store.get_topics(2027)), 300)
        self.assertGreater(len(self.store.get_topics(2028)), 300)

    def test_turkish_search_and_filters(self) -> None:
        # Büyük/küçük Türkçe İ/i eşleşmesi
        res = self.store.filter_topics(year=2027, session="AYT", query="trigonometri")
        self.assertTrue(any("Trigonometri" in t.unite for t in res))

        res_dil = self.store.filter_topics(year=2028, session="YDT", puan_turu="DİL")
        self.assertGreater(len(res_dil), 20)

    def test_single_week_excel_export_with_overflow(self) -> None:
        plan = {day: [] for day in DAYS_OF_WEEK}
        plan["PAZARTESİ"] = [
            ScheduleTask("TYT Matematik", f"Konu {i}", "Soru Çözümü", 25)
            for i in range(1, 10)  # 9 görev -> 6 satırlık bloku taşırır
        ]
        plan["PAZAR"] = [
            ScheduleTask("Genel Deneme", "TYT Deneme", "Deneme", 120)
        ]

        raw = export_single_week_excel(
            plan, student_name="Zeynep Kaya", week_title="16 Şubat"
        )
        wb = openpyxl.load_workbook(io.BytesIO(raw))
        ws = wb.active
        self.assertEqual(ws.title, "16 Şubat")
        self.assertIn("ZEYNEP KAYA", ws["A1"].value)
        self.assertEqual(ws["B5"].value, "TYT Matematik")
        self.assertEqual(ws["B13"].value, "TYT Matematik")  # 9. satır (5+8=13)
        self.assertEqual(ws["F5"].value, 25)
        self.assertEqual(ws["B45"].value, "Genel Deneme")  # 3 satır kaymış Pazar başlangıcı (42+3=45)
        with zipfile.ZipFile(io.BytesIO(raw)) as zf:
            self.assertIn("xl/media/image1.png", zf.namelist())

    def test_compact_four_weeks_export(self) -> None:
        w1 = {day: [] for day in DAYS_OF_WEEK}
        w1["SALI"] = [
            ScheduleTask("AYT Fizik", f"Kuvvet {i}", "Konu", 30)
            for i in range(1, 9)  # 8 görev -> 6 satırlık Salı'yı genişletir
        ]
        w3 = {day: [] for day in DAYS_OF_WEEK}
        w3["CUMA"] = [
            ScheduleTask("TYT Türkçe", f"Paragraf {i}", "Soru Çözümü", 20)
            for i in range(1, 9)  # 3-4. Hafta'da Cuma 6 satır -> 2 satır taşar
        ]
        raw = export_compact_four_weeks_excel(
            [w1, {}, w3, {}],
            student_name="Ali Demir",
            week_labels=["16 Şub", "23 Şub", "2 Mar", "9 Mar"],
        )
        wb = openpyxl.load_workbook(io.BytesIO(raw))
        self.assertEqual(wb.sheetnames, ["1-2. Hafta Programı", "3-4. Hafta Programı"])
        ws12 = wb["1-2. Hafta Programı"]
        self.assertIn("ALİ DEMİR", ws12["C1"].value)
        # Salı 2 satır uzadığı için toplam satırı 49 -> 51 olmalı
        self.assertEqual(ws12["D51"].value, "=SUM(D5:D50)")
        self.assertEqual(ws12["M51"].value, "=SUM(M5:M50)")

        ws34 = wb["3-4. Hafta Programı"]
        # 3-4. Hafta taban toplam satırı 47, Cuma 2 satır uzayınca 49 olmalı
        self.assertEqual(ws34["D49"].value, "=SUM(D5:D48)")
        self.assertEqual(ws34["M49"].value, "=SUM(M5:M48)")

    def test_monthly_matrix_export_and_mapping(self) -> None:
        self.assertEqual(map_to_matrix_subject("TYT", "Matematik"), "TEMEL MATEMATİK")
        self.assertEqual(map_to_matrix_subject("AYT", "Tarih", "AYT TARİH 2"), "AYT TARİH 2")
        self.assertEqual(len(MATRIX_SUBJECTS), 22)

        matrix_weeks = [
            {
                "TÜRKÇE": {
                    "topics": ["Sözcükte Anlam", "Cümlede Anlam", "Paragraf Taktikleri"],
                    "target_questions": 180,
                },
                "AYT MATEMATİK": {
                    "topics": ["Türev", "İntegral"],
                    "target_questions": 120,
                },
            },
            {
                "TÜRKÇE": {"topics": ["Ses Bilgisi"], "target_questions": 100},
            },
            {},
            {},
        ]
        raw = export_monthly_matrix_excel(
            matrix_weeks=matrix_weeks,
            student_name="Elif Yılmaz",
            week_labels=["16 Şub", "23 Şub", "2 Mar", "9 Mar"],
        )
        wb = openpyxl.load_workbook(io.BytesIO(raw))
        ws = wb.active
        self.assertIn("ELİF YILMAZ", ws["A1"].value)
        self.assertEqual(ws["B2"].value, "1. HAFTA (16 Şub)")
        self.assertEqual(ws["B3"].value, "Sözcükte Anlam\nCümlede Anlam\nParagraf Taktikleri")
        self.assertEqual(ws["C3"].value, 180)
        self.assertEqual(ws["D3"].value, "Ses Bilgisi")
        self.assertEqual(ws["E3"].value, 100)
        self.assertEqual(ws["J3"].value, '=IF(SUM(C3,E3,G3,I3)>0,SUM(C3,E3,G3,I3),"")')
        self.assertEqual(ws["B13"].value, "Türev\nİntegral")
        self.assertEqual(ws["C13"].value, 120)
        self.assertEqual(ws["C25"].value, '=IF(SUM(C3:C24)>0,SUM(C3:C24),"")')
        self.assertEqual(ws["J25"].value, '=IF(SUM(J3:J24)>0,SUM(J3:J24),"")')
        self.assertGreater(ws.row_dimensions[3].height or 0, 19.5)


if __name__ == "__main__":
    unittest.main()
