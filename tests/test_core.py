"""Müfredat ve Excel aktarım modülleri için temel testler."""

import io
import unittest
import openpyxl

from curriculum import DAYS_OF_WEEK, CurriculumStore
from excel_exporter import (
    ScheduleTask,
    export_compact_four_weeks_excel,
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
            for i in range(1, 10)  # 9 görev -> 7 satırlık bloku taşırır
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

    def test_compact_four_weeks_export(self) -> None:
        w1 = {day: [] for day in DAYS_OF_WEEK}
        w1["SALI"] = [
            ScheduleTask("AYT Fizik", f"Kuvvet {i}", "Konu", 30)
            for i in range(1, 9)  # 8 görev -> 6 satırlık Salı'yı genişletir
        ]
        raw = export_compact_four_weeks_excel(
            [w1, {}, {}, {}],
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


if __name__ == "__main__":
    unittest.main()
