"""2027 ve 2028 YKS müfredat Excel dosyalarını JSON formatına dönüştürür."""

import json
from pathlib import Path
import openpyxl


def clean_str(val) -> str:
    if val is None:
        return ""
    return " ".join(str(val).strip().split())


def parse_curriculum_workbook(xlsx_path: Path, year: int) -> list[dict]:
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    items: list[dict] = []

    sheet_specs = [
        ("TYT Müfredatı", "TYT"),
        ("AYT Müfredatı", "AYT"),
        ("YDT Müfredatı", "YDT"),
    ]

    for sheet_name, session in sheet_specs:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue

        for row in rows[1:]:
            if not row or row[0] is None:
                continue
            sira_raw = str(row[0]).strip()
            if not sira_raw.replace(".", "", 1).isdigit():
                continue

            sira_no = int(float(sira_raw))

            if session == "TYT":
                # Sıra No, Oturum, Test Grubu, Ders Adı, Sınıf Düzeyi, Ünite, Konu, Kazanım
                puan_turu = "TYT (Ortak)"
                test_grubu = clean_str(row[2])
                ders = clean_str(row[3])
                sinif = clean_str(row[4])
                unite = clean_str(row[5])
                konu = clean_str(row[6])
                kazanim = clean_str(row[7]) if len(row) > 7 else ""
                agirlik = ""
            else:
                # Sıra No, Oturum, Hedef Puan Türü, Test Grubu, Ders Adı, Sınıf Düzeyi, Ünite, Konu, Kazanım, [ÖSYM Ağırlık]
                puan_turu = clean_str(row[2])
                test_grubu = clean_str(row[3])
                ders = clean_str(row[4])
                sinif = clean_str(row[5])
                unite = clean_str(row[6])
                konu = clean_str(row[7])
                kazanim = clean_str(row[8]) if len(row) > 8 else ""
                agirlik = clean_str(row[9]) if len(row) > 9 else ""

            items.append(
                {
                    "id": f"{year}-{session}-{sira_no:03d}",
                    "year": year,
                    "sira": sira_no,
                    "session": session,
                    "puan_turu": puan_turu,
                    "test_grubu": test_grubu,
                    "ders": ders,
                    "sinif": sinif,
                    "unite": unite,
                    "konu": konu,
                    "kazanim": kazanim,
                    "agirlik": agirlik,
                }
            )

    return items


def main() -> None:
    root = Path(__file__).resolve().parent.parent
    data_dir = root / "data"
    data_dir.mkdir(parents=True, exist_ok=True)

    sources = [
        (2027, Path("/tmp/2027_mufredat.xlsx"), data_dir / "mufredat_2027.json"),
        (2028, Path("/tmp/2028_mufredat.xlsx"), data_dir / "mufredat_2028.json"),
    ]

    for year, src_path, dst_path in sources:
        if not src_path.exists():
            raise FileNotFoundError(f"Kaynak dosya bulunamadı: {src_path}")
        items = parse_curriculum_workbook(src_path, year)
        payload = {
            "year": year,
            "total_topics": len(items),
            "topics": items,
        }
        dst_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"{year} müfredatı yazıldı -> {dst_path} ({len(items)} konu)")


if __name__ == "__main__":
    main()
