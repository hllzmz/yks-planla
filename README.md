# YKS Çalışma Programı Planlayıcı

Öğrenciler için **2027 YKS (MEB 2018)** ve **2028 YKS (Türkiye Yüzyılı Maarif Modeli)** müfredatlarını kullanarak hızlıca haftalık / 4 haftalık ders çalışma programları hazırlayan ve doğrudan orijinal Excel şablonlarınıza aktaran **Python (`FastAPI` + `openpyxl`)** web uygulaması.

## Özellikler

- **Gömülü Güncel Müfredatlar:**
  - `2027 YKS`: TYT, AYT, YDT (311 konu/alt başlık + MEB kazanım detayları)
  - `2028 YKS`: TYT, AYT, YDT (314 konu/alt başlık + Maarif Modeli öğrenme çıktıları)
- **Tıkla-Ata Hızlı Planlama:** Sol tarafta **Aktif Gün** seçip sağdaki müfredat ağacından **`+ Ekle`** (veya tüm üniteyi eklemek için **`+ Üniteyi Ekle`**) butonlarına basarak konuları anında haftaya yerleştirme.
- **Rutin Dağıtıcı:** *"20 Paragraf"*, *"20 Problem"*, *"20 Geometri"* veya özel rutinleri seçili günlere tek tıkla dağıtma.
- **Otomatik Taslak Kaydı:** Çalışmalarınız tarayıcıda (`localStorage`) otomatik saklanır; sayfa yenilense bile kaybolmaz.
- **İki Farklı Excel Export Formatı:**
  - **Tek Haftalık Şablon** (`templates/haftalik_bos.xlsx`)
  - **4 Haftalık Kompakt Şablon** (`templates/kompakt_sablon.xlsx` — `1-2. Hafta Programı` ve `3-4. Hafta Programı` sekmeleriyle)
  - Bir güne şablondaki boş satır kapasitesinden fazla görev eklense dahi birleşik hücreleri, kenarlıkları ve toplam formüllerini otomatik korur.

## Kurulum ve Çalıştırma

```bash
# Sanal ortam oluşturun ve bağımlılıkları kurun
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# Uygulamayı başlatın
python app.py
```

Ardından tarayıcınızda `http://localhost:8080` adresini açın.

## Proje Yapısı

- `app.py`: FastAPI tabanlı Python arka uç ve API sunucusu
- `web/index.html`: Harici bağımlılığı olmayan, tek dosyalık web arayüzü
- `curriculum.py`: 2027 & 2028 müfredat yükleme ve Türkçe duyarlı filtreleme
- `excel_exporter.py`: `openpyxl` tabanlı şablon doldurma ve dinamik satır genişletme motoru
- `data/`: JSON formatında gömülü 2027 ve 2028 müfredat verileri
- `templates/`: Orijinal tek haftalık ve 4 haftalık kompakt Excel şablonları
- `scripts/import_from_excel.py`: Kaynak Excel müfredat dosyalarından `data/*.json` üreten yardımcı script
