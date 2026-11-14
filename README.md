# YKS Çalışma Programı Planlayıcı

Öğrenciler için **2027 YKS (MEB 2018)** ve **2028 YKS (Türkiye Yüzyılı Maarif Modeli)** müfredatlarını kullanarak hızlıca haftalık / 4 haftalık ders çalışma programları hazırlayan ve doğrudan orijinal Excel şablonlarınıza aktaran Python web uygulaması.

🌐 **Canlı Kullanım (GitHub Pages):** https://hllzmz.github.io/yks-planla/  

## Özellikler

- **Gömülü Güncel Müfredatlar:**
  - `2027 YKS`: TYT, AYT, YDT (311 konu/alt başlık + MEB kazanım detayları)
  - `2028 YKS`: TYT, AYT, YDT (314 konu/alt başlık + Maarif Modeli öğrenme çıktıları)
- **Tıkla-Ata & Sürükle-Bırak Planlama:** Sağdaki müfredat ağacından **`+ Ekle`** / **`+ Üniteyi Ekle`** butonlarıyla veya **`⠿` tutamacından sürükle-bırak** yaparak konuları ve tüm üniteleri istediğiniz günün istediğiniz sırasına yerleştirme; gün içi ve günler arası görev taşıma.
- **Rutin Dağıtıcı & Hızlı Araçlar:** *"20 Paragraf"*, *"20 Problem"*, *"20 Geometri"* veya özel rutinleri seçili günlere tek tıkla dağıtma; **`🆕 Yeni Program`** ile sıfırdan yeni öğrenci programı başlatma.
- **Otomatik Hafta Tarihi (+7 Gün):** Takvim seçiciden (`📅`) veya 1. hafta başlığından tarih girildiğinde 2–4. hafta başlıklarını otomatik +7 gün ilerletme (`↻ +7 Gün`).
- **Otomatik Taslak Kaydı:** Çalışmalarınız tarayıcıda (`localStorage`) otomatik saklanır; sayfa yenilense bile kaybolmaz.
- **Üç Farklı Excel Export Formatı:**
  - **Tek Haftalık Şablon** (`templates/haftalik_bos.xlsx`)
  - **4 Haftalık Kompakt Şablon** (`templates/kompakt_sablon.xlsx` — `1-2. Hafta Programı` ve `3-4. Hafta Programı` sekmeleriyle)
  - **Aylık Gelişim Matrisi** (`templates/aylik_matris.xlsx` — 22 ders satırı üzerinde 4 haftalık konu ve soru hedefi matrisi)
  - Bir güne şablondaki kapasiteden fazla görev eklendiğinde satırları otomatik genişletir; isteğe bağlı **`Boş satırları daralt`** seçeneğiyle gün içindeki kullanılmayan boş satırları sıkıştırır.

## Kurulum ve Çalıştırma

```bash
# Sanal ortam oluşturun ve bağımlılıkları kurun
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# Uygulamayı başlatın
python app.py

# Testleri çalıştırın
python -m unittest discover -s tests -v
```

Ardından tarayıcınızda `http://localhost:8080` adresini açın.

## Proje Yapısı

- `app.py`: FastAPI tabanlı Python arka uç ve API sunucusu
- `web/index.html`: Harici bağımlılığı olmayan, tek dosyalık web arayüzü
- `curriculum.py`: 2027 & 2028 müfredat yükleme, Türkçe duyarlı filtreleme ve matris ders eşleme
- `excel_exporter.py`: `openpyxl` tabanlı şablon doldurma, dinamik satır genişletme ve daraltma motoru
- `data/`: JSON formatında gömülü 2027 ve 2028 müfredat verileri
- `templates/`: Orijinal tek haftalık, 4 haftalık kompakt ve aylık gelişim matrisi Excel şablonları
- `scripts/import_from_excel.py`: Kaynak Excel müfredat dosyalarından `data/*.json` üreten yardımcı script
- `tests/test_core.py`: Müfredat, Excel dışa aktarım ve API uç nokta testleri

## Lisans

Bu proje [MIT Lisansı](LICENSE) ile lisanslanmıştır. Kullanım ve dağıtım serbesttir; kopyalarda telif hakkı ve geliştirici atıf bildiriminin korunması gerekir.

