# Ekip Araçları - Kategorili Sade Yapı

Bu sürüm küçük ekip kullanımı için sade tutulmuştur.

## Yapı

- `app.py` -> ana Flask uygulaması
- `tool_registry.py` -> araç kartları ve kategoriler
- `tools/` -> her araç ayrı `.py` dosyası
- `templates/` -> sayfalar
- `static/` -> CSS ve JS

## Yeni araç ekleme

1. `tools/` içine yeni aracın `.py` dosyasını ekleyin.
2. Route'u `app.py` içine ekleyin.
3. Aracın kart ve kategori bilgisini `tool_registry.py` içine ekleyin.
4. Gerekirse ilgili HTML şablonunu oluşturun.

## Kurulum

```bash
pip install -r requirements.txt
python app.py
```

# OCR desteği

Taranmış gümrük tahsilat PDF'lerindeki beyanname numarası RapidOCR ile okunur.
Gerekli OCR paketleri `requirements.txt` üzerinden kurulur; ayrıca bir masaüstü
uygulaması veya Tesseract kurulması gerekmez. Sunucuda proje bağımlılıklarını
yeniden kurmak yeterlidir.

Minimal Linux sunucularda OpenCV'nin ihtiyaç duyduğu `libGL1` paketi de gerekebilir:

```bash
sudo apt-get update
sudo apt-get install -y libgl1
python -m rapidocr check
```
