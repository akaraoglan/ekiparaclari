"""
KDV İadesi Kontrol Raporu - "Vermedi" veya "0.00" Olan 1.Alt Mükellefleri Çıkarıcı

2. Alt Mükellef tablosunda, "Karşılaştırmalı Alış-Satış Analizi Dönemindeki Mal ve
Hizmet Alım Karşılığını Teşkil Eden Bedel / 1. Alt Mükellefin Muhtasar Beyanname
Kontrolü" sütununda Vermedi veya 0.00 olan satırları getirir.
"""

import re
import subprocess
import os
import uuid
import pandas as pd


# ---------------------------------------------------------------------------
# PDF OKUMA
# ---------------------------------------------------------------------------

def pdf_metnini_oku_layout(pdf_yolu: str) -> str:
    """pdftotext -layout ile metin çıkarır (sütun hizaları korunur)."""
    try:
        result = subprocess.run(
            ["pdftotext", "-layout", pdf_yolu, "-"],
            capture_output=True, text=True, encoding="utf-8"
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout
    except FileNotFoundError:
        pass

    # Yedek: PyMuPDF
    try:
        import fitz
        doc = fitz.open(pdf_yolu)
        pages = []
        for page in doc:
            pages.append(page.get_text("text", sort=True))
        doc.close()
        return "\n".join(pages)
    except Exception:
        pass

    # Yedek: pdfplumber
    try:
        import pdfplumber
        with pdfplumber.open(pdf_yolu) as pdf:
            return "\n".join(
                p.extract_text() or "" for p in pdf.pages
            )
    except Exception as e:
        raise RuntimeError(f"PDF okunamadı: {e}")


# ---------------------------------------------------------------------------
# PARSING
# ---------------------------------------------------------------------------

def satirlari_birlestir(lines: list, start: int) -> str:
    """
    Çok satıra bölünmüş '1.Alt Mükellef:' girişini tek satıra birleştirir.
    PDF'de bazen VKN/Dönem/Tutar kısmı bir sonraki satıra taşar.
    """
    birlesik = lines[start].strip()

    # Durum 1: VKN yok → devam satırlarını birleştir
    if not re.search(r"\d{10,11}", birlesik):
        for j in range(start + 1, min(start + 5, len(lines))):
            sonraki = lines[j].strip()
            if sonraki and not sonraki.isdigit():
                birlesik = birlesik + " " + sonraki
                if re.search(r"\d{10,11}", birlesik):
                    break

    # Durum 2: VKN var ama dönem+tutar eksik
    if re.search(r"\d{10,11}", birlesik) and not re.search(r"/\s*\d{6}\s*/", birlesik):
        for j in range(start + 1, min(start + 5, len(lines))):
            sonraki = lines[j].strip()
            if sonraki and not sonraki.isdigit():
                birlesik = birlesik + " " + sonraki
                if re.search(r"/\s*\d{6}\s*/", birlesik):
                    break

    return birlesik


_TAM_RE = re.compile(
    r"1\.Alt Mükellef:\s*(.+?)\s*/\s*(\d{10,11})\s*/\s*(\d{6})\s*/\s*([\d.]+)",
    re.IGNORECASE,
)


def blok_bilgisi_coz(birlesik_satir: str):
    """'1.Alt Mükellef: ...' satırından cari, vkn, dönem, tutar çıkarır."""
    m = _TAM_RE.search(birlesik_satir)
    if m:
        return m.group(1).strip(), m.group(2), m.group(3), m.group(4)
    return None


_GURULTU = re.compile(
    r"https?://|"
    r"\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}|"
    r"2\.\s*ALT MÜKELLEFİN|"
    r"Unvanı / Adı|Karşılaştırmalı|"
    r"Vergi Kimlik|Noksan Beyan|"
    r"Döneminde|Kontrolü",
    re.IGNORECASE,
)


def kontrol_degeri_bul(lines: list, baslangic: int) -> str:
    """
    'Vermedi' veya '0.00' / '0,00' değerini blok içinde arar.
    Sayfa kırılması satırlarını atlar.
    """
    for j in range(baslangic + 1, min(baslangic + 45, len(lines))):
        satir = lines[j]
        s = satir.strip()

        # Bir sonraki bloğa geçtik
        if "1.Alt Mükellef:" in satir:
            break

        # Gürültü satırlarını atla
        if _GURULTU.search(satir):
            continue

        # Durum: Satır yalnızca "Vermedi" içeriyor
        if re.match(r"^Vermedi$", s, re.IGNORECASE):
            return "Vermedi"

        # Durum: Satır yalnızca "0.00" veya "0,00" içeriyor
        if re.match(r"^0[.,]00$", s):
            return s

        # Durum: Satır VKN içermiyor ama sonda 0.00 var
        if not re.search(r"\d{10,11}", satir):
            m = re.search(r"\b0[.,]00\s*$", satir)
            if m:
                return "0.00"

        # Gerçek pozitif sayısal değer → blok bu değil
        if re.match(r"^\d[\d.,]+$", s) and not re.match(r"^0[.,]00$", s):
            if re.match(r"^\d+$", s) and int(s) <= 999:
                continue
            return None

    return None


def _parca_temizle(parca: str) -> str:
    """Cari adı parçasından 0.00/Vermedi gibi gürültüleri temizler."""
    parca = re.sub(r"\b0[.,]00\b", "", parca)
    parca = re.sub(r"\bVermedi\b", "", parca, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", parca).strip()


def iki_alt_muk_bul(lines: list, baslangic: int, bir_alt_vkn: str) -> tuple:
    """
    Bloktaki 2.Alt Mükellef VKN ve cari adını bulur.
    VKN: satırda geçen 10-11 haneli sayı (1.Alt VKN'si hariç).
    Cari: VKN satırının sol tarafı + üst/alt komşu metin satırları.
    """
    vkn_satir_no = None
    vkn_degeri = ""

    for j in range(baslangic + 1, min(baslangic + 45, len(lines))):
        if "1.Alt Mükellef:" in lines[j]:
            break
        if _GURULTU.search(lines[j]):
            continue
        m = re.search(r"\b(\d{10,11})\b", lines[j])
        if m and m.group(1) != bir_alt_vkn:
            vkn_degeri = m.group(1)
            vkn_satir_no = j
            break

    if vkn_satir_no is None:
        return "", ""

    # VKN satırında solundaki metin
    sol = lines[vkn_satir_no][:lines[vkn_satir_no].find(vkn_degeri)].strip()

    # VKN satırının üstündeki cari parçaları
    yukari = []
    for j in range(vkn_satir_no - 1, max(baslangic, vkn_satir_no - 8), -1):
        s = lines[j].strip()
        if not s:
            break
        if _GURULTU.search(lines[j]) or re.search(r"\d{10,11}", s):
            break
        if re.match(r"^\d+[.,]\d", s):
            break
        if re.match(r"^\d+$", s) and int(s) <= 999:
            break
        yukari.insert(0, _parca_temizle(s))

    # VKN satırının altındaki cari parçaları
    asagi = []
    for j in range(vkn_satir_no + 1, min(vkn_satir_no + 6, len(lines))):
        s = lines[j].strip()
        if not s or _GURULTU.search(lines[j]):
            break
        if re.search(r"\d{10,11}", s) or re.match(r"^\d", s):
            break
        if "1.Alt Mükellef:" in lines[j]:
            break
        asagi.append(_parca_temizle(s))

    parcalar = yukari + ([_parca_temizle(sol)] if sol else []) + asagi
    cari = re.sub(r"\s+", " ", " ".join(p for p in parcalar if p)).strip()
    return cari, vkn_degeri


def pdf_analiz_et(pdf_yolu: str) -> tuple:
    """
    PDF'den '2. Alt Mükellef' tablosunu okuyup
    Vermedi / 0.00 olan 1.Alt Mükellefleri döndürür.
    Her kayda karşılık gelen 2.Alt Mükellef bilgisi de eklenir.
    """
    metin = pdf_metnini_oku_layout(pdf_yolu)
    lines = metin.split("\n")

    sonuc = []
    kontrol = []

    i = 0
    while i < len(lines):
        if "1.Alt Mükellef:" in lines[i]:
            birlesik = satirlari_birlestir(lines, i)
            cozum = blok_bilgisi_coz(birlesik)

            if cozum:
                cari, vkn, donem, tutar = cozum
                deger = kontrol_degeri_bul(lines, i)
                durum = "Vermedi/0.00" if deger else "Değer var"
                kontrol.append([i + 1, durum, cari, vkn, birlesik[:300]])

                if deger:
                    cari2, vkn2 = iki_alt_muk_bul(lines, i, vkn)
                    sonuc.append([cari, vkn, donem, tutar, deger, cari2, vkn2])
            else:
                kontrol.append([i + 1, "PARSE HATASI", "", "", birlesik[:300]])

        i += 1

    return sonuc, kontrol


# ---------------------------------------------------------------------------
# EXCEL YAZMA
# ---------------------------------------------------------------------------

def excele_yaz(sonuc: list, kontrol: list, cikti_yolu: str) -> None:
    """İki sayfalı Excel dosyası oluşturur: Vermedi_Sifir + Kontrol."""
    with pd.ExcelWriter(cikti_yolu, engine="openpyxl") as writer:
        if sonuc:
            df = pd.DataFrame(sonuc, columns=[
                "1.Alt Mükellef Cari Adı", "1.Alt Mükellef VKN",
                "Dönem", "Tutar", "Kontrol Değeri",
                "2.Alt Mükellef Cari Adı", "2.Alt Mükellef VKN",
            ])
            df.to_excel(writer, sheet_name="Vermedi_Sifir", index=False)

        if kontrol:
            df_k = pd.DataFrame(kontrol, columns=[
                "Satır No", "Durum", "Cari Adı", "Vergi No", "Ham Metin"
            ])
            df_k.to_excel(writer, sheet_name="Kontrol", index=False)

        # Sütun genişlikleri ve formatı
        from openpyxl.styles import Font, PatternFill
        for sheet_name, sheet in writer.sheets.items():
            for col in sheet.columns:
                max_len = max(
                    (len(str(c.value)) for c in col if c.value), default=10
                )
                sheet.column_dimensions[col[0].column_letter].width = min(max_len + 4, 80)

            # Başlık satırı formatı
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", start_color="366092")


# ---------------------------------------------------------------------------
# FLASK ENTEGRASYON
# ---------------------------------------------------------------------------

def process_vermedi_sifir_pdf(input_path: str, output_dir: str) -> tuple:
    """
    Ana işlev: PDF'yi işler ve sonuç veya kontrol bilgisi dönmek.
    Döner: (status, dosya_yolu, mesaj)
    """
    try:
        metin = pdf_metnini_oku_layout(input_path)
        if not metin.strip():
            return "error", None, "PDF'den metin okunamadı. Lütfen başka bir dosya seçiniz."

        sonuc, kontrol = pdf_analiz_et(input_path)

        if not sonuc:
            # Kontrol dosyasını yine de oluştur
            cikti = os.path.join(output_dir, f"vermedi_sifir_kontrol_{uuid.uuid4().hex[:8]}.xlsx")
            excele_yaz([], kontrol, cikti)
            hata_sayisi = sum(1 for k in kontrol if k[1] == "PARSE HATASI")
            mesaj = f"Vermedi/0.00 olan kayıt bulunamadı. Toplam blok: {len(kontrol)}, Hatalı: {hata_sayisi}"
            return "partial", cikti, mesaj

        # Sonuç dosyası oluştur
        cikti = os.path.join(output_dir, f"vermedi_sifir_raporlar_{uuid.uuid4().hex[:8]}.xlsx")
        excele_yaz(sonuc, kontrol, cikti)

        hata_sayisi = sum(1 for k in kontrol if k[1] == "PARSE HATASI")
        mesaj = f"Başarıyla işlendi. Bulunan: {len(sonuc)} kayıt, Hatalı blok: {hata_sayisi}"
        return "success", cikti, mesaj

    except Exception as e:
        return "error", None, f"Hata oluştu: {str(e)}"
