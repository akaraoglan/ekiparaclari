"""
KDV İadesi Kontrol Raporu - 1.Alt Mükellef Çıkarıcı
"""

import re
import pandas as pd
from tools.common import zip_files
import os
import uuid


# ---------------------------------------------------------------------------
# PDF OKUMA
# ---------------------------------------------------------------------------

def pdf_metnini_oku(pdf_yolu: str) -> str:
    """PDF'den düz metin okur. PyMuPDF → pdfplumber sırasıyla dener."""
    metinler = []

    # 1) PyMuPDF
    try:
        import fitz
        doc = fitz.open(pdf_yolu)
        for page in doc:
            txt = page.get_text("text", sort=True)
            if txt and txt.strip():
                metinler.append(txt)
        doc.close()
        if metinler:
            return "\n".join(metinler)
    except Exception:
        pass

    # 2) pdfplumber yedek
    try:
        import pdfplumber
        with pdfplumber.open(pdf_yolu) as pdf:
            for sayfa in pdf.pages:
                txt = sayfa.extract_text()
                if txt and txt.strip():
                    metinler.append(txt)
        return "\n".join(metinler)
    except Exception as e:
        raise RuntimeError(f"PDF okunamadı: {e}")


# ---------------------------------------------------------------------------
# METİN TEMİZLEME
# ---------------------------------------------------------------------------

def temizle_metin(text: str) -> str:
    """Temel unicode/whitespace temizliği."""
    if not text:
        return ""
    text = text.replace("\xa0", " ")   # non-breaking space
    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")
    return text


def duzlestir(text: str) -> str:
    """Tüm whitespace'leri tek boşluğa indirir (satır sonları dahil)."""
    return re.sub(r"\s+", " ", text).strip()


# ---------------------------------------------------------------------------
# BLOK AYIRMA
# ---------------------------------------------------------------------------

_MARKER = re.compile(
    r"1\s*\.\s*Alt\s+M[üu]kellef\s*:",
    re.IGNORECASE | re.UNICODE,
)


def bloklari_bul(metin: str) -> list:
    """Her '1.Alt Mükellef:' işaretçisinden bir sonrakine kadar olan metin parçasını blok olarak döndürür."""
    metin = temizle_metin(metin)
    eslesmeler = list(_MARKER.finditer(metin))

    bloklar = []
    for i, m in enumerate(eslesmeler):
        start = m.end()
        end = eslesmeler[i + 1].start() if i + 1 < len(eslesmeler) else len(metin)
        blok = metin[start:end].strip()
        if blok:
            bloklar.append(blok)

    return bloklar


# ---------------------------------------------------------------------------
# TEK BLOKTAN KAYIT ÇÖZME
# ---------------------------------------------------------------------------

_VKN_RE = re.compile(r"\b(\d{10,11})\b")
_DONEM_RE = re.compile(r"\b(\d{6})\b")


def bloktan_kayit_coz(blok: str):
    """Bloğu düzleştirir ve içinden Cari Adı ile VKN'yi çıkarmaya çalışır."""
    tek = duzlestir(blok)
    tek = re.sub(r"^\d+\s+", "", tek).strip()

    # Desen 1
    m = re.search(
        r"^(?P<cari>.+?)"
        r"\s*/\s*"
        r"(?:\d+\s+)?"
        r"(?P<vkn>\d{10,11})"
        r"\s*/\s*"
        r"(?P<donem>\d{6})"
        r"\s*/",
        tek,
        re.IGNORECASE,
    )

    if m:
        cari = _cari_temizle(m.group("cari"))
        vkn = m.group("vkn").strip()
        if cari:
            return cari, vkn, tek

    # Desen 2 (fallback)
    m2 = re.search(
        r"/\s*(?:\d+\s+)?(?P<vkn>\d{10,11})\s*/\s*(?P<donem>\d{6})\s*/",
        tek,
    )
    if m2:
        vkn = m2.group("vkn").strip()
        cari = _cari_temizle(tek[: m2.start()])
        if cari:
            return cari, vkn, tek

    # Desen 3
    tum_vknler = _VKN_RE.findall(tek)
    tum_donemler = _DONEM_RE.findall(tek)
    gecerli_donemler = [d for d in tum_donemler if int(d) > 202000]

    if tum_vknler and gecerli_donemler:
        vkn = tum_vknler[0]
        vkn_pos = tek.find(vkn)
        cari_ham = tek[:vkn_pos]
        slash_pos = cari_ham.rfind("/")
        cari = _cari_temizle(cari_ham[:slash_pos] if slash_pos != -1 else cari_ham)
        if cari:
            return cari, vkn, tek

    return None


def _cari_temizle(cari: str) -> str:
    """Cari adından baştaki/sondaki istenmeyen karakterleri temizler."""
    cari = cari.strip(" /\t\n")
    cari = re.sub(r"^\d+\s+", "", cari)
    cari = re.sub(r"\s+", " ", cari)
    return cari.strip()


# ---------------------------------------------------------------------------
# TOPLU İŞLEM
# ---------------------------------------------------------------------------

def alt_mukellefleri_cek(metin: str) -> tuple:
    """Metinden tüm 1.Alt Mükellefleri çıkarır. Döner: (benzersiz_liste, kontrol_listesi)"""
    bloklar = bloklari_bul(metin)

    sonuc = []
    kontrol = []

    for idx, blok in enumerate(bloklar, start=1):
        cozum = bloktan_kayit_coz(blok)

        if cozum:
            cari, vkn, ham = cozum
            sonuc.append([cari, vkn])
            kontrol.append([idx, "OK", cari, vkn, ham[:600]])
        else:
            temiz = duzlestir(blok)
            kontrol.append([idx, "HATA", "", "", temiz[:600]])

    # VKN bazlı tekilleştirme
    benzersiz = []
    gorulen_vkn = {}

    for row in sonuc:
        cari, vkn = row
        if vkn not in gorulen_vkn:
            gorulen_vkn[vkn] = len(benzersiz)
            benzersiz.append([cari, vkn])
        else:
            mevcut_idx = gorulen_vkn[vkn]
            if len(cari) > len(benzersiz[mevcut_idx][0]):
                benzersiz[mevcut_idx][0] = cari

    return benzersiz, kontrol


# ---------------------------------------------------------------------------
# EXCEL YAZMA
# ---------------------------------------------------------------------------

def excele_yaz(veriler: list, kontrol_listesi: list, dosya_yolu: str) -> None:
    """İki sayfalı Excel dosyası oluşturur: Sonuc + Kontrol."""
    with pd.ExcelWriter(dosya_yolu, engine="openpyxl") as writer:
        if veriler:
            df = pd.DataFrame(veriler, columns=["Cari Adı", "Vergi No"])
            df.to_excel(writer, sheet_name="Sonuc", index=False)

        if kontrol_listesi:
            df_kontrol = pd.DataFrame(
                kontrol_listesi,
                columns=["Blok No", "Durum", "Cari Adı", "Vergi No", "Ham Blok"],
            )
            df_kontrol.to_excel(writer, sheet_name="Kontrol", index=False)

        # Sütun genişliklerini otomatik ayarla
        for sheet in writer.sheets.values():
            for col in sheet.columns:
                max_len = max((len(str(c.value)) for c in col if c.value), default=10)
                sheet.column_dimensions[col[0].column_letter].width = min(max_len + 4, 80)


# ---------------------------------------------------------------------------
# ANA İŞLEV (Flask Integration)
# ---------------------------------------------------------------------------

def process_kdviadesi_pdf(input_path: str, output_dir: str) -> tuple:
    """
    Ana işlev: PDF'yi işler ve sonuç veya kontrol bilgisi dönmek.
    Döner: (status, dosya_yolu, mesaj)
    """
    try:
        metin = pdf_metnini_oku(input_path)
        if not metin.strip():
            return "error", None, "PDF'den metin okunamadı. Lütfen başka bir dosya seçiniz."

        veriler, kontrol = alt_mukellefleri_cek(metin)

        if not veriler:
            # Kontrol dosyasını yine de oluştur
            cikti = os.path.join(output_dir, f"kdviadesi_kontrol_{uuid.uuid4().hex[:8]}.xlsx")
            excele_yaz([], kontrol, cikti)
            hata_sayisi = sum(1 for k in kontrol if k[1] == "HATA")
            mesaj = f"Hiç kayıt bulunamadı. Toplam blok: {len(kontrol)}, Hatalı: {hata_sayisi}"
            return "partial", cikti, mesaj

        # Sonuç dosyası oluştur
        cikti = os.path.join(output_dir, f"kdviadesi_raporlar_{uuid.uuid4().hex[:8]}.xlsx")
        excele_yaz(veriler, kontrol, cikti)

        hata_sayisi = sum(1 for k in kontrol if k[1] == "HATA")
        mesaj = f"Başarıyla işlendi. Bulunan: {len(veriler)} kayıt, Hatalı blok: {hata_sayisi}"
        return "success", cikti, mesaj

    except Exception as e:
        return "error", None, f"Hata oluştu: {str(e)}"
