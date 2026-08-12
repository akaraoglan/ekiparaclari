TOOLS = [
    {
        "id": "pdfayir",
        "name": "PDF Ayır",
        "category": "PDF Araçları",
        "description": "PDF dosyasından belirli sayfaları alır veya tek tek sayfalara böler.",
        "keywords": ["pdf", "ayır", "split", "sayfa", "böl"],
        "url": "/pdf/ayir"
    },
    {
        "id": "pdfbirlestir",
        "name": "PDF Birleştir",
        "category": "PDF Araçları",
        "description": "Birden fazla PDF dosyasını tek bir PDF dosyasında birleştirir.",
        "keywords": ["pdf", "birleştir", "merge"],
        "url": "/pdf/birlestir"
    },
    {
        "id": "pdfdondur",
        "name": "PDF Döndür",
        "category": "PDF Araçları",
        "description": "PDF dosyasındaki tüm sayfaları veya seçilen sayfaları döndürür.",
        "keywords": ["pdf", "döndür", "rotate", "çevir"],
        "url": "/pdf/dondur"
    },
    {
        "id": "pdfcikar",
        "name": "PDF Çıkar",
        "category": "PDF Araçları",
        "description": "PDF'den istenmeyen sayfaları siler; geri kalan sayfalar yeni bir PDF olarak indirilir.",
        "keywords": ["pdf", "çıkar", "sil", "sayfa", "kaldır", "remove"],
        "url": "/pdf/cikar"
    },
    {
        "id": "pdfekle",
        "name": "PDF Ekle",
        "category": "PDF Araçları",
        "description": "Bir PDF dosyasını ana PDF'nin istediğiniz sayfasından sonra ekler.",
        "keywords": ["pdf", "ekle", "insert", "sayfa", "araya", "eklemek"],
        "url": "/pdf/ekle"
    },
    {
        "id": "kdviadesi_kontrol",
        "name": "KDV İadesi Mükellef Çıkartıcı",
        "category": "KDV İadesi Kontrol Raporu",
        "description": "KDV İadesi Kontrol Raporlarından 1. Alt Mükellef bilgilerini çıkararak Excel'e aktarır.",
        "keywords": ["kdv", "iades", "mükellef", "kontrol", "rapor", "çıkart"],
        "url": "/kdv/iades-kontrol"
    },
    {
        "id": "kdvermedi_sifir",
        "name": "Vermeyenler ve Sıfırlar",
        "category": "KDV İadesi Kontrol Raporu",
        "description": "KDV İadesi Kontrol Raporundan Vermedi veya 0.00 olan 1. Alt Mükellefleri bulur.",
        "keywords": ["kdv", "iades", "vermedi", "sıfır", "0.00", "mükellef"],
        "url": "/kdv/vermedi-sifir"
    },
    {
        "id": "xml_to_excel",
        "name": "E-Fatura XML → Excel",
        "category": "E-Fatura Araçları",
        "description": "Gelen e-fatura XML dosyalarını Excel formatına dönüştürür. Birden fazla dosya seçilebilir.",
        "keywords": ["xml", "excel", "efatura", "fatura", "dönüştür", "xlsx"],
        "url": "/xml/efatura"
    },
    {
        "id": "irsaliye_xml_to_excel",
        "name": "İrsaliye XML → Excel",
        "category": "E-Fatura Araçları",
        "description": "İrsaliye XML dosyalarını Excel'e dönüştürür. Müşteri adı, tarih, irsaliye no, ürün, miktar ve birim bilgilerini çıkarır.",
        "keywords": ["xml", "irsaliye", "excel", "ürün", "miktar", "müşteri", "dönüştür"],
        "url": "/xml/irsaliye-excel"
    },
    {
        "id": "irsaliye_no",
        "name": "Giden Fatura İrsaliye No Çıkartıcı",
        "category": "E-Fatura Araçları",
        "description": "Giden e-fatura XML dosyalarından irsaliye numaralarını çıkararak Excel'e aktarır. Birden fazla dosya seçilebilir.",
        "keywords": ["xml", "irsaliye", "giden", "fatura", "efatura", "numara", "çıkart", "waybill"],
        "url": "/xml/irsaliye-no"
    },
    {
        "id": "ekstre_boyama",
        "name": "Ekstre Boyama ve Gönderme İşlemi",
        "category": "Starwood",
        "description": "VakıfBank ekstresindeki Gümrük Vergi Tahsilatı satırlarını beyanname tipine göre renklendirir (IM=sarı, EX=mavi), OneDrive'a kaydeder ve Outlook taslak maili açar.",
        "keywords": ["ekstre", "vakıfbank", "boya", "gümrük", "beyanname", "im", "ex", "outlook", "starwood"],
        "url": "/starwood/ekstre-boyama"
    },
    {
        "id": "ihrac_kayitli_hazirlama",
        "name": "İhraç Kayıtlı Hazırlama",
        "category": "Starwood",
        "description": "Detay ve Özet Excel dosyalarından ihraç kayıtlı satış faturası listesini hazırlar.",
        "keywords": ["ihrac", "ihraç", "kayıtlı", "starwood", "excel", "detay", "özet", "gtip"],
        "url": "/starwood/ihrac-kayitli-hazirlama"
    },
    {
        "id": "ithalde_indirilecek_kdv",
        "name": "İthalde İndirilecek KDV Listesi",
        "category": "Starwood",
        "description": "Muavin ve İthalat Raporu dosyalarından İthalde İndirilecek KDV listesini hazırlar; eşleşmeyen ve farklı birimli kayıtları incelemeye ayırır.",
        "keywords": ["ithalat", "ithalde", "indirilecek", "kdv", "muavin", "beyanname", "ggb", "starwood", "excel"],
        "url": "/starwood/ithalde-indirilecek-kdv"
    },
    {
        "id": "ithaldeindirilecekfinal",
        "name": "ithaldeindirilecekfinal",
        "category": "Starwood",
        "description": "İthalde İndirilecek KDV Excel dosyasına %20 KDV ve fark kontrolü ekler; gerektiğinde matrahı düzeltir ve I/J sütunlarını değer olarak sabitler.",
        "keywords": ["ithaldeindirilecekfinal", "ithalde", "indirilecek", "kdv", "matrah", "fark", "starwood", "excel"],
        "url": "/starwood/ithaldeindirilecekfinal"
    }
]

def get_categories():
    return sorted({tool["category"] for tool in TOOLS})
