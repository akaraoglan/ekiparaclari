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
    }
]

def get_categories():
    return sorted({tool["category"] for tool in TOOLS})
