# AGENTS.md

Bu dosya, Codex'in bu projede calisirken proje mantigini her seferinde yeniden anlatmaya gerek kalmadan anlamasi icin hazirlandi. Buradaki notlari README ile birlikte dikkate al.

## Proje Ozeti

Bu proje Flask tabanli, kucuk ekip kullanimina yonelik bir web uygulamasidir.

Ana sayfa (`/`) kullaniciyi iki bolume yonlendirir:

- `Ekip Araclari`: Python ile yazilmis PDF, XML, Excel ve KDV odakli araclarin listelendigi bolum.
- `Rehber`: Telefonda tutulmak istenmeyen firma/iletisim numaralarini kategori, firma adi ve etiketlerle aramaya yarayan HTML rehber.

Projeye domain, sunucu adresi, sifre, token, gercek ortam yolu veya benzeri gizli bilgiler eklenmemelidir.

## Uygulama Yapisi

- `app.py`: Ana Flask uygulamasi, route'lar, dosya yukleme/indirme akisleri.
- `tool_registry.py`: Ekip Araclari ekraninda gorunen arac kartlari, kategoriler, anahtar kelimeler ve URL'ler.
- `tools/`: Her aracin asil Python is mantigi.
- `templates/`: Flask/Jinja HTML sablonlari.
- `static/css/style.css`: Ekip Araclari ortak tasarimi.
- `static/js/main.js`: Tema degistirme ve bazi form davranislari.
- `uploads/`: Calisma aninda yuklenen dosyalar. Git'e eklenmez; 24 saatten eski dosyalar otomatik temizlenir.
- `outputs/`: Uretilen indirme dosyalari. Git'e eklenmez; 24 saatten eski dosyalar otomatik temizlenir.

`uploads/` ve `outputs/` runtime klasorleridir. Kullanici dosyalari ve uretilen raporlar kalici veri gibi dusunulmemeli; repo'ya dahil edilmemelidir.

## Ana Akislar

### Ana sayfa

`/` route'u `templates/home.html` dosyasini gosterir. Bu sayfa Ekip Araclari ve Rehber kartlarini sunar.

### Ekip Araclari

`/ekip-araclari` route'u `tool_registry.TOOLS` listesinden araclari alir. Kategori filtresi ve arama sorgusu burada uygulanir.

Yeni bir arac eklerken genel sira:

1. Aracin is mantigini `tools/` altinda ayri bir `.py` dosyasina koy.
2. Gerekli route'u `app.py` icine ekle.
3. Aracin HTML formunu `templates/` altinda olustur; mumkunse `base.html` uzerinden ilerle.
4. Arac kartini `tool_registry.py` icindeki `TOOLS` listesine ekle.
5. Yeni kutuphane gerekiyorsa `requirements.txt` dosyasini guncelle.

### Rehber

`/rehber` route'u `templates/rehber.html` dosyasini gosterir. Rehber su an tek dosyalik, bagimsiz bir HTML/JS sayfasi gibi calisir ve CSV verisini Google Sheets yayin linkinden PapaParse ile okur.

Rehberde firma adi, kategori, adres, telefon ve etiket alanlari istemci tarafinda filtrelenir. Rehberle calisirken gercek telefon numaralarini veya ozel kisi/firma verilerini gereksiz yere kod, dokumantasyon veya commit mesajlarina tasima.

## Mevcut Arac Kategorileri

PDF araclari:

- PDF ayirma
- PDF birlestirme
- PDF dondurme
- PDF sayfa cikarma
- PDF icine PDF ekleme

KDV Iadesi Kontrol Raporu araclari:

- 1. Alt Mukellef bilgilerini PDF'den Excel'e cikarma
- Vermedi veya 0.00 olan alt mukellefleri bulma

E-Fatura araclari:

- E-fatura XML dosyalarini Excel'e donusturme
- Irsaliye XML dosyalarini Excel'e donusturme
- Giden fatura XML dosyalarindan irsaliye numarasi cikarma

Starwood araclari:

- VakifBank ekstresinde gumruk vergi tahsilati satirlarini beyanname tipine gore renklendirme
- Detay ve ozet Excel dosyalarindan ihrac kayitli satis faturasi listesi hazirlama

## Dosya Isleme Kurallari

- Yuklenen dosyalar `uploads/` altina UUID eklenmis guvenli dosya adlariyla kaydedilir.
- Cikti dosyalari `outputs/` altinda uretilir ve `send_file(..., as_attachment=True)` ile indirilir.
- `tools.common.secure_tr_filename` Turkce karakterli dosya adlarini guvenli ASCII dosya adlarina cevirir.
- `tools.common.cleanup_old_files` `uploads/` ve `outputs/` altindaki 24 saatten eski dosyalari temizler.
- Yeni araclarda gecici veya islenmis dosya gerekiyorsa proje icindeki bu ortak runtime klasorleri kullan; sistem temp klasorune kalici takip edilemeyen cikti birakma.
- Coklu ciktilarda `tools.common.zip_files` kullanilabilir.

## Linux VPS Notlari

Gelistirme Windows ortaminda yapilabilir, ancak yayin ortami Linux VPS'tir. Degisiklik yaparken platform bagimliligini mutlaka dusun.

- Dosya yollari icin string birlestirmek yerine `os.path` veya `pathlib` kullan.
- Windows'a ozel `os.startfile`, COM, Outlook otomasyonu, OneDrive yerel yolu veya masaustu uygulama entegrasyonu sunucuda calismayabilir.
- Boyle bir ozellik gerekiyorsa Linux'ta sessizce hata vermeyecek sekilde opsiyonel yap ve kullaniciya anlasilir mesaj dondur.
- PDF metin cikarma tarafinda `pdftotext` gibi harici Linux paketleri bulunmayabilir. Kodda PyMuPDF/pdfplumber yedekleri var; yeni arac yazarken benzer fallback mantigi kullan.
- Yeni kutuphanelerin Linux'ta kurulabilir oldugunu kontrol et ve gerekirse sistem paketi ihtiyacini dokumante et.
- Sunucuda guncelleme akisi genelde GitHub'a push, VPS'te pull ve servis restart seklindedir. Repo icine domain veya sunucuya ozel gizli bilgi yazma.

## Bagimliliklar

Temel Python bagimliliklari `requirements.txt` icindedir:

- Flask
- pypdf
- pandas
- openpyxl
- pymupdf
- pdfplumber

Yeni bir Python paketi kullanildiginda `requirements.txt` guncellenmelidir. Sistem seviyesinde paket gerekiyorsa bunu kodun icine gizli varsayim olarak koyma; not olarak belirt.

## Kod Stili ve Bakim

- Arac mantigini mumkun oldugunca `tools/` altinda tut; `app.py` route ve HTTP akisini yonetsin.
- Kullaniciya gosterilen mesajlar Turkce ve anlasilir olmali.
- Turkce karakterleri koru; dosyalari UTF-8 olarak kaydet.
- Hata durumlarinda Flask `flash` mesajlariyla kullaniciya net bilgi ver.
- Kullanici dosyalarini, uretilen raporlari, `.env` dosyalarini ve gizli bilgileri commit etme.
- Sadece ilgili aracin dosyalarini degistir; alakasiz refactor yapma.

## Arayuz Notlari

Ekip Araclari bolumu sade ve is odaklidir:

- Kategori sol menusu `tool_registry.py` kategorilerinden gelir.
- Arama, arac adi/aciklama/kategori/anahtar kelimelerde calisir.
- Arac sayfalari genellikle `base.html` sablonunu ve `static/css/style.css` stilini kullanir.
- Tema degistirme `static/js/main.js` ile localStorage uzerinden calisir.

Rehber sayfasi ayri bir tasarima sahiptir ve kendi CSS/JS kodunu `templates/rehber.html` icinde tasir.

## Dogrulama

Kod degisikliginden sonra, mumkunse en az su kontrolleri yap:

```bash
python -m py_compile app.py tool_registry.py tools/*.py
```

Arayuz veya route degisikliginde Flask uygulamasini yerelde calistirip ilgili sayfayi manuel kontrol et:

```bash
python app.py
```

PDF, XML veya Excel araclarinda degisiklik yapildiysa uygun ornek dosyayla gercek cikti uretmeyi dene. Ornek dosya yoksa en azindan import/compile kontrolunu yap ve test edilemeyen kismi kullaniciya acikca soyle.
