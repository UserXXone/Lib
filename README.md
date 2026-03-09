# Lib - Excel Uyumlu Kütüphane Programı

Python Tkinter ile yazılmış, modernleştirilmiş (simgeli) arayüze sahip kütüphane takip uygulaması.

## Özellikler (Sürüm 1.1)
- Modern görünüm ve simgeli sol menü
- Ana ekran istatistik kartları (Toplam / Rafta / Ödünçte / Geciken)
- Raf durum tablosu
- Sağ altta canlı tarih-saat (saniyeli), Türkçe gün adı ile
- Kitap sorgulama (metin arama + durum filtresi)
- Kitap ekleme
- Excel/CSV dosyasından toplu kitap içe aktarma
- Import sırasında sütun eşleştirme ekranı (hangi sütun = kitap adı/yazar/raf/qr)
- İşlem ekranlarında yüzdelik progress bar
- Kitap düzenleme / silme
- Öğrenci kontrol paneli (numara + isim ile manuel kaydet/güncelle/sil)
- QR / Kitap ID ile ödünç verme
- QR / Kitap ID ile iade alma
- QR / Kitap ID ile teslim süresi uzatma
- Geciken kitapları listeleme
- Tek tıkla tam yedek alma (`yedekler/` klasörüne uygulamadaki tüm temel dosyalar paketlenir)
- Menüden güvenli çıkış butonu
- Alt+1 kısayolu ile veritabanını temizleme
- Veriler `kutuphane.csv` dosyasında saklanır (Excel ile açılıp düzenlenebilir)

## Çalıştırma
```bash
python library_app.py
```

İlk açılışta örnek veriler otomatik oluşturulur.


### Import sütun notu
Örnek görseldeki dosya için önerilen eşleştirme:
- `Kitap Adı` → `aciklama`
- `Yazar` → `DUZELTILMIS KATEGORI` (istersen `hazirlayan`)
- `Raf` → `RAF`
- `QR` (opsiyonel) → `QR`
- QR importunda tüm `-` karakterleri otomatik `*` olarak dönüştürülür
