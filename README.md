# Excel Rapor Otomasyonu

Farklı şubelerden **farklı formatlarda** gelen satış Excel'lerini tek komutla birleştirir, temizler ve grafikli bir özet rapor üretir. Her ay elle saatler süren iş birkaç saniyeye iner.

*Merges messy sales spreadsheets from multiple branches (different column names, date formats, typos, duplicates) into one clean, formatted Excel report with charts — in one click.*

## Çözdüğü sorunlar

| Gerçek hayattaki sorun | Örnek | Program ne yapıyor |
|---|---|---|
| Her şube farklı sütun adı kullanıyor | `Adet` / `miktar` / `ADET` | Hepsini tek standart ada eşler |
| Farklı tarih formatları | `01.09.2026` / `2026-09-01` / `01/09/2026` | Hepsini doğru okur (gün-ay karışmaz) |
| Ürün adı yazım farkları | `Çay 1kg` / `çay 1kg` / ` Çay 1kg ` | Aynı ürün olarak birleştirir |
| Araya girmiş boş satırlar | | Siler |
| Aynı kaydın iki kez girilmesi | | Siler |
| Bozuk satırlar | metin yazılmış adet alanı | Atlar ve kaydını tutar |

## Üretilen rapor

- **Özet**: toplam ciro, satış adedi, en çok satan ürün, şube sıralaması ve grafik
- **Ürünler**: ürün bazında adet/ciro ve grafik
- **Ürün x Şube**: hangi şubede hangi ürün ne kadar sattı
- **Günlük**: günlük ciro trendi (çizgi grafik)
- **Temiz Veri**: birleştirilmiş, filtrelenebilir tüm kayıtlar
- **İşlem Kaydı**: kaç satır silindi/atlandı. Müşteri hiçbir verinin sessizce kaybolmadığını görür

## Kullanım

```bash
pip install -r requirements.txt
python ornek_veri_olustur.py   # demo için örnek dağınık dosyalar (girdi/)
python rapor.py                # cikti/rapor_YYYY-AA-GG.xlsx
```

Kendi klasörünle: `python rapor.py --girdi "D:/Satışlar/Eylül" --cikti eylul.xlsx`

**Windows**: `girdi` klasörüne dosyaları at, `calistir.bat`'a çift tıkla.

---

Benzer bir ihtiyacınız mı var? Excel/CSV raporlama, veri birleştirme, otomatik e-posta ile rapor gönderimi gibi işler için iletişime geçin.
