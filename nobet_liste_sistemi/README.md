# Nöbet Liste Sistemi

Tek klasörde çalışan, ay/yıl/bina bazlı nöbet listesi yönetim modülü.

## Klasör yapısı

```
nobet_liste_sistemi/
  index.asp                 → Halka açık liste sayfası
  ayarlar.asp               → Modül ayarları
  admin/
    login.asp               → Yönetici girişi
    panel.asp               → Dosya durumu ve yönetim
    yukle.asp               → Dosya yükleme
    dosya_sil.asp           → Dosya silme
  database/
    connection.asp          → Veritabanı bağlantısı
    schema.sql              → Tablo kurulum scripti
  lib/                      → Ortak fonksiyonlar
  listeler/                 → Yüklenen dosyalar
    2026/
      pendik/
        haziran/
          pendik_acil_radyoloji_nobet_listesi.pdf
      basibuyuk/
        haziran/
          asaf_ataseven_idari_hekim_nobet_listesi.pdf
```

## Kurulum

1. `nobet_liste_sistemi` klasörünü sunucuya kopyalayın:
   `inetpub/wwwroot/Admin/nobet_liste_sistemi/`

2. `database/connection.asp` içindeki SQL Server bilgilerini güncelleyin.

3. `database/schema.sql` scriptini veritabanında çalıştırın.

4. `ayarlar.asp` içindeki `MODUL_WEB_YOLU` değerini kontrol edin:
   `/Admin/nobet_liste_sistemi/`

5. IIS uygulama havuzuna `nobet_liste_sistemi`, `listeler` ve `tmp` klasörlerinde **okuma/yazma** izni verin.

6. Dosya yükleme için sunucuda **Persits.Upload** veya **ABCUpload** bileşenlerinden biri kurulu olmalıdır.

## Kullanım

| Sayfa | Adres |
|-------|-------|
| Liste sayfası | `/Admin/nobet_liste_sistemi/index.asp` |
| Yönetici paneli | `/Admin/nobet_liste_sistemi/admin/login.asp` |

**Varsayılan giriş:** `admin` / `admin123` (kurulumdan sonra değiştirin)

## Mantık

- Sistem yılı → `listeler/2026/`
- Bina → `pendik` veya `basibuyuk`
- Ay → `haziran`, `temmuz` vb. (otomatik oluşur)
- Dosya varsa link **aktif**, yoksa **pasif**
- Yönetici panelinden dosya yüklenince link otomatik aktif olur

## Portal entegrasyonu

Mevcut portal menüsündeki nöbet listeleri linkini şu adrese yönlendirin:

`/Admin/nobet_liste_sistemi/index.asp`
