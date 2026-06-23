# Nöbet Liste Sistemi

Tek klasörde çalışan, ay/yıl/bina bazlı nöbet listesi yönetim modülü.  
Veritabanı: **Microsoft Access** (`database/nobet_liste.mdb`) — otomatik oluşturulur.

## Klasör yapısı

```
nobet_liste_sistemi/
  index.asp                 → Halka açık liste sayfası
  kurulum.asp               → İlk kurulum (veritabanı oluşturur)
  ayarlar.asp               → Modül ayarları
  admin/                    → Yönetici paneli
  database/
    nobet_liste.mdb         → Otomatik oluşur
  listeler/
    2026/pendik/haziran/    → Dosyalar
    2026/basibuyuk/haziran/
```

## Kurulum

1. `nobet_liste_sistemi` klasörünü sunucuya kopyalayın
2. IIS'e `database`, `listeler`, `tmp` klasörlerinde yazma izni verin
3. Tarayıcıda `/Admin/nobet_liste_sistemi/kurulum.asp` adresini açın
4. Yönetici girişi: `admin` / `admin123`

## Sayfalar

| Sayfa | Adres |
|-------|-------|
| Kurulum | `/Admin/nobet_liste_sistemi/kurulum.asp` |
| Liste | `/Admin/nobet_liste_sistemi/index.asp` |
| Yönetim | `/Admin/nobet_liste_sistemi/admin/login.asp` |

## Gereksinimler

- IIS + Classic ASP
- Microsoft Jet OLEDB 4.0 veya Access Database Engine (ACE)
- Dosya yükleme: Persits.Upload veya ABCUpload
