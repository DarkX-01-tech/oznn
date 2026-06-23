# Nöbet Liste Sistemi

Tek klasörde çalışan nöbet listesi modülü. Veritabanı: **Microsoft Access** (otomatik oluşur).

## Klasör yapısı

```
nobet_liste_sistemi/
├── index.asp              Liste sayfası (halka açık)
├── kurulum.asp            İlk kurulum
├── ayarlar.asp            Ayarlar
├── web.config             IIS yükleme limiti
├── admin/                 Yönetici paneli
├── assets/                CSS
├── database/              Access veritabanı (otomatik oluşur)
├── lib/                   Ortak kodlar
├── listeler/              PDF/XLS dosyaları
│   └── 2026/
│       ├── pendik/haziran/
│       └── basibuyuk/haziran/
└── tmp/                   Geçici yükleme klasörü
```

## Kurulum (3 adım)

1. Bu klasörü sunucuya kopyala: `C:\inetpub\wwwroot\Admin\nobet_liste_sistemi\`
2. IIS'e `database`, `listeler`, `tmp` klasörlerinde yazma izni ver
3. Tarayıcıda aç: `http://SUNUCU/Admin/nobet_liste_sistemi/kurulum.asp`

## Sayfalar

| Sayfa | Adres |
|-------|-------|
| Kurulum | `/Admin/nobet_liste_sistemi/kurulum.asp` |
| Liste | `/Admin/nobet_liste_sistemi/index.asp` |
| Yönetim | `/Admin/nobet_liste_sistemi/admin/login.asp` |

**Giriş:** admin / admin123

## Gereksinimler

- IIS + Classic ASP
- Microsoft Jet OLEDB 4.0 veya Access Database Engine
- Dosya yükleme: Persits.Upload veya ABCUpload
