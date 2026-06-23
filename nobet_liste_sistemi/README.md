# Nöbet Liste Sistemi

Tek klasörde çalışan nöbet listesi modülü. Veritabanı: **Microsoft Access** (otomatik oluşur).

## Yönetim akışı

1. **panel.asp** — Yıl/ay seçimi, ana menü
2. **listeler.asp** — Seçilen dönem için yükleme/güncelleme/silme
3. **istatistikler.asp** — Yıl → ay → yükleme/güncelleme istatistikleri
4. **gecmis.asp** — Geçmiş dönem listelerini görüntüleme

## Klasör yapısı

```
nobet_liste_sistemi/
├── index.asp
├── kurulum.asp
├── admin/
│   ├── panel.asp
│   ├── listeler.asp
│   ├── istatistikler.asp
│   ├── gecmis.asp
│   ├── yukle.asp
│   └── login.asp
├── listeler/2026/pendik/haziran/
└── database/nobet_liste.mdb
```

## Özellikler

- Dönem bazlı liste yönetimi (yıl + ay seçimi)
- Dosya adı otomatik standartlaştırma (ör. her zaman `hekim_calisma_listeleri.pdf`)
- Yükleme/güncelleme sonrası panele dönüş
- İstatistik ve geçmiş dönem görüntüleme
- Halka açık sayfa her zaman güncel ayı gösterir

## Kurulum

1. Klasörü `C:\inetpub\wwwroot\Admin\nobet_liste_sistemi\` altına kopyalayın
2. IIS yazma izni verin (`database`, `listeler`, `tmp`)
3. `/Admin/nobet_liste_sistemi/kurulum.asp` adresini açın

**Giriş:** admin / admin123
