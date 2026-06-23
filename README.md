# MÜ Pendik E.A.H. — Dinamik Nöbet Listesi Sistemi

Bu proje, hastane portalındaki nöbet listeleri sayfasını ay bazlı dinamik dosya kontrolü ile çalışacak şekilde günceller.

## Nasıl çalışır?

1. Sayfa açıldığında sistem o anki ay için klasör adını belirler (ör. `haziran 2026`).
2. `Admin/nobet_liste_sistemi/` altında bu klasör yoksa otomatik oluşturulur.
3. Her liste satırı sabit bir dosya adına bağlıdır.
4. İlgili dosya ay klasöründe varsa link **aktif**, yoksa **pasif** görünür.
5. PDF/XLS dosyasını klasöre attığınızda sayfa yenilendiğinde link otomatik aktif olur.

## Kurulum

### 1. Dosyaları sunucuya kopyalayın

```
inetpub/wwwroot/
  tumnobetlistesi.asp
  lib/
    nobet_liste_lib.asp
    nobet_liste_config.asp
  Admin/
    nobet_liste_sistemi/    ← boş klasör (otomatik alt klasörler oluşur)
```

Mevcut `Connection.asp`, `ayarlar.asp` ve `sol.asp` dosyalarınız aynı kalır.

### 2. Veritabanı (isteğe bağlı)

Dosya durumunu izlemek için `sql/nobet_liste_dosyalar.sql` scriptini çalıştırın. Bağlantı yoksa sistem yine de dosya sistemi üzerinden çalışır.

### 3. IIS izinleri

`IIS_IUSRS` veya uygulama havuzu kimliğinin `Admin/nobet_liste_sistemi` klasöründe **okuma + yazma** izni olmalıdır (ay klasörü otomatik oluşması için).

## Klasör yapısı örneği

```
Admin/nobet_liste_sistemi/
  haziran 2026/
    pendik_acil_radyoloji_nobet_listesi.pdf
    pendik_idari_hekim_nobet_listesi.pdf
    klinik_icapci_hekim_nobet_listesi.xls
```

## Dosya adları

Tüm eşleşmeler `lib/nobet_liste_config.asp` içinde tanımlıdır. Yeni liste eklemek veya dosya adını değiştirmek için bu dosyayı düzenleyin.

## Pasif / aktif link

| Durum | Görünüm |
|-------|---------|
| Dosya yok | Gri metin, tıklanamaz |
| Dosya var | Turkuaz link, yeni sekmede açılır |
