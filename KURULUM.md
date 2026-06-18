# Duyuru Tablo Görünümü Düzeltmesi

Bu paket, `index.asp` ve `tumduyuru.asp` sayfalarında duyuru tablolarının her iki sayfada da aynı görünmesini ve tabloların gizleme butonunun arkasına düşmemesini sağlar.

## Dosyalar

| Dosya | Açıklama |
|-------|----------|
| `duyuru-tablo.css.inc` | Ortak tablo stilleri (her iki sayfaya include edilir) |
| `duyuru-icerik-render.inc` | Ortak içerik yazdırma mantığı — tablolar her zaman görünür |
| `index.asp` | Güncellenmiş anasayfa |
| `tumduyuru.asp` | Güncellenmiş tüm duyurular sayfası (DOCTYPE + UTF-8) |
| `js/duyuru-modal.js` | Görsel modal ve aç/kapa scriptleri |
| `js/duyuru-search.js` | Tüm duyurular arama scripti |
| `apply-duyuru-tablo-patch.py` | Mevcut dosyalara otomatik yama (opsiyonel) |

## Sunucuya Kurulum

1. `duyuru-tablo.css.inc` dosyasını `index.asp` ile aynı klasöre kopyalayın.
2. Mevcut `index.asp` ve `tumduyuru.asp` dosyalarınızı yedekleyin.
3. Bu repodaki güncel dosyaları sunucuya yükleyin.

## Manuel Uygulama (mevcut dosyaları korumak istiyorsanız)

### 1. Her iki dosyada `<style>` bloğuna ekleyin

`.duyuru-icerik strong` kuralından hemen sonra:

```html
        <!--#include file="duyuru-tablo.css.inc"-->
```

### 2. `tumduyuru.asp` dosya başına ekleyin

```html
<%@ Language="VBScript" %>
<!DOCTYPE html>
<html>
```

`charset=windows-1254` satırını şununla değiştirin:

```html
<meta charset="utf-8">
```

### 3. Her iki dosyada VBScript değişikliği

Resim **olmayan** içerikte, gizleme kontrolünü şu şekilde değiştirin:

```vbscript
' ESKİ:
If gorsellerKapaliMi Then

' YENİ:
If gorsellerKapaliMi And InStr(LCase(icerik), "<table") = 0 Then
```

Bu sayede tablo içeren duyurular her zaman doğrudan görünür.
