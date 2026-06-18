#!/usr/bin/env python3
"""Mevcut index.asp ve tumduyuru.asp dosyalarına tablo düzeltmesini uygular."""

from pathlib import Path

ROOT = Path(__file__).resolve().parent

INDEX_PATCHES = [
    (
        ".duyuru-icerik strong {\n          font-weight: 800 !important;\n        }",
        ".duyuru-icerik strong {\n          font-weight: 800 !important;\n        }\n\n        <!--#include file=\"duyuru-tablo.css.inc\"-->",
    ),
    (
        '<!-- #include file="ayarlar.asp" -->',
        '<!-- #include file="ayarlar.asp" -->\n    <!-- #include file="duyuru-icerik-render.inc" -->',
    ),
    (
        "If gorsellerKapaliMi Then\n                imageCounter = imageCounter + 1\n                linkId2 = \"imgLink\"",
        "If gorsellerKapaliMi And InStr(LCase(icerik), \"<table\") = 0 Then\n                imageCounter = imageCounter + 1\n                linkId2 = \"imgLink\"",
    ),
]

TUM_PATCHES = [
    (
        '<%\nsession("ok") = false\n%>',
        '<%@ Language="VBScript" %>\n<!DOCTYPE html>\n<html>\n<%\nsession("ok") = false\n%>',
    ),
    (
        '<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">',
        '<meta charset="utf-8">',
    ),
    (
        ".duyuru-icerik strong  { font-weight: 800 !important; }",
        ".duyuru-icerik strong  { font-weight: 800 !important; }\n\n    <!--#include file=\"duyuru-tablo.css.inc\"-->",
    ),
    (
        '<!-- #include file="ayarlar.asp" -->',
        '<!-- #include file="ayarlar.asp" -->\n<!-- #include file="duyuru-icerik-render.inc" -->',
    ),
    (
        "If gorsellerKapaliMi Then\n            Dim linkId2, contentId2\n            imageCounter = imageCounter + 1",
        "If gorsellerKapaliMi And InStr(LCase(icerik), \"<table\") = 0 Then\n            Dim linkId2, contentId2\n            imageCounter = imageCounter + 1",
    ),
]


def apply_patches(path: Path, patches: list[tuple[str, str]]) -> bool:
    if not path.exists():
        print(f"ATLANDI (dosya yok): {path}")
        return False

    content = path.read_text(encoding="utf-8", errors="replace")
    original = content

    for old, new in patches:
        if old not in content:
            print(f"UYARI: Desen bulunamadı ({path.name}): {old[:60]}...")
        else:
            content = content.replace(old, new, 1)

    if content != original:
        path.write_text(content, encoding="utf-8")
        print(f"GÜNCELLENDİ: {path}")
        return True

    print(f"DEĞİŞİKLİK YOK: {path}")
    return False


def main() -> None:
    apply_patches(ROOT / "index.asp", INDEX_PATCHES)
    apply_patches(ROOT / "tumduyuru.asp", TUM_PATCHES)
    print("Bitti.")


if __name__ == "__main__":
    main()
