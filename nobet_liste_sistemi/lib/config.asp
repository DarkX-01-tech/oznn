<%
Dim pendikNobetListeleri
pendikNobetListeleri = Array( _
    Array("Hekim Çalışma Listeleri", "hekim_calisma_listeleri.pdf"), _
    Array("İdari Hekim Nöbet Listesi", "pendik_idari_hekim_nobet_listesi.pdf"), _
    Array("Klinik İcapçı Hekim Nöbet Listesi", "klinik_icapci_hekim_nobet_listesi.xls"), _
    Array("Hemşire Süpervizör Nöbet Listesi", "pendik_hemsire_supervizor_nobet_listesi.pdf"), _
    Array("Destek Kalite Süpervizör Nöbet Listesi", "pendik_destek_kalite_supervizor_nobet_listesi.pdf"), _
    Array("Askom Nöbet Listesi", "pendik_askom_nobet_listesi.pdf"), _
    Array("Acil Radyoloji Nöbet Listesi", "pendik_acil_radyoloji_nobet_listesi.pdf"), _
    Array("Acil Laboratuvar Nöbet Listesi", "pendik_acil_laboratuvar_nobet_listesi.pdf"), _
    Array("Acil Transfer Ekip Nöbet Listesi", "pendik_acil_transfer_ekip_nobet_listesi.pdf"), _
    Array("Ambulans Şoför Nöbet Listesi", "pendik_ambulans_sofor_nobet_listesi.pdf"), _
    Array("Kan Bankası Nöbet Listesi", "pendik_kan_bankasi_nobet_listesi.pdf"), _
    Array("Bilgi İşlem Nöbet Listesi", "pendik_bilgi_islem_nobet_listesi.pdf"), _
    Array("Eczane Nöbet Listesi", "pendik_eczane_nobet_listesi.pdf"), _
    Array("Memur Nöbet Listesi", "pendik_memur_nobet_listesi.pdf"), _
    Array("Santral Nöbet Listesi", "pendik_santral_nobet_listesi.pdf") _
)

Dim basibuyukNobetListeleri
basibuyukNobetListeleri = Array( _
    Array("İdari Hekim Nöbet Listesi", "asaf_ataseven_idari_hekim_nobet_listesi.pdf"), _
    Array("Hemşire Süpervizör Nöbet Listesi", "asaf_ataseven_hemsire_supervizor_nobet_listesi.pdf"), _
    Array("Destek Kalite Süpervizör Nöbet Listesi", "asaf_ataseven_destek_kalite_supervizor_nobet_listesi.pdf"), _
    Array("Bilgi İşlem Nöbet Listesi", "asaf_ataseven_bilgi_islem_nobet_listesi.pdf"), _
    Array("Eczane Nöbet Listesi", "asaf_ataseven_eczane_nobet_listesi.pdf"), _
    Array("Memur Nöbet Listesi", "asaf_ataseven_memur_nobet_listesi.pdf") _
)

Function BinaAdiGoster(binaKodu)
    Select Case LCase(binaKodu)
        Case BINA_PENDIK
            BinaAdiGoster = "Pendik E.A.H."
        Case BINA_BASIBUYUK
            BinaAdiGoster = "Prof. Dr. Asaf Ataseven Ek Hizmet Binası"
        Case Else
            BinaAdiGoster = binaKodu
    End Select
End Function

Function ListeDizisiGetir(binaKodu)
    Select Case LCase(binaKodu)
        Case BINA_PENDIK
            ListeDizisiGetir = pendikNobetListeleri
        Case BINA_BASIBUYUK
            ListeDizisiGetir = basibuyukNobetListeleri
        Case Else
            ListeDizisiGetir = Array()
    End Select
End Function

Function DosyaAdiGecerliMi(binaKodu, dosyaAdi)
    Dim liste, i, satir
    DosyaAdiGecerliMi = False
    liste = ListeDizisiGetir(binaKodu)

    For i = 0 To UBound(liste)
        satir = liste(i)
        If LCase(satir(1)) = LCase(dosyaAdi) Then
            DosyaAdiGecerliMi = True
            Exit Function
        End If
    Next
End Function

Function BaslikGetir(binaKodu, dosyaAdi)
    Dim liste, i, satir
    BaslikGetir = dosyaAdi
    liste = ListeDizisiGetir(binaKodu)

    For i = 0 To UBound(liste)
        satir = liste(i)
        If LCase(satir(1)) = LCase(dosyaAdi) Then
            BaslikGetir = satir(0)
            Exit Function
        End If
    Next
End Function
%>
