<!-- #include file="../lib/database_setup.asp" -->
<%
Dim conn, connHata
connHata = ""

On Error Resume Next
Call EnsureAccessDatabase()
If Err.Number <> 0 Then
    connHata = "Veritabani hazirlanamadi: " & Err.Description
    Err.Clear
End If

If connHata = "" Then
    Call EnsureAccessTables()
    If Err.Number <> 0 Then
        connHata = "Veritabani tablolari hazirlanamadi: " & Err.Description
        Err.Clear
    End If
End If

If connHata = "" Then
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open AccessBaglantiMetni()
    If Err.Number <> 0 Then
        connHata = "Veritabani baglantisi kurulamadi: " & Err.Description
        Err.Clear
        Set conn = Nothing
    End If
End If

On Error GoTo 0

Function VeritabaniHazirMi()
    VeritabaniHazirMi = (connHata = "" And IsObject(conn))
End Function
%>
