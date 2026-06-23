<%
' Veritabanı bağlantısı - kurulumda bilgileri güncelleyin.
Dim conn

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLOLEDB;Data Source=SUNUCU_ADI;Initial Catalog=VERITABANI_ADI;User ID=KULLANICI;Password=SIFRE;"
%>
