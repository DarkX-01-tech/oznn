<%
' Development connection for AxonASP on Linux (MariaDB).
' Production deployments use Windows IIS + Access and keep this file out of git.
Set ConnYemek = Server.CreateObject("ADODB.Connection")
ConnYemek.ConnectionString = "Driver=mysql;Server=localhost;Database=oznn_yemek;Uid=oznn;Pwd=oznn_dev"
ConnYemek.Open
%>
