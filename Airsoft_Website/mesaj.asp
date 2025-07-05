<%@ Language=VBScript %>
<% Option Explicit %>

<%
Dim isim, soyisim, email, mesaj
isim = Request.Form("Name")
soyisim = Request.Form("PhoneNumber")
email = Request.Form("Email")
mesaj = Request.Form("Message")

Dim mdb
mdb = Server.MapPath("veritabani.mdb")

Dim connStr
connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdb & ";"

Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open connStr

Dim sql
sql = "INSERT INTO Mesaj (isim, soyisim, email, mesaj) VALUES ('" & isim & "', '" & soyisim & "', '" & email & "', '" & mesaj & "')"

conn.Execute sql

conn.Close
Set conn = Nothing
%>

<!DOCTYPE html>
<html>
<head>
    <title>Form Gönderimi</title>
</head>
<body>
    <h3>Form basariyla gonderildi!</h3>
</body>
</html>

