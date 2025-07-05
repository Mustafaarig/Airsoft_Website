<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Kayıt Ol</title>
</head>
<body>

<%
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("veritabani.mdb")

Dim kullaniciAdi, sifre
kullaniciAdi = Request.Form("Kadi")
sifre = Request.Form("sifre")

Dim kontrolSQL, kontrolRS
kontrolSQL = "SELECT COUNT(*) AS sayi FROM Kullanicilar WHERE Kadi = '" & kullaniciAdi & "'"
Set kontrolRS = conn.Execute(kontrolSQL)
If kontrolRS("sayi") > 0 Then
    Response.Write "Bu kullanıcı adı zaten kullanılmaktadır. Lütfen başka bir kullanıcı adı seçin."
Else
    Dim ekleSQL
    ekleSQL = "INSERT INTO Kullanicilar (Kadi, sifre) VALUES ('" & kullaniciAdi & "', '" & sifre & "')"
    conn.Execute ekleSQL

    Response.Write "Kayıt başarıyla eklendi."
End If
conn.Close
Set conn = Nothing
%>

</body>
</html>