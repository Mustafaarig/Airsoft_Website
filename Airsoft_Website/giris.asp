<!DOCTYPE html>
<html lang="en">
<head>
    <style>
        body {
            background-color: black; 
            color: white; 
            font-family: 'Poppins', sans-serif;
        }

        .error-message {
            font-size: 20px;
            font-weight: bold;
            color: red;
            text-align: center;
            text-decoration: underline;
            margin-top: 20px;
            cursor: pointer; 
        }
    </style>
</head>
<body>
    <%
    Dim username, password
    username = Request.Form("username")
    password = Request.Form("password")
    Dim conn, rs
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("veritabani.mdb")
    Dim strSQL
    strSQL = "SELECT * FROM Kullanicilar WHERE Kadi = '" & username & "' AND sifre = '" & password & "'"
    Set rs = conn.Execute(strSQL)
    If Not rs.EOF Then
        Response.Redirect("anasayfa.html")
    Else
    %>
        <div class="error-message">Hatali kullanici adi veya sifre. Giris sayfasina geri don</div>
    <%
    End If
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    %>
</body>
</html>

