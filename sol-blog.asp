<span class="solMenuBas">BLOG</span>
 
<% 
Function cevir(data)

    If Not IsNull(data) Then
        data = LCase(Trim(data & ""))

        data = Replace(data, "ç", "c")
        data = Replace(data, "ğ", "g")
        data = Replace(data, "ı", "i")
        data = Replace(data, "i", "i")
        data = Replace(data, "ö", "o")
        data = Replace(data, "ş", "s")
        data = Replace(data, "ü", "u")

        data = Replace(data, "Ç", "c")
        data = Replace(data, "Ğ", "g")
        data = Replace(data, "İ", "i")
        data = Replace(data, "I", "i")
        data = Replace(data, "Ö", "o")
        data = Replace(data, "Ş", "s")
        data = Replace(data, "Ü", "u")

        data = Replace(data, "'", "")
        data = Replace(data, """", "")
        data = Replace(data, ".", "")
        data = Replace(data, ",", "")
        data = Replace(data, ":", "")
        data = Replace(data, ";", "")
        data = Replace(data, "/", "-")
        data = Replace(data, "\", "-")
        data = Replace(data, " ", "-")
    Else
        data = ""
    End If

    cevir = data

End Function 
%> 

<%
Dim bloglar, Sorgula

Set bloglar = Server.CreateObject("ADODB.RecordSet")

Sorgula = "Select * From sayfalar where ana=3 order by sira asc"

bloglar.Open Sorgula, baglanti, 1, 1

If bloglar.EOF Then
%>

    <div>Blog yaz&#305;s&#305; bulunamad&#305;.</div>

<%
Else

    Do While Not bloglar.EOF
%>   

    <a href="blog.asp?k=<%=Server.URLEncode(cevir(bloglar("isim"))) %>&id=<%=bloglar("id") %>">
        <div>
            <i class="fas fa-chevron-right"></i> <%=bloglar("isim") %>
        </div>
    </a>   
 
<%  
        bloglar.MoveNext
    Loop

End If

If IsObject(bloglar) Then
    If bloglar.State = 1 Then bloglar.Close
    Set bloglar = Nothing
End If
%>