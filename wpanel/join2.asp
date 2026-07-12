<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>

<!--#INCLUDE file="menu.asp"-->

<CENTER>
<font class=fy>
<font color=#FF0000>
<BR><BR><BR>





<%
Function KarakterTemizle(Veri)
KarakterTemizle = Replace(Veri, "'", "''")
Veri = KarakterTemizle
End Function



stokk=10


sql1 = "INSERT INTO products(isim, tip, anagrup, surface, notlar,  delivery, yayin, kodu, stok, keyw, descr) Values('" & request("isim") & "','" & request("tip") & "','" & request("anagrup") & "', '" & request("surface") & "', '" & KarakterTemizle(Request.Form("notlar")) & "','" & request("delivery") & "', " & request("yayin") & ", '" & request("kodu") & "'," & stokk & ",'" & KarakterTemizle(Request.Form("keyw")) & "', '" & KarakterTemizle(Request.Form("descr")) & "')"

baglanti.Execute(sql1)

%>


 <font color=#0080FF size=3>
<BR><BR>
<B>Ürün Bilgisi Kaydedildi.
<BR><BR>

<%
set object3 = Server.CreateObJect("ADODB.RecordSet")
sql = "Select * from products order by AffiliateID desc"
object3.Open sql,baglanti,1,3
%>

 <a href="edits.asp?id=<%=object3("AffiliateID")%>&g=<%=request("anagrup")%>&tip=<%=request("tip")%>"><font color=#E10000>Görsel Eklemek İçin Tıklayın</B> </a>



	</CENTER>
