<!--#INCLUDE file="menu.asp"-->        

<div class="baslik">Ürün Düzenle </div>

  <body OnLoad="initialiseWebWizRTE();">  

<% id=request.querystring("id") 
 tip=request.querystring("tip")
grup=request.querystring("g")
rela=request.querystring("rela")
rel1=request.querystring("rel1")
rel2=request.querystring("rel2")
toplam = Cint(Request.Form("toplam"))  


    	
    	Function cevir(data)   
if not data="" then
data = Replace (data ,"ı","i",1,-1,1) 
data = Replace (data ,"ğ","g",1,-1,1) 
data = Replace (data ,"ü","u",1,-1,1) 
data = Replace (data ,"ş","s",1,-1,1) 
data = Replace (data ,"ö","o",1,-1,1) 
data = Replace (data ,"ç","c",1,-1,1) 
data = Replace (data ,"I","i",1,-1,1) 
data = Replace (data ,"Ğ","g",1,-1,1) 
data = Replace (data ,"Ü","u",1,-1,1) 
data = Replace (data ,"Ş","s",1,-1,1) 
data = Replace (data ,"İ","i",1,-1,1) 
data = Replace (data ,"Ç","c",1,-1,1) 
data = Replace (data ,"Ö","o",1,-1,1) 
data = Replace (data ," ","-",1,-1,1)   
end if
cevir=Lcase(data)
End Function     %>
  
            
  

        	<script src="../js/jscolor.js"></script>           
        	


 <% if  request.querystring("pro")="add"  then  
                                                                   
for i = 1 to toplam  
if Request.Form("Veri"&i) <> "" then 
set ilg = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From related where urun ="& Request.Form("urun") &" and ilgili ="& Request.Form("Veri"&i) &"  "
ilg.open Sorgula,baglanti,1,3      
 

 if ilg.eof then
set Rs = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From related"
Rs.open Sorgula,baglanti,1,3    


Rs.AddNew         
Rs("urun")=   Request.Form("urun") 
Rs("ilgili")=    Request.Form("Veri"&i)      
Rs.Update     


Rs.AddNew         
Rs("urun")=   Request.Form("Veri"&i)   
Rs("ilgili")=    Request.Form("urun") 
Rs.Update       
end if 
end if     
Next 
 
 Response.Redirect Request.ServerVariables("HTTP_REFERER")  

 end if


 if  request.querystring("pro")="del"  then  
                                                                   
  
set ilg = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From related where urun ="& Request("rel1") &" and ilgili ="& Request("rel2") &"  "    
ilg.open Sorgula,baglanti,1,3      

ilg.Delete 
ilg.Update    
  

set ilg2 = Server.CreateObJect("ADODB.RecordSet")
Sorgula2 = "Select * From related where  ilgili ="& Request("rel1") &" and urun ="& Request("rel2") &" "
ilg2.open Sorgula2,baglanti,1,3

ilg2.Delete 
ilg2.Update     

      

Response.Redirect Request.ServerVariables("HTTP_REFERER")  

 end if %>         
          
        
       

<BR><BR>         



<%if request("submit")<>"" then          

fiyatg= request ("fiyatg")      
if fiyatg = "" then
fiyatg =0
end if       

yayin= request ("yayin")      
if yayin = "" then
yayin =0
end if         
              

yeni= request ("yeni")      
if yeni = "" then
yeni =0
end if     

indi= request ("indirim")      
if indi = "" then
indi =0
end if    



listele=request ("listele")      
if listele = "" then
listele =0
end if   

stok= request ("stok")      
if stok = "" then
stok =0
end if           


Function KarakterTemizle(Veri)
KarakterTemizle = Replace(Veri, "'", "''")
Veri = KarakterTemizle                                                         
End Function            


Function KarakterTemizle1(Veri1)
KarakterTemizle1 = Replace(Veri1, "," , ".")
Veri1 = KarakterTemizle1                                                         
End Function       

kodu = request("kodu")
if  kodu ="" then
kodu =0
end if


SQL = "Update products Set  isim = '"& request ("isim") &"' , anagrup='"& request ("anagrup") &"', tip='"& request ("tipi") &"', surface = '"& request ("surface") &"',  renk1 = '"& request ("renk1") &"',  renk2 = '"& request ("renk2") &"',  renk3 = '"& request ("renk3") &"',  renk4 = '"& request ("renk4") &"',   renk5 = '"& request ("renk5") &"', delivery = '"& request ("delivery") &"',   sira = '"& request ("sira") &"' , anas = '"& request ("anas") &"' , notlar = '"& KarakterTemizle(request ("notlar")) &"' , keyw = '"& KarakterTemizle(request ("keyw")) &"' ,  descr = '"& KarakterTemizle(request ("descr")) &"' , kodu = '"& kodu &"', listele = '"& listele &"' ,  indirim = "& indi &" , fiyat = "& request ("fiyat") &" ,  video = '"& request ("video") &"' , yeni = "& yeni &"  , yayin = '"& yayin &"' , stok = '"& stok &"'  Where AffiliateID = " & request("AffiliateID") 
Set RS = baglanti.Execute(SQL)    
Response.Redirect Request.ServerVariables("HTTP_REFERER")


else %>          
                 

<%

set rs = Server.CreateObject("ADODB.recordset")
sql="SELECT * FROM products where AffiliateID=" & request("id") &  " "
rs.Open sql, baglanti
%>
              



<%       

set object3 = Server.CreateObJect("ADODB.RecordSet")
sql = "Select * from products where AffiliateID ="& id
object3.Open sql,baglanti,1,3

%>                                           
  
  
                          
  
  
<FORM  name="f"  action=edits.asp method=post onSubmit="return CheckPollForm();" >          


<Table  border=1  bordercolor=#E6E6E6 cellpadding=6 cellspacing=6 width=1000 class=font align=center bgcolor=#FFFFFF>



<tr><td colspan=2> 
       
<B>Ürün Görsel Düzenleme: </B><BR>      

<table class=font border=1 cellspacing=0 cellpadding=0 bordercolor="#E9E9E9" width=650 align=center><tr>     


<td align=center>
<% if object3("foto1") <> "" then %><img src="../urunler/<%=object3("foto1")%>" width=100 height=100><BR>
<a href="imageupdate.asp?id=<%=object3("AffiliateID")%>&g=<%=object3("anagrup")%>&tip=<%=object3("tip")%>&isim=<%=cevir(object3("isim"))%>&n=1"><img src="images/edit.png" width=20> </a>
<a href="fotosil.asp?id=<%=object3("AffiliateID")%>&foto1=&foto2=<%=object3("foto2")%>&foto3=<%=object3("foto3")%>&foto4=<%=object3("foto4")%>&foto5=<%=object3("foto5")%>&sil=<%=object3("foto1")%>"><img src="images/del.png" ></a>
 <% else %> 
 <a href="imageupdate.asp?id=<%=object3("AffiliateID")%>&g=<%=object3("anagrup")%>&tip=<%=object3("tip")%>&isim=<%=cevir(object3("isim"))%>&n=1" title="Görsel Ekle"><img src="images/add.png"> </a> <% end if %>
</td>


<td align=center>
<% if object3("foto2") <> "" then %><img src="../urunler/<%=object3("foto2")%>" width=100 height=100><BR>
<a href="imageupdate.asp?id=<%=object3("AffiliateID")%>&g=<%=object3("anagrup")%>&tip=<%=object3("tip")%>&isim=<%=cevir(object3("isim"))%>&n=2"><img src="images/edit.png" width=20>  </a>  
<a href="fotosil.asp?id=<%=object3("AffiliateID")%>&foto1=<%=object3("foto1")%>&foto2=&foto3=<%=object3("foto3")%>&foto4=<%=object3("foto4")%>&foto5=<%=object3("foto5")%>&sil=<%=object3("foto2")%>"><img src="images/del.png" ></a>
<% else %> 
 <a href="imageupdate.asp?id=<%=object3("AffiliateID")%>&g=<%=object3("anagrup")%>&tip=<%=object3("tip")%>&isim=<%=cevir(object3("isim"))%>&n=2" title="Görsel Ekle"><img src="images/add.png"> </a> <% end if %>
</td>
          


<td align=center>   
<% if object3("foto3") <> "" then %><img src="../urunler/<%=object3("foto3")%>" width=100 height=100><BR>
<a href="imageupdate.asp?id=<%=object3("AffiliateID")%>&g=<%=object3("anagrup")%>&tip=<%=object3("tip")%>&isim=<%=cevir(object3("isim"))%>&n=3"><img src="images/edit.png" width=20>  </a> 
<a href="fotosil.asp?id=<%=object3("AffiliateID")%>&foto1=<%=object3("foto1")%>&foto2=<%=object3("foto2")%>&foto3=&foto4=<%=object3("foto4")%>&foto5=<%=object3("foto5")%>&sil=<%=object3("foto3")%>"><img src="images/del.png" ></a>
<% else %> 
 <a href="imageupdate.asp?id=<%=object3("AffiliateID")%>&g=<%=object3("anagrup")%>&tip=<%=object3("tip")%>&isim=<%=cevir(object3("isim"))%>&n=3" title="Görsel Ekle"><img src="images/add.png"> </a> <% end if %>
</td>


<td align=center>
<% if object3("foto4") <> "" then %><img src="../urunler/<%=object3("foto4")%>" width=100 height=100><BR>
<a href="imageupdate.asp?id=<%=object3("AffiliateID")%>&g=<%=object3("anagrup")%>&tip=<%=object3("tip")%>&isim=<%=cevir(object3("isim"))%>&n=4"><img src="images/edit.png" width=20>  </a>  
<a href="fotosil.asp?id=<%=object3("AffiliateID")%>&foto1=<%=object3("foto1")%>&foto2=<%=object3("foto2")%>&foto3=<%=object3("foto3")%>&foto4=&foto5=<%=object3("foto5")%>&sil=<%=object3("foto4")%>"><img src="images/del.png" ></a>
<% else %> 
 <a href="imageupdate.asp?id=<%=object3("AffiliateID")%>&g=<%=object3("anagrup")%>&tip=<%=object3("tip")%>&isim=<%=cevir(object3("isim"))%>&n=4" title="Görsel Ekle"><img src="images/add.png"> </a> <% end if %>
</td>


<td align=center>
<% if object3("foto5") <> "" then %><img src="../urunler/<%=object3("foto5")%>" width=100 height=100><BR>
<a href="imageupdate.asp?id=<%=object3("AffiliateID")%>&g=<%=object3("anagrup")%>&tip=<%=object3("tip")%>&isim=<%=cevir(object3("isim"))%>&n=5"><img src="images/edit.png" width=20>  </a> 
<a href="fotosil.asp?id=<%=object3("AffiliateID")%>&foto1=<%=object3("foto1")%>&foto2=<%=object3("foto2")%>&foto3=<%=object3("foto3")%>&foto4=<%=object3("foto4")%>&foto5=&sil=<%=object3("foto5")%>"><img src="images/del.png" ></a>
<% else %> 
 <a href="imageupdate.asp?id=<%=object3("AffiliateID")%>&g=<%=object3("anagrup")%>&tip=<%=object3("tip")%>&isim=<%=cevir(object3("isim"))%>&n=5" title="Görsel Ekle"><img src="images/add.png"></a> <% end if %>
</td></tr></table>


</td></tr>
<tr>             

<tr><td><B>Ürün Kodu: </B></td><td>       <INPUT  style="WIDTH: 100px"  name="kodu"  value="<%=object3("kodu")%>"> </td>  </tr>


 
<script>
    $(function(){
      // bind change event to select
      $('#anagrup').bind('change', function () {  
      
          var url = '?id=<%=id%>&tip=<%=tip%>&g='+ $(this).val(); // get selected value
      

   if (url) { // require a URL
              window.location = url; // redirect
          }
          return false;
      });
    });        

</script>    
       

<script>
    $(function(){
      // bind change event to select
      $('#tip').bind('change', function () {  
      
          var url = '?g=<%=grup%>&tip'+ $(this).val(); // get selected value
      

   if (url) { // require a URL
              window.location = url; // redirect
          }
          return false;
      });
    });        

</script>  



<tr bgcolor=#F7F4F0>
<td><B>Ürün Kategori: </B></td>
<td>    

<select name="anagrup"  id="anagrup">  
<%  
           
set mods = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From anagrup order by isim desc"
mods.open Sorgula,baglanti,1,3      

Do while not mods.Eof      
%>   
  
<option  value="<%=mods("id") %>" <% if grup = "" then %><%if object3("anagrup") =mods("id") then %> selected <% end if %> <% else %><%if Cint(grup) =mods("id") then %> selected <% end if %> <% end if %>><%=mods("isim")%></option>    
 

<%  
mods.MoveNExt
Loop
%>   
</select>   

</td>  
</tr>



<tr bgcolor=#F7F4F0>
<td><B>Alt Kategori:</B></td>
<td>

<select name="tipi"  id="tipi">  
<%  
           
set mods = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From tip where grup = "& grup &"  order by isim desc"
mods.open Sorgula,baglanti,1,3      

Do while not mods.Eof      
%>   
  
<option  value="<%=mods("id")%>" <% if tip = "" then %><%if Cint(object3("tip")) =mods("id") then %> selected <% end if %> <% else %><%if Cint(tip) =mods("id") then %> selected <% end if %> <% end if %>><%=mods("isim")%></option>    
 

<%  
mods.MoveNExt
Loop
%>   
</select> 

</td>  
</tr>      




<tr bgcolor=#F7F4F0>
<td><B>Ürün Grubu:</B></td>
<td>

<select name="surface"  id="surface">  
<%  
           
set mods = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From surface where grup = "& tip &"  order by isim desc"
mods.open Sorgula,baglanti,1,3      

Do while not mods.Eof      
%>   
  
<option  value="<%=mods("id")%>" <% if tip = "" then %><%if Cint(object3("tip")) =mods("id") then %> selected <% end if %> <% else %><%if Cint(tip) =mods("id") then %> selected <% end if %> <% end if %>><%=mods("isim")%></option>    
 

<%  
mods.MoveNExt
Loop
%>   

<% if object3("surface")=0 then %> <option value="0" selected>Yok</option>    <% end if %>
</select> 

</td>  
</tr>

  
<tr><td><B>İsim:</B> </td><td>  <TEXTAREA   style="padding:5px; WIDTH: 400px; HEIGHT: 50px"  name="isim"><%=object3("isim")%></TEXTAREA></td>  </tr>
 
 
 <tr><td><B>Renkler:</B> </td><td>  

<input class="jscolor"  name="renk1" value="<%=object3("renk1")%>" style="width:50px; -webkit-border-radius: 40px;-moz-border-radius: 40px;border-radius: 40px; border:1px solid #C0C0C0">   
<input class="jscolor"  name="renk2" value="<%=object3("renk2")%>" style="width:50px; -webkit-border-radius: 40px;-moz-border-radius: 40px;border-radius: 40px; border:1px solid #C0C0C0">   
<input class="jscolor"  name="renk3" value="<%=object3("renk3")%>" style="width:50px; -webkit-border-radius: 40px;-moz-border-radius: 40px;border-radius: 40px; border:1px solid #C0C0C0">   
<input class="jscolor"  name="renk4" value="<%=object3("renk4")%>" style="width:50px; -webkit-border-radius: 40px;-moz-border-radius: 40px;border-radius: 40px; border:1px solid #C0C0C0">   
<input class="jscolor"  name="renk5" value="<%=object3("renk5")%>" style="width:50px; -webkit-border-radius: 40px;-moz-border-radius: 40px;border-radius: 40px; border:1px solid #C0C0C0">   </td></tr>

       
<tr><td><B>Fiyat:</B> </td><td>   <INPUT  type="text"  name="fiyat"  value="<%=object3("fiyat")%>" size=10> TL</td></tr>

<tr><td><B>İndirimli Fiyat:</B> </td><td>   <INPUT  type="text"  name="indirim"  value="<%=object3("indirim")%>" size=10> TL </td></tr>

<tr><td><B>Yeni Ürün:</B> </td><td><input type="checkbox" name="yeni"  value="1" <% if object3("yeni")="1" then %> checked <% end if %>></td></tr>     
  
<tr><td><B>Stok Adet:</B> </td><td><INPUT name="stok" size="5" value="<%=object3("stok")%>"> </td></tr>    


<INPUT  type="hidden"  name="listele"  value="1">   
<INPUT  type="hidden"  name="anas"  value="<%=object3("anas")%>"> 
  <INPUT  type="hidden"   name="sira"  value="<%=object3("sira")%>">


<tr><td><B>Kargoya Veriliş Süresi :</B>  </td><td><INPUT  style="WIDTH: 100px"  name="delivery"  value="<%=object3("delivery")%>">  İçerisinde Kargoda!</td></tr>


<tr><td><B>Yayın Durumu:</B> </td><td><input type="checkbox" name="yayin"  value="1" <% if object3("yayin")="1" then %> checked <% end if %>></td></tr>   


<tr><td><B>Özet - Tanımlama: </B></td><td><TEXTAREA   style="WIDTH: 550px; HEIGHT: 90px"  name="descr"><%=object3("descr")%></TEXTAREA></td></tr>


	<%
strFormName = "f"
strTextAreaName = "notlar"
%> 
<tr><td colspan=2>   
<B>Ürün Detayları:</B> <BR><BR>
<!--#include file="RTE_editor_inc.asp" --><BR>
<TEXTAREA   style="padding:10px; WIDTH: 1000px; HEIGHT: 450px"  name="notlar" id="notlar"><%=object3("notlar")%></TEXTAREA></td></tr>      
      
                                                                                                                                                                 

<tr><td><B>Döküman Yükle</B></td><td>
<% if   object3.fields("dokuman")&"" = "" then  %>     
  <a href="dokuman.asp?id=<%=object3("AffiliateID")%>&g=<%=grup%>&tip=<%=tip%>"  style="color:#EC0000"><B>YÜKLE (+)</B> </a> 
  <% else %>
  <a href="../urunler/<%=object3("dokuman")%>"  style="color:#0070DF"  target=_blank><B>Döküman Yüklü - İncele </B> </a> 
  &nbsp;&nbsp;&nbsp;&nbsp;
  
  <a href="dokumansil.asp?id=<%=object3("AffiliateID")%>"  onclick="return confirm('Kaydı Silmek İstiyor Musunuz?')"> <B>DÖKÜMANI SİL (-)</B></a>
  
  <% end if %>  
  </td></tr>      
  
  
  
  <tr><td><B>Teknik Özellik Yükle</B></td><td>
<% if   object3.fields("uretici")&"" = "" then  %>     
  <a href="teknik.asp?id=<%=object3("AffiliateID")%>&g=<%=grup%>&tip=<%=tip%>"  style="color:#EC0000"><B>YÜKLE (+)</B> </a> 
  <% else %>
  <a href="../urunler/<%=object3("uretici")%>"  style="color:#0070DF"  target=_blank><B>Teknik Özellik Yüklü - İncele </B> </a> 
  &nbsp;&nbsp;&nbsp;&nbsp;
  
  <a href="tekniksil.asp?id=<%=object3("AffiliateID")%>"  onclick="return confirm('Kaydı Silmek İstiyor Musunuz?')"> <B>TEKNİK ÖZELLİK SİL (-)</B></a>
  
  <% end if %>  
  </td></tr>      
  

<tr><td><B>Video Embed Kodu </B></td><td><TEXTAREA   style="WIDTH: 550px; HEIGHT: 90px"  name="video"><%=object3("video")%></TEXTAREA></td></tr>   
 
<tr><td><B>Anahtar Kelimeler: [SEO]</B> <BR> Önerilen kelime sayısı: 4-6 kelime</td><td><TEXTAREA   style="WIDTH: 550px; HEIGHT: 90px"  name="keyw"><%=object3("keyw")%></TEXTAREA></td></tr>



<%
set mevc = Server.CreateObJect("ADODB.RecordSet")
Sor = "Select * From related where urun = "& id &" "
mevc.open Sor,baglanti,1,3             
 
Do while not mevc.Eof      

set urunler = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From products where AffiliateID = "& mevc("ilgili") &" "
urunler.open Sorgula,baglanti,1,3 


set tipler = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From tip where id = "& urunler("tip") &" "
tipler.open Sorgula,baglanti,1,3 

%>
   

 
<%  
mevc.MoveNExt
Loop
%>    



<script>
    $(function(){
      // bind change event to select
      $('#tip').bind('change', function () {  
      
          var url = '?id=<%=id%>&g=<%=grup%>&tip=<%=tip%>&rela='+ $(this).val()+'#rel'; // get selected value
      

   if (url) { // require a URL
              window.location = url; // redirect
          }
          return false;
      });
    });        

</script>    


<tr><td><a name="rel"></a><B>Benzer Ürünler:</B></td>

<td>

<select name="tip" id="tip" style="font-size:15px; ">  
<option value1="" >Lütfen Seçiniz</option>   
<%  
           
set mods = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select id, isim,grup From tip group by grup, isim, id"
mods.open Sorgula,baglanti,1,3      

Do while not mods.Eof      
%>   
  
<option   value="<%=mods("id")%>" <% if rela = "" then %><%if Cint(rela) =mods("id") then %> selected <% end if %> <% else %><%if Cint(rela) =mods("id") then %> selected <% end if %> <% end if %>><%=mods("isim")%> </option>    
 

<%  
mods.MoveNExt
Loop
%>   
</select>   

   
<BR><BR>

<% 

set pros = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From products where tip = "& Cint(rela) &" "
pros.open Sorgula,baglanti,1,3
   
i=1
Do while not pros.Eof 


set ilg = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From related where urun= "& id &" and ilgili = "& pros("AffiliateID") &" "
ilg.open Sorgula,baglanti,1,3     

  %>
              

<% if ilg.eof then %> 
<input type="checkbox" name="Veri<%=i%>" value="<%=pros("AffiliateID")%>"><%=pros("isim")%> <BR>           
<% else %>             

<input type="checkbox" name="Veri<%=i%>" value="<%=pros("AffiliateID")%>" <% if ilg.eof then %> <% else %><% if  ilg("ilgili") =  pros("AffiliateID")  then%>checked><a href="edits.asp?id=<%=id%>&g=<%=grup%>&tip=<%=tip%>&rela=<%=rela%>&pro=del&rel1=<%=ilg("urun")%>&rel2=<%=ilg("ilgili")%>#rel">Kaldır </a> <% end if %><% end if %>><%=pros("isim")%>  <BR> 
<% end if %>    



<% 
i=i+1
pros.MoveNExt     
Loop
%>                      

<input type="hidden" name="urun"  value="<%=id%>"> 

<input type="hidden" name="toplam" value="<%=Int(i-1)%>">
<BR> 
<input type="submit" value="Benzer Ürünleri Belirle"   onclick="f.action='edits.asp?pro=add#rel'; return true; " style="font-size:12px; background-color:#A4A4A4">   

</td></tr>
<tr>           
<input type="hidden" name="AffiliateID" value="<%=object3("AffiliateID")%>">       

<tr><td colspan=2 align=center>   
<input type="submit" value="GÜNCELLE" name="submit"> 
</td></tr></table>
 </FORM>


 
 <%end if %>
         
 <BR><BR>
 
</body>
</html>
