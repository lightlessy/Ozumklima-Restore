
  <!--#include file="menu.asp"-->     
  

<div class="baslik">Ürün Ekle </div>
  
  <body OnLoad="initialiseWebWizRTE();">        
  
  
     	<script src="../js/jscolor.js"></script>            

                                               
<% grup= request.querystring("g")
 tip= request.querystring("tip")   


 if  request.querystring("grup")="add"  then  


set Rs = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From anagrup"
Rs.open Sorgula,baglanti,1,3    

Rs.AddNew            
Rs("isim")=   Request.Form("isim1")    
Rs("durum")=   1
Rs.Update            

Response.Redirect Request.ServerVariables("HTTP_REFERER")        

end if %>       




 
 
 <% if  request.querystring("tip")="add"  then  


set Rs = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From tip"
Rs.open Sorgula,baglanti,1,3    

Rs.AddNew            
Rs("isim")=   Request.Form("isim2") 
Rs("grup")=   Request.Form("grup") 
Rs.Update            

Response.Redirect Request.ServerVariables("HTTP_REFERER")        

end if %>       
                                             


             
<% if  request.querystring("surf")="add"  then  


set Rs = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From surface"
Rs.open Sorgula,baglanti,1,3    

Rs.AddNew            
Rs("isim")=   Request.Form("isim3") 
Rs("grup")=   Request.Form("grubu") 
Rs.Update            

Response.Redirect Request.ServerVariables("HTTP_REFERER")        

end if %>     
                          

  <BR><BR>

                               
<FORM  name="f"  action=join2.asp method=post >                    


<Table  border=1  bordercolor=#DFDFDF bgcolor=#FFFFFF cellpadding=5 cellspacing=5 width=900 class=font align=center>    


 
<script>
    $(function(){
      // bind change event to select
      $('#anagrup').bind('change', function () {  
      
          var url = '?g='+ $(this).val(); // get selected value
      

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
      
          var url = '?g=<%=grup%>&tip='+ $(this).val(); // get selected value
      

   if (url) { // require a URL
              window.location = url; // redirect
          }
          return false;
      });
    });        

</script>  


<tr><td><B>Ürün Kategori: </B></td><td>

<select name="anagrup"  id="anagrup">
<option value1="" value="">Lütfen Seçiniz</option>   
<%  
           
set mods = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From anagrup order by sira desc"
mods.open Sorgula,baglanti,1,3      

Do while not mods.Eof      
%>   
  
<option value="<%=mods("id")%>" <%if Cint(grup) =mods("id") then %> selected <% end if %>><%=mods("isim")%></option>    
 

<%  
mods.MoveNExt
Loop
%>   
</select>   

</td><td colspan=2 width=250>
                                     
<INPUT   style="WIDTH: 150px"  name="isim1">       
<input type="submit" value="+ Ana Kategori Ekle" onclick="f.action='join.asp?grup=add'; return true; ">   

 &nbsp; | &nbsp; <a href="javascript:penc1()"><img src="images/del.png" width=20></a> 
 </td></tr> 
 
 
  <% if grup <> "" then %>
  
<tr><td><B>Alt Kategori:</B> </td><td>
<select name="tip" id="tip"> 
<option value1=""  value="">Lütfen Seçiniz</option>   
<%  
           
set mods = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From tip where grup = "& grup &"  order by sira desc"
mods.open Sorgula,baglanti,1,3      

Do while not mods.Eof      
%>   
  
<option  value="<%=mods("id")%>" <%if Cint(tip) =mods("id") then %> selected <% end if %>><%=mods("isim")%></option>    
 

<%  
mods.MoveNExt
Loop
%>   
</select>   
</div>
</td><td colspan=2>
                                     
<INPUT   style="WIDTH: 150px"  name="isim2">       
<INPUT type="hidden"  name="grup" value="<%=grup%>">  
<input type="submit" value="+ Alt Kategori Ekle" onclick="f.action='join.asp?tip=add'; return true; ">   

 &nbsp; | &nbsp; <a href="javascript:penc2()"><img src="images/del.png" width=20></a> 
 </td></tr> 


<% end if %>

                      


  <% if tip <> "" then %>    
  
  
  <tr><td><B>Ürün Grubu:</B> </td><td>
<select name="surface" id="surface"> 
<option value1=""  value="0">Lütfen Seçiniz</option>   
<%  
           
set mods = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From surface where grup = "& grup &"  order by isim desc"
mods.open Sorgula,baglanti,1,3      

Do while not mods.Eof      
%>   
  
<option  value="<%=mods("id")%>" <%if Cint(surface) =mods("id") then %> selected <% end if %>><%=mods("isim")%></option>    
 

<%  
mods.MoveNExt
Loop
%>   
</select>   
</div>
</td><td colspan=2>
                                     
<INPUT   style="WIDTH: 100px"  name="isim3">       
<INPUT type="hidden"  name="grubu" value="<%=grup%>">  
<input type="submit" value="+ Ekle" onclick="f.action='join.asp?surf=add'; return true; ">   

 &nbsp; | &nbsp; <a href="javascript:penc3()"><img src="images/del.png" width=20></a> 
 </td></tr>    
  

<tr><td><B>Ürün Kodu: </B></td><td><INPUT  style="WIDTH: 100px"  name="kodu" value="0" ></td>

<tr><td><B>Ürün Adý: </B></td><td><INPUT  style="WIDTH: 350px"  name="isim"></td>      </tr>  

<tr><td><B>Renk Seçenekleri:</B>  </td><td colspan=3> 
<input class="jscolor"  name="renk1" value="0" style="width:50px; -webkit-border-radius: 40px;-moz-border-radius: 40px;border-radius: 40px; border:1px solid #C0C0C0">   
<input class="jscolor"  name="renk2" value="0" style="width:50px; -webkit-border-radius: 40px;-moz-border-radius: 40px;border-radius: 40px; border:1px solid #C0C0C0">   
<input class="jscolor"  name="renk3" value="0" style="width:50px; -webkit-border-radius: 40px;-moz-border-radius: 40px;border-radius: 40px; border:1px solid #C0C0C0">   
<input class="jscolor"  name="renk4" value="0" style="width:50px; -webkit-border-radius: 40px;-moz-border-radius: 40px;border-radius: 40px; border:1px solid #C0C0C0">   
<input class="jscolor"  name="renk5" value="0" style="width:50px; -webkit-border-radius: 40px;-moz-border-radius: 40px;border-radius: 40px; border:1px solid #C0C0C0">   
<input class="jscolor"  name="renk6" value="0" style="width:50px; -webkit-border-radius: 40px;-moz-border-radius: 40px;border-radius: 40px; border:1px solid #C0C0C0">   
<input class="jscolor"  name="renk7" value="0" style="width:50px; -webkit-border-radius: 40px;-moz-border-radius: 40px;border-radius: 40px; border:1px solid #C0C0C0">   
<input class="jscolor"  name="renk8" value="0" style="width:50px; -webkit-border-radius: 40px;-moz-border-radius: 40px;border-radius: 40px; border:1px solid #C0C0C0">   
<input class="jscolor"  name="renk9" value="0" style="width:50px; -webkit-border-radius: 40px;-moz-border-radius: 40px;border-radius: 40px; border:1px solid #C0C0C0">   
<input class="jscolor"  name="renk10" value="0" style="width:50px; -webkit-border-radius: 40px;-moz-border-radius: 40px;border-radius: 40px; border:1px solid #C0C0C0">   
 </td></tr> 


<tr><td><B>Beden / Numara Seçenekleri:</B>  </td><td colspan=3> 
<input  name="beden1" value="0" style="width:40px;">   
<input  name="beden2" value="0" style="width:40px;">
<input  name="beden3" value="0" style="width:40px;">
<input  name="beden4" value="0" style="width:40px;">
<input  name="beden5" value="0" style="width:40px;">
<input  name="beden6" value="0" style="width:40px;">
<input  name="beden7" value="0" style="width:40px;">
<input  name="beden8" value="0" style="width:40px;">
<input  name="beden9" value="0" style="width:40px;">
<input  name="beden10" value="0" style="width:40px;">

 </td></tr>  
 
 
<tr bgcolor=#F7F4F0><td><B>Özet - Tanýmlama:</B></td><td colspan=3>
<TEXTAREA   style="WIDTH: 700px; HEIGHT: 50px"  name="descr"></TEXTAREA></td></tr>                                                           

	<%
strFormName = "f"
strTextAreaName = "notlar"
%> 

<tr bgcolor=#F7F4F0><td colspan=4>    
<B>Ürün Detaylarý:</B> <BR><BR>

<!--#include file="RTE_editor_inc.asp" --><BR>
<TEXTAREA   style="WIDTH: 870px; HEIGHT: 390px"  name="notlar" id="notlar"></TEXTAREA></td></tr> 
  
   
  <tr><td><B>Kargoya Veriliþ Süresi:</B>  </td><td> <INPUT  style="WIDTH: 100px"  name="delivery" value="3">    </td></tr>


<tr><td><B>Anahtar Kelimeler: </B> </td><td colspan=3>
<TEXTAREA   style="WIDTH: 700px; HEIGHT: 50px"  name="keyw"></TEXTAREA></td></tr>


<input type="hidden" name="yayin" value="0">  

<tr><td colspan=4 align=center>              
<input type="submit" value="ÜRÜN EKLE"> 
</td></tr></table>
 </FORM>
               
 <BR><BR>&nbsp;
       
 <% end if %>  
 
