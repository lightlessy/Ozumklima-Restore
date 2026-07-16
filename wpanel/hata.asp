<HTML>
<HEAD> 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">     
       
<title>HATALI BİLGİ GİRİŞİ</title>
<script>


var alertmessage="Hatalı Kullanıcı Bilgileri --- Bilgilerinizi Doğru Girdiğinizden Emin Olunuz"



//Alert only once per browser session (0=no, 1=yes)
var once_per_session=0


function get_cookie(Name) {
  var search = Name + "="
  var returnvalue = "";
  if (document.cookie.length > 0) {
    offset = document.cookie.indexOf(search)
    if (offset != -1) { // if cookie exists
      offset += search.length
      // set index of beginning of value
      end = document.cookie.indexOf(";", offset);
      // set index of end of cookie value
      if (end == -1)
         end = document.cookie.length;
      returnvalue=unescape(document.cookie.substring(offset, end))
      }
   }
  return returnvalue;
}

function alertornot(){
if (get_cookie('alerted')==''){
loadalert()
document.cookie="alerted=yes"
}
}

function loadalert(){
alert(alertmessage)
}

if (once_per_session==0)
loadalert()
else
alertornot()

</script>        

</HEAD>
<meta http-equiv="refresh"
content="0; URL=index.asp">
</body>
</html>