<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
Response.ContentType = "text/html"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>  
 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"> 
   <meta name="viewport" content="width=device-width, initial-scale=1">

<title>Antalya Klima Servisi ve Mitsubishi Electric Klima &Ccedil;&ouml;z&uuml;mleri | &Ouml;z&uuml;m Klima</title>

<meta name="author" content="MEGA TASARIM">    

  <meta name="description" content="&Ouml;z&uuml;m Klima, Antalya&#39;da Mitsubishi Electric klima, VRF ve havaland&#305;rma sistemlerinde ke&#351;if, montaj, bak&#305;m ve teknik servis &#231;&ouml;z&uuml;mleri sunar.">
  <link rel="canonical" href="https://ozumklima.com/" />


<meta name="author" content="Mega Tasar&#305;m">   

    <link rel="icon" type="image/png" href="fav.png" />   

<link href="css/font.css" rel="stylesheet" type="text/css"/>   

<script src="js/jquery-2.1.1.js"></script>  
<script src="js/jquery-ui.min.js"></script>               

   <link href="css/fontawesome-all.css" rel="stylesheet">
                                                                                 
   
           
   
      <link href="css/textslide.css" rel="stylesheet">
   
   
<link rel="stylesheet" href="css/owl.carousel.min.css">
<link rel="stylesheet" href="css/owl.theme.default.min.css">
<script src="js/owl.carousel.min.js"></script>



<link rel="stylesheet" href="css/flexslider.css" type="text/css">
<script src="js/jquery.flexslider.js"></script> 

 <style>
  .flexslider .slides > li {
    background-size: cover; 
    background-position: center ;
    height: 600px; 
    max-width: 1270px;
    display: none; 
    -webkit-backface-visibility:   hidden;  
  
}   


.flexslider-container, .flexslider .slides, .flex-viewport {
    height: 600px; 
    border: 0px solid #FF0000;
}
html, body {
    height: 100%;

    
}
body {
    position: relative;

    margin: 0;
    padding: 0;
}

.flexslider {margin: 0; padding: 0; margin:0px auto; max-width: 1270px; height: 600px; border: 0px solid; overflow: hidden; }
.flexslider .slides img {width: auto; height: 100%; display: block;}
.flex-pauseplay span {text-transform: capitalize;}     
.hizmetGovdesi {width:92%; max-width:1120px; margin:35px auto 45px auto; padding:35px 4%; background:#FFFFFF; border-top:1px solid #E9E9E9; border-bottom:1px solid #E9E9E9; color:#222222; font-family:Arial, Helvetica, sans-serif; box-sizing:border-box;}
.hizmetGovdesi h1 {margin:0 0 12px 0; padding:0; font-size:1.7em; line-height:1.35em; font-weight:700; color:#171717; text-align:center;}
.hizmetGovdesi .hizmetAltBaslik {display:block; margin:0 auto 24px auto; max-width:760px; color:#777777; font-size:1em; line-height:1.65em; text-align:center;}
.hizmetGovdesi .hizmetMetin {max-width:900px; margin:0 auto; column-count:2; column-gap:42px;}
.hizmetGovdesi p {margin:0 0 18px 0; padding:0; font-size:1em; line-height:1.75em; color:#333333;}
.hizmetGovdesi a {color:#DE0303; text-decoration:none; font-weight:700;}
.hizmetGovdesi a:hover {text-decoration:underline;}
.hizmetGovdesi .hizmetAksiyon {max-width:900px; margin:12px auto 0 auto; padding-top:18px; border-top:1px solid #EFEFEF; text-align:center; color:#555555; font-size:1em; line-height:1.65em;}
@media only screen and (max-width: 760px) {
.hizmetGovdesi {width:92%; margin:24px auto 32px auto; padding:26px 5%;}
.hizmetGovdesi h1 {font-size:1.35em;}
.hizmetGovdesi .hizmetMetin {column-count:1;}
}

/* =========================
   MODERN MÜŞTERİ YORUMLARI
   ========================= */
.musteriYorumlari {
  width: 92%;
  max-width: 1120px;
  margin: 38px auto 50px auto;
  padding: 48px 4%;
  box-sizing: border-box;
  background: #F8F8F8;
  border-top: 1px solid #E9E9E9;
  border-bottom: 1px solid #E9E9E9;
  font-family: Arial, Helvetica, sans-serif;
}

.musteriYorumlariBaslik {
  max-width: 720px;
  margin: 0 auto 28px auto;
  text-align: center;
}

.musteriYorumlariEtiket {
  display: inline-block;
  margin-bottom: 10px;
  color: #DE0303;
  font-size: 0.78em;
  line-height: 1;
  font-weight: 700;
  letter-spacing: 0.12em;
  text-transform: uppercase;
}

.musteriYorumlariBaslik h2 {
  margin: 0 0 10px 0;
  color: #171717;
  font-size: 1.75em;
  line-height: 1.25em;
  font-weight: 700;
}

.musteriYorumlariBaslik p {
  margin: 0;
  color: #6C6C6C;
  font-size: 0.98em;
  line-height: 1.65em;
}

.yorumSlider {
  position: relative;
  max-width: 860px;
  margin: 0 auto;
}

.yorumKart {
  display: none;
  min-height: 210px;
  padding: 42px 70px 38px 70px;
  box-sizing: border-box;
  background: #FFFFFF;
  border: 1px solid #E7E7E7;
  border-radius: 14px;
  box-shadow: 0 12px 34px rgba(0, 0, 0, 0.055);
  text-align: center;
  animation: yorumGiris 0.35s ease;
}

.yorumKartAktif {
  display: block;
}

.yorumIkon {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 42px;
  height: 42px;
  margin: 0 auto 20px auto;
  border-radius: 50%;
  background: #FFF1F1;
  color: #DE0303;
  font-size: 1.25em;
}

.yorumMetni {
  margin: 0;
  color: #292929;
  font-size: 1.12em;
  line-height: 1.8em;
  font-style: normal;
}

.yorumAltBilgi {
  display: block;
  margin-top: 22px;
  color: #777777;
  font-size: 0.82em;
  font-weight: 700;
  letter-spacing: 0.08em;
  text-transform: uppercase;
}

.yorumYon {
  position: absolute;
  top: 50%;
  z-index: 3;
  width: 42px;
  height: 42px;
  margin-top: -21px;
  padding: 0;
  border: 1px solid #E1E1E1;
  border-radius: 50%;
  background: #FFFFFF;
  color: #222222;
  font-size: 1.15em;
  line-height: 40px;
  text-align: center;
  cursor: pointer;
  transition: border-color 0.2s ease, color 0.2s ease, transform 0.2s ease;
}

.yorumYon:hover {
  border-color: #DE0303;
  color: #DE0303;
  transform: translateY(-1px);
}

.yorumOnceki {
  left: -21px;
}

.yorumSonraki {
  right: -21px;
}

.yorumNoktalar {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  margin-top: 20px;
}

.yorumNokta {
  width: 8px;
  height: 8px;
  padding: 0;
  border: 0;
  border-radius: 50%;
  background: #CFCFCF;
  cursor: pointer;
  transition: width 0.2s ease, border-radius 0.2s ease, background 0.2s ease;
}

.yorumNoktaAktif {
  width: 24px;
  border-radius: 8px;
  background: #DE0303;
}

.yorumAltAlan {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 28px;
  max-width: 860px;
  margin: 30px auto 0 auto;
  padding-top: 24px;
  border-top: 1px solid #E4E4E4;
}

.yorumAltMetin {
  max-width: 590px;
}

.yorumAltMetin strong {
  display: block;
  margin-bottom: 5px;
  color: #222222;
  font-size: 0.98em;
}

.yorumAltMetin span {
  color: #707070;
  font-size: 0.9em;
  line-height: 1.55em;
}

.yorumAksiyonlar {
  display: flex;
  align-items: center;
  gap: 10px;
  flex-shrink: 0;
}

.yorumButon {
  display: inline-block;
  padding: 12px 17px;
  border: 1px solid #DE0303;
  border-radius: 6px;
  color: #DE0303;
  background: transparent;
  font-size: 0.86em;
  font-weight: 700;
  text-decoration: none;
  transition: background 0.2s ease, color 0.2s ease;
}

.yorumButon:hover {
  background: #DE0303;
  color: #FFFFFF;
}

.yorumSosyal {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 39px;
  height: 39px;
  border: 1px solid #DDDDDD;
  border-radius: 50%;
  color: #444444;
  background: #FFFFFF;
  text-decoration: none;
  transition: border-color 0.2s ease, color 0.2s ease;
}

.yorumSosyal:hover {
  border-color: #DE0303;
  color: #DE0303;
}

.yorumBos {
  padding: 34px;
  background: #FFFFFF;
  border: 1px solid #E7E7E7;
  border-radius: 14px;
  color: #666666;
  text-align: center;
}

@keyframes yorumGiris {
  from { opacity: 0; transform: translateY(5px); }
  to { opacity: 1; transform: translateY(0); }
}

@media only screen and (max-width: 760px) {
  .musteriYorumlari {
    width: 92%;
    margin: 28px auto 38px auto;
    padding: 36px 5%;
  }

  .musteriYorumlariBaslik h2 {
    font-size: 1.4em;
  }

  .yorumKart {
    min-height: 235px;
    padding: 34px 28px 30px 28px;
  }

  .yorumMetni {
    font-size: 1em;
    line-height: 1.7em;
  }

  .yorumYon {
    top: auto;
    bottom: -62px;
  }

  .yorumOnceki {
    left: calc(50% - 52px);
  }

  .yorumSonraki {
    right: calc(50% - 52px);
  }

  .yorumNoktalar {
    margin-top: 24px;
    padding-bottom: 52px;
  }

  .yorumAltAlan {
    display: block;
    margin-top: 18px;
    text-align: center;
  }

  .yorumAksiyonlar {
    justify-content: center;
    flex-wrap: wrap;
    margin-top: 18px;
  }
}



/* =========================
   MODERN REFERANS AKIŞI
   ========================= */
.referansBolumu {
  width: 100%;
  margin: 42px 0 58px 0;
  padding: 46px 0 38px 0;
  box-sizing: border-box;
  overflow: hidden;
  background: #FFFFFF;
  border-top: 1px solid #E9E9E9;
  border-bottom: 1px solid #E9E9E9;
  font-family: Arial, Helvetica, sans-serif;
}

.referansBaslik {
  width: 92%;
  max-width: 1120px;
  margin: 0 auto 30px auto;
  text-align: center;
}

.referansEtiket {
  display: inline-block;
  margin-bottom: 10px;
  color: #DE0303;
  font-size: 0.78em;
  line-height: 1;
  font-weight: 700;
  letter-spacing: 0.12em;
  text-transform: uppercase;
}

.referansBaslik h2 {
  margin: 0 0 10px 0;
  color: #171717;
  font-size: 1.75em;
  line-height: 1.25em;
  font-weight: 700;
}

.referansBaslik p {
  max-width: 720px;
  margin: 0 auto;
  color: #6C6C6C;
  font-size: 0.98em;
  line-height: 1.65em;
}

.referansAkis {
  position: relative;
  width: 100%;
  overflow: hidden;
}

.referansAkis::before,
.referansAkis::after {
  content: "";
  position: absolute;
  top: 0;
  bottom: 0;
  z-index: 2;
  width: 90px;
  pointer-events: none;
}

.referansAkis::before {
  left: 0;
  background: linear-gradient(to right, #FFFFFF 10%, rgba(255,255,255,0));
}

.referansAkis::after {
  right: 0;
  background: linear-gradient(to left, #FFFFFF 10%, rgba(255,255,255,0));
}

.referansTrack {
  display: flex;
  align-items: center;
  width: max-content;
  animation: referansKaydir 42s linear infinite;
  will-change: transform;
}

.referansGrup {
  display: flex;
  align-items: center;
  flex-shrink: 0;
  gap: 24px;
  padding-right: 24px;
}

.referansLogo {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 210px;
  height: 112px;
  padding: 20px 24px;
  box-sizing: border-box;
  background: #FFFFFF;
  border: 1px solid #E8E8E8;
  border-radius: 10px;
  transition: border-color 0.2s ease, transform 0.2s ease, box-shadow 0.2s ease;
}

.referansLogo:hover {
  border-color: #D5D5D5;
  transform: translateY(-2px);
  box-shadow: 0 10px 24px rgba(0,0,0,0.05);
}

.referansLogo img {
  display: block;
  width: 100%;
  height: 100%;
  max-width: 160px;
  max-height: 70px;
  object-fit: contain;
  filter: grayscale(100%);
  opacity: 0.78;
  transition: filter 0.2s ease, opacity 0.2s ease;
}

.referansLogo:hover img {
  filter: grayscale(0%);
  opacity: 1;
}

.referansAlt {
  width: 92%;
  max-width: 1120px;
  margin: 26px auto 0 auto;
  text-align: center;
}

.referansAlt a {
  display: inline-block;
  color: #DE0303;
  font-size: 0.9em;
  line-height: 1.5em;
  font-weight: 700;
  text-decoration: none;
  border-bottom: 1px solid transparent;
}

.referansAlt a:hover {
  border-bottom-color: #DE0303;
}

.referansBos {
  width: 92%;
  max-width: 1120px;
  margin: 0 auto;
  padding: 28px;
  box-sizing: border-box;
  background: #F8F8F8;
  border: 1px solid #E8E8E8;
  border-radius: 10px;
  color: #666666;
  text-align: center;
}

@keyframes referansKaydir {
  from { transform: translate3d(0, 0, 0); }
  to { transform: translate3d(-50%, 0, 0); }
}

.referansAkis:hover .referansTrack {
  animation-play-state: paused;
}

@media (prefers-reduced-motion: reduce) {
  .referansTrack {
    animation: none;
  }

  .referansAkis {
    overflow-x: auto;
  }
}

@media only screen and (max-width: 760px) {
  .referansBolumu {
    margin: 30px 0 42px 0;
    padding: 36px 0 30px 0;
  }

  .referansBaslik {
    margin-bottom: 24px;
  }

  .referansBaslik h2 {
    font-size: 1.4em;
  }

  .referansAkis::before,
  .referansAkis::after {
    width: 34px;
  }

  .referansTrack {
    animation-duration: 42s;
  }

  .referansGrup {
    gap: 14px;
    padding-right: 14px;
  }

  .referansLogo {
    width: 164px;
    height: 94px;
    padding: 17px 20px;
  }
}



/* =========================
   2026 UYUMLU İKİ CAROUSEL
   ========================= */
.ozumCarousel {
  position: relative;
  width: min(100%, 1270px);
  margin: 0 auto;
  overflow: hidden;
  background: #F4F4F4;
  isolation: isolate;
}

.ozumCarousel .flexslider,
.ozumCarousel .flex-viewport,
.ozumCarousel .slides {
  width: 100%;
  border: 0;
}

.ozumCarousel .flexslider {
  position: relative;
  margin: 0;
  background: #F4F4F4;
  box-shadow: none;
}

.ozumHeroCarousel,
.ozumHeroCarousel .flexslider,
.ozumHeroCarousel .flex-viewport,
.ozumHeroCarousel .slides,
.ozumHeroCarousel .slides > li {
  height: clamp(420px, 47vw, 600px);
}

.ozumGaleriCarousel,
.ozumGaleriCarousel .flexslider,
.ozumGaleriCarousel .flex-viewport,
.ozumGaleriCarousel .slides,
.ozumGaleriCarousel .slides > li {
  height: clamp(390px, 44vw, 560px);
}

.ozumCarousel .slides > li {
  max-width: none;
  background-position: center;
  background-repeat: no-repeat;
  background-size: cover;
}

.ozumCarousel::after {
  content: "";
  position: absolute;
  inset: auto 0 0;
  z-index: 1;
  height: 22%;
  pointer-events: none;
  background: linear-gradient(to top, rgba(0,0,0,0.2), rgba(0,0,0,0));
}

.ozumCarousel .flex-direction-nav a {
  z-index: 5;
  display: flex;
  align-items: center;
  justify-content: center;
  width: 46px;
  height: 46px;
  margin-top: -23px;
  overflow: hidden;
  border: 1px solid rgba(255,255,255,0.55);
  border-radius: 50%;
  background: rgba(18,18,18,0.48);
  color: #FFFFFF;
  opacity: 1;
  text-shadow: none;
  backdrop-filter: blur(8px);
  -webkit-backdrop-filter: blur(8px);
  transition: background 0.2s ease, border-color 0.2s ease, transform 0.2s ease;
}

.ozumCarousel .flex-direction-nav a::before {
  color: #FFFFFF;
  font-size: 24px;
  line-height: 1;
  text-shadow: none;
}

.ozumCarousel .flex-direction-nav .flex-prev {
  left: 22px;
}

.ozumCarousel .flex-direction-nav .flex-next {
  right: 22px;
}

.ozumCarousel:hover .flex-direction-nav .flex-prev {
  left: 22px;
}

.ozumCarousel:hover .flex-direction-nav .flex-next {
  right: 22px;
}

.ozumCarousel .flex-direction-nav a:hover,
.ozumCarousel .flex-direction-nav a:focus-visible {
  border-color: #FFFFFF;
  background: rgba(222,3,3,0.9);
  transform: scale(1.04);
  outline: 3px solid rgba(255,255,255,0.72);
  outline-offset: 2px;
}

.ozumCarousel .flex-control-nav {
  z-index: 5;
  bottom: 18px;
  display: flex;
  justify-content: center;
  gap: 8px;
  height: auto;
  line-height: 1;
}

.ozumCarousel .flex-control-nav li {
  margin: 0;
}

.ozumCarousel .flex-control-paging li a {
  width: 8px;
  height: 8px;
  background: rgba(255,255,255,0.62);
  box-shadow: none;
  transition: width 0.2s ease, border-radius 0.2s ease, background 0.2s ease;
}

.ozumCarousel .flex-control-paging li a.flex-active {
  width: 25px;
  border-radius: 8px;
  background: #FFFFFF;
}

.ozumCarousel .flex-control-paging li a:focus-visible {
  outline: 3px solid #FFFFFF;
  outline-offset: 3px;
}

.carouselDuraklat {
  position: absolute;
  right: 20px;
  bottom: 13px;
  z-index: 6;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 7px;
  min-width: 42px;
  min-height: 42px;
  padding: 0 13px;
  border: 1px solid rgba(255,255,255,0.55);
  border-radius: 22px;
  background: rgba(18,18,18,0.48);
  color: #FFFFFF;
  font: 700 12px/1 Arial, Helvetica, sans-serif;
  cursor: pointer;
  backdrop-filter: blur(8px);
  -webkit-backdrop-filter: blur(8px);
}

.carouselDuraklat:hover,
.carouselDuraklat:focus-visible {
  border-color: #FFFFFF;
  background: rgba(222,3,3,0.9);
  outline: 3px solid rgba(255,255,255,0.72);
  outline-offset: 2px;
}

.carouselDuraklat i {
  font-size: 11px;
}

.ozumGaleriCarousel {
  margin-top: 8px;
  border-top: 1px solid #E8E8E8;
  border-bottom: 1px solid #E8E8E8;
}

.ozumGaleriCarousel .slides > li {
  cursor: pointer;
}

@media only screen and (max-width: 760px) {
  .ozumHeroCarousel,
  .ozumHeroCarousel .flexslider,
  .ozumHeroCarousel .flex-viewport,
  .ozumHeroCarousel .slides,
  .ozumHeroCarousel .slides > li {
    height: clamp(330px, 112vw, 500px);
  }

  .ozumGaleriCarousel,
  .ozumGaleriCarousel .flexslider,
  .ozumGaleriCarousel .flex-viewport,
  .ozumGaleriCarousel .slides,
  .ozumGaleriCarousel .slides > li {
    height: clamp(300px, 88vw, 440px);
  }

  .ozumCarousel .flex-direction-nav a {
    width: 44px;
    height: 44px;
    margin-top: -22px;
  }

  .ozumCarousel .flex-direction-nav .flex-prev,
  .ozumCarousel:hover .flex-direction-nav .flex-prev {
    left: 10px;
  }

  .ozumCarousel .flex-direction-nav .flex-next,
  .ozumCarousel:hover .flex-direction-nav .flex-next {
    right: 10px;
  }

  .ozumCarousel .flex-control-nav {
    bottom: 15px;
    max-width: calc(100% - 110px);
    margin-left: 55px;
  }

  .carouselDuraklat {
    right: 10px;
    bottom: 8px;
    width: 44px;
    min-width: 44px;
    height: 44px;
    padding: 0;
    border-radius: 50%;
  }

  .carouselDuraklat span {
    position: absolute;
    width: 1px;
    height: 1px;
    overflow: hidden;
    clip: rect(0 0 0 0);
    white-space: nowrap;
  }
}

@media only screen and (max-width: 420px) {
  .ozumHeroCarousel,
  .ozumHeroCarousel .flexslider,
  .ozumHeroCarousel .flex-viewport,
  .ozumHeroCarousel .slides,
  .ozumHeroCarousel .slides > li {
    height: 390px;
  }

  .ozumGaleriCarousel,
  .ozumGaleriCarousel .flexslider,
  .ozumGaleriCarousel .flex-viewport,
  .ozumGaleriCarousel .slides,
  .ozumGaleriCarousel .slides > li {
    height: 330px;
  }
}

@media (prefers-reduced-motion: reduce) {
  .ozumCarousel *,
  .ozumCarousel *::before,
  .ozumCarousel *::after {
    scroll-behavior: auto !important;
    transition-duration: 0.01ms !important;
    animation-duration: 0.01ms !important;
    animation-iteration-count: 1 !important;
  }
}



/* =========================
   CAROUSEL SON CİLA
   ========================= */
.ozumCarousel {
  width: 92%;
  max-width: 1180px;
  margin-left: auto;
  margin-right: auto;
  border: 1px solid #E8E8E8;
  border-radius: 12px;
  background: #F3F3F3;
  box-shadow: 0 14px 38px rgba(0,0,0,0.055);
}

.ozumHeroCarousel {
  margin-top: 28px;
  margin-bottom: 20px;
}

.ozumHeroCarousel,
.ozumHeroCarousel .flexslider,
.ozumHeroCarousel .flex-viewport,
.ozumHeroCarousel .slides,
.ozumHeroCarousel .slides > li {
  height: clamp(410px, 36vw, 540px);
}

.ozumGaleriCarousel {
  margin-top: 32px;
  margin-bottom: 48px;
}

.ozumGaleriCarousel,
.ozumGaleriCarousel .flexslider,
.ozumGaleriCarousel .flex-viewport,
.ozumGaleriCarousel .slides,
.ozumGaleriCarousel .slides > li {
  height: clamp(340px, 30vw, 455px);
}

/* FlexSlider'ın font ikonunu kapat; oku geometrik olarak ortala. */
.ozumCarousel .flex-direction-nav a {
  width: 46px;
  height: 46px;
  margin-top: -23px;
  font-size: 0;
  line-height: 1;
  overflow: visible;
  border: 1px solid rgba(255,255,255,0.66);
  background: rgba(20,20,20,0.46);
}

.ozumCarousel .flex-direction-nav a::before {
  content: "";
  display: block;
  width: 10px;
  height: 10px;
  margin: 0;
  border-top: 2px solid #FFFFFF;
  border-right: 2px solid #FFFFFF;
  color: transparent;
  font-size: 0;
  line-height: 0;
}

.ozumCarousel .flex-direction-nav .flex-prev::before {
  transform: rotate(-135deg);
  margin-left: 4px;
}

.ozumCarousel .flex-direction-nav .flex-next::before {
  transform: rotate(45deg);
  margin-right: 4px;
}

.ozumCarousel .flex-direction-nav .flex-prev,
.ozumCarousel:hover .flex-direction-nav .flex-prev {
  left: 18px;
}

.ozumCarousel .flex-direction-nav .flex-next,
.ozumCarousel:hover .flex-direction-nav .flex-next {
  right: 18px;
}

/* Üstte az sayıda slide için sade pagination. */
.ozumHeroCarousel .flex-control-nav {
  bottom: 15px;
}

/* Alt galeride yüzlerce nokta asla gösterilmez. */
.ozumGaleriCarousel .flex-control-nav {
  display: none !important;
}

.ozumGaleriCarousel .carouselDuraklat {
  bottom: 14px;
}

/* Görselleri kırpmadan daha editoryal bir odak. */
.ozumHeroCarousel .slides > li {
  background-position: center center;
}

.ozumGaleriCarousel .slides > li {
  background-position: center 44%;
}

@media only screen and (max-width: 760px) {
  .ozumCarousel {
    width: calc(100% - 24px);
    border-radius: 10px;
  }

  .ozumHeroCarousel {
    margin-top: 14px;
    margin-bottom: 14px;
  }

  .ozumHeroCarousel,
  .ozumHeroCarousel .flexslider,
  .ozumHeroCarousel .flex-viewport,
  .ozumHeroCarousel .slides,
  .ozumHeroCarousel .slides > li {
    height: clamp(300px, 88vw, 420px);
  }

  .ozumGaleriCarousel {
    margin-top: 24px;
    margin-bottom: 34px;
  }

  .ozumGaleriCarousel,
  .ozumGaleriCarousel .flexslider,
  .ozumGaleriCarousel .flex-viewport,
  .ozumGaleriCarousel .slides,
  .ozumGaleriCarousel .slides > li {
    height: clamp(260px, 70vw, 350px);
  }

  .ozumCarousel .flex-direction-nav a {
    width: 42px;
    height: 42px;
    margin-top: -21px;
  }

  .ozumCarousel .flex-direction-nav .flex-prev,
  .ozumCarousel:hover .flex-direction-nav .flex-prev {
    left: 9px;
  }

  .ozumCarousel .flex-direction-nav .flex-next,
  .ozumCarousel:hover .flex-direction-nav .flex-next {
    right: 9px;
  }

  .ozumHeroCarousel .flex-control-nav {
    max-width: calc(100% - 96px);
    margin-left: 48px;
  }
}

@media only screen and (max-width: 420px) {
  .ozumHeroCarousel,
  .ozumHeroCarousel .flexslider,
  .ozumHeroCarousel .flex-viewport,
  .ozumHeroCarousel .slides,
  .ozumHeroCarousel .slides > li {
    height: 330px;
  }

  .ozumGaleriCarousel,
  .ozumGaleriCarousel .flexslider,
  .ozumGaleriCarousel .flex-viewport,
  .ozumGaleriCarousel .slides,
  .ozumGaleriCarousel .slides > li {
    height: 280px;
  }
}



/* =========================
   OK MERKEZLEME KESİN DÜZELTME
   ========================= */
.ozumCarousel .flex-direction-nav a {
  position: absolute;
  display: block;
  width: 46px;
  height: 46px;
  margin-top: -23px;
  padding: 0;
  overflow: hidden !important;
  text-indent: -9999px !important;
  font-size: 0 !important;
  line-height: 0 !important;
  color: transparent !important;
}

.ozumCarousel .flex-direction-nav a::before {
  content: "" !important;
  position: absolute !important;
  top: 50% !important;
  left: 50% !important;
  display: block !important;
  width: 10px !important;
  height: 10px !important;
  margin: 0 !important;
  padding: 0 !important;
  border: 0 !important;
  border-top: 2px solid #FFFFFF !important;
  border-right: 2px solid #FFFFFF !important;
  background: transparent !important;
  color: transparent !important;
  font-size: 0 !important;
  line-height: 0 !important;
  text-indent: 0 !important;
  transform-origin: center center !important;
}

.ozumCarousel .flex-direction-nav .flex-prev::before {
  transform: translate(-42%, -50%) rotate(-135deg) !important;
}

.ozumCarousel .flex-direction-nav .flex-next::before {
  transform: translate(-58%, -50%) rotate(45deg) !important;
}

/* FlexSlider teması after veya başka pseudo içerik ekliyorsa kapat. */
.ozumCarousel .flex-direction-nav a::after {
  content: none !important;
  display: none !important;
}

@media only screen and (max-width: 760px) {
  .ozumCarousel .flex-direction-nav a {
    width: 42px;
    height: 42px;
    margin-top: -21px;
  }

  .ozumCarousel .flex-direction-nav a::before {
    width: 9px !important;
    height: 9px !important;
  }
}



/* =========================
   MOBİL TAM GÖRSEL + AUTOPLAY FIX
   ========================= */
@media only screen and (max-width: 760px) {
  .ozumHeroCarousel {
    aspect-ratio: 16 / 9;
    height: auto !important;
    background: #111111;
  }

  .ozumGaleriCarousel {
    aspect-ratio: 4 / 3;
    height: auto !important;
    background: #F3F3F3;
  }

  .ozumHeroCarousel .flexslider,
  .ozumHeroCarousel .flex-viewport,
  .ozumHeroCarousel .slides,
  .ozumHeroCarousel .slides > li,
  .ozumGaleriCarousel .flexslider,
  .ozumGaleriCarousel .flex-viewport,
  .ozumGaleriCarousel .slides,
  .ozumGaleriCarousel .slides > li {
    height: 100% !important;
    min-height: 0 !important;
  }

  .ozumHeroCarousel .slides > li,
  .ozumGaleriCarousel .slides > li {
    background-size: contain !important;
    background-position: center center !important;
    background-repeat: no-repeat !important;
  }
}



/* =========================
   KALICI MOBİL RENDER DÜZELTMESİ
   ========================= */
@media only screen and (max-width: 760px) {
  .ozumHeroCarousel,
  .ozumHeroCarousel .flexslider,
  .ozumHeroCarousel .flex-viewport,
  .ozumHeroCarousel .slides,
  .ozumHeroCarousel .slides > li {
    width: 100% !important;
    height: clamp(230px, 56.25vw, 420px) !important;
    min-height: 230px !important;
  }

  .ozumGaleriCarousel,
  .ozumGaleriCarousel .flexslider,
  .ozumGaleriCarousel .flex-viewport,
  .ozumGaleriCarousel .slides,
  .ozumGaleriCarousel .slides > li {
    width: 100% !important;
    height: clamp(260px, 75vw, 390px) !important;
    min-height: 260px !important;
  }

  .ozumHeroCarousel .slides > li,
  .ozumGaleriCarousel .slides > li {
    background-size: contain !important;
    background-position: center center !important;
    background-repeat: no-repeat !important;
  }

  .ozumHeroCarousel {
    background: #090909 !important;
  }

  .ozumGaleriCarousel {
    background: #F2F2F2 !important;
  }
}

/* FlexSlider'ın JS yüklenmesinden önce ilk görsel görünür kalsın. */
.ozumCarousel .slides > li:first-child {
  display: block;
}



/* =========================
   VANILLA CAROUSEL MİMARİSİ
   ========================= */
.carouselDuraklat,
.ozumCarousel .flex-direction-nav,
.ozumCarousel .flex-control-nav {
  display: none !important;
}

.ozumCarousel.vanillaCarouselHazir {
  position: relative;
  overflow: hidden;
}

.ozumCarousel.vanillaCarouselHazir .flexslider,
.ozumCarousel.vanillaCarouselHazir .slides,
.ozumCarousel.vanillaCarouselHazir .slides > li {
  width: 100% !important;
  height: 100% !important;
  margin: 0 !important;
  padding: 0 !important;
}

.ozumCarousel.vanillaCarouselHazir .flexslider {
  position: static !important;
  overflow: visible !important;
}

.ozumCarousel.vanillaCarouselHazir .slides {
  position: relative !important;
  display: block !important;
}

.ozumCarousel.vanillaCarouselHazir .slides > li {
  position: absolute !important;
  inset: 0 !important;
  display: block !important;
  opacity: 0;
  visibility: hidden;
  pointer-events: none;
  transform: scale(1.008);
  transition: opacity 0.55s ease, visibility 0.55s ease, transform 0.8s ease;
  background-repeat: no-repeat !important;
  background-position: center center !important;
  background-size: cover !important;
}

.ozumCarousel.vanillaCarouselHazir .slides > li.vanillaAktif {
  z-index: 1;
  opacity: 1;
  visibility: visible;
  pointer-events: auto;
  transform: scale(1);
}

.vanillaCarouselOk {
  position: absolute;
  top: 50%;
  z-index: 8;
  width: 46px;
  height: 46px;
  margin-top: -23px;
  padding: 0;
  border: 1px solid rgba(255,255,255,0.72);
  border-radius: 50%;
  background: rgba(20,20,20,0.5);
  cursor: pointer;
  backdrop-filter: blur(7px);
  -webkit-backdrop-filter: blur(7px);
}

.vanillaCarouselOk::before {
  content: "";
  position: absolute;
  top: 50%;
  left: 50%;
  width: 10px;
  height: 10px;
  border-top: 2px solid #FFFFFF;
  border-right: 2px solid #FFFFFF;
}

.vanillaCarouselOnceki {
  left: 18px;
}

.vanillaCarouselOnceki::before {
  transform: translate(-38%, -50%) rotate(-135deg);
}

.vanillaCarouselSonraki {
  right: 18px;
}

.vanillaCarouselSonraki::before {
  transform: translate(-62%, -50%) rotate(45deg);
}

.vanillaCarouselOk:hover,
.vanillaCarouselOk:focus-visible {
  background: rgba(222,3,3,0.9);
  outline: 3px solid rgba(255,255,255,0.75);
  outline-offset: 2px;
}

.vanillaCarouselNoktalar {
  position: absolute;
  left: 50%;
  bottom: 16px;
  z-index: 8;
  display: flex;
  align-items: center;
  gap: 8px;
  transform: translateX(-50%);
}

.vanillaCarouselNokta {
  width: 8px;
  height: 8px;
  padding: 0;
  border: 0;
  border-radius: 999px;
  background: rgba(255,255,255,0.65);
  cursor: pointer;
  transition: width 0.2s ease, background 0.2s ease;
}

.vanillaCarouselNokta.vanillaNoktaAktif {
  width: 24px;
  background: #FFFFFF;
}

@media only screen and (min-width: 761px) {
  .ozumHeroCarousel {
    height: clamp(390px, 34vw, 530px) !important;
  }

  .ozumGaleriCarousel {
    height: clamp(330px, 29vw, 450px) !important;
  }
}

@media only screen and (max-width: 760px) {
  .ozumCarousel {
    width: calc(100% - 20px) !important;
    min-height: 0 !important;
    border-radius: 10px;
  }

  .ozumHeroCarousel,
  .ozumGaleriCarousel {
    height: 280px;
  }

  .ozumCarousel.vanillaCarouselHazir .slides > li {
    background-size: contain !important;
    background-color: transparent !important;
  }

  .ozumHeroCarousel {
    background: #090909 !important;
  }

  .ozumGaleriCarousel {
    background: #F2F2F2 !important;
  }

  .vanillaCarouselOk {
    width: 42px;
    height: 42px;
    margin-top: -21px;
  }

  .vanillaCarouselOnceki {
    left: 8px;
  }

  .vanillaCarouselSonraki {
    right: 8px;
  }

  .vanillaCarouselNoktalar {
    bottom: 12px;
  }
}

@media (prefers-reduced-motion: reduce) {
  .ozumCarousel.vanillaCarouselHazir .slides > li {
    transition: none !important;
  }
}



/* =========================
   RESPONSIVE GALERİ ORAN DÜZELTMESİ
   ========================= */
.ozumGaleriCarousel {
  transition: height 0.38s ease;
  background: #F2F2F2 !important;
}

.ozumGaleriCarousel.vanillaCarouselHazir .slides > li {
  background-size: contain !important;
  background-position: center center !important;
  background-repeat: no-repeat !important;
  background-color: #F2F2F2 !important;
}

/* Eski sabit masaüstü yüksekliğini geçersiz kıl. */
@media only screen and (min-width: 761px) {
  .ozumGaleriCarousel {
    height: 520px;
  }

  .ozumGaleriCarousel,
  .ozumGaleriCarousel .flexslider,
  .ozumGaleriCarousel .slides,
  .ozumGaleriCarousel .slides > li {
    min-height: 320px !important;
    max-height: 760px !important;
  }
}

@media only screen and (max-width: 760px) {
  .ozumGaleriCarousel {
    transition-duration: 0.28s;
  }
}

@media (prefers-reduced-motion: reduce) {
  .ozumGaleriCarousel {
    transition: none !important;
  }
}



/* =========================
   RESPONSIVE HİZMET KARTLARI
   ========================= */
.urunkutular {
  display: grid !important;
  grid-template-columns: minmax(0, 1fr) minmax(0, 1.08fr);
  grid-template-areas:
    "ev hava"
    "ticari hava"
    "vrf vrf";
  gap: 18px;
  width: 92% !important;
  max-width: 1120px !important;
  margin: 42px auto 56px auto !important;
  padding: 0 !important;
  box-sizing: border-box;
}

.urunkutular .solSutun,
.urunkutular .sagSutun {
  display: contents !important;
}

.urunkutular .solSutun > a:nth-child(1) {
  grid-area: ev;
}

.urunkutular .solSutun > a:nth-child(2) {
  grid-area: ticari;
}

.urunkutular .solSutun > a:nth-child(3) {
  grid-area: vrf;
}

.urunkutular .sagSutun > a.vrfm {
  display: none !important;
}

.urunkutular .sagSutun > a:last-child {
  grid-area: hava;
}

.urunkutular a {
  position: relative;
  display: block;
  min-width: 0;
  overflow: hidden;
  border-radius: 12px;
  text-decoration: none;
  background: #F2F2F2;
  box-shadow: 0 10px 28px rgba(0,0,0,0.055);
  transition: transform 0.22s ease, box-shadow 0.22s ease;
}

.urunkutular a:hover {
  transform: translateY(-3px);
  box-shadow: 0 16px 34px rgba(0,0,0,0.085);
}

.urunkutular a:focus-visible {
  outline: 3px solid #DE0303;
  outline-offset: 3px;
}

.urunkutular a > div {
  position: relative;
  width: 100% !important;
  max-width: none !important;
  height: auto !important;
  min-height: 0 !important;
  overflow: hidden;
  border-radius: inherit;
  background-position: center center !important;
  background-repeat: no-repeat !important;
  background-size: cover !important;
}

.urunkutular .solSutun > a:nth-child(1) > div,
.urunkutular .solSutun > a:nth-child(2) > div {
  aspect-ratio: 16 / 8.2;
}

.urunkutular .solSutun > a:nth-child(3) > div {
  aspect-ratio: 16 / 5.2;
}

.urunkutular .sagSutun > a:last-child > div {
  height: 100% !important;
  min-height: 100% !important;
}

.urunkutular .sagSutun > a:last-child {
  min-height: 100%;
}

.urunkutular a > div::before {
  content: "";
  position: absolute;
  inset: 0;
  z-index: 1;
  background: linear-gradient(
    to top,
    rgba(0,0,0,0.56) 0%,
    rgba(0,0,0,0.15) 42%,
    rgba(0,0,0,0.02) 70%
  );
  transition: background 0.22s ease;
}

.urunkutular a:hover > div::before {
  background: linear-gradient(
    to top,
    rgba(0,0,0,0.62) 0%,
    rgba(0,0,0,0.2) 44%,
    rgba(0,0,0,0.04) 72%
  );
}

.urunkutular .Absolute-Center {
  position: absolute !important;
  left: 24px !important;
  right: 24px !important;
  bottom: 22px !important;
  top: auto !important;
  z-index: 2;
  width: auto !important;
  height: auto !important;
  transform: none !important;
  text-align: left !important;
}

.urunkutular .Absolute-Center span {
  display: inline-block;
  color: #FFFFFF;
  font-family: Arial, Helvetica, sans-serif;
  font-size: clamp(1.05rem, 1.5vw, 1.45rem);
  line-height: 1.25;
  font-weight: 700;
  text-shadow: 0 2px 12px rgba(0,0,0,0.35);
}

/* Tablet: iki sütun korunur ama koreografi sadeleşir. */
@media only screen and (max-width: 900px) {
  .urunkutular {
    grid-template-columns: repeat(2, minmax(0, 1fr));
    grid-template-areas:
      "ev ticari"
      "hava hava"
      "vrf vrf";
    gap: 14px;
    margin-top: 32px !important;
    margin-bottom: 44px !important;
  }

  .urunkutular .solSutun > a:nth-child(1) > div,
  .urunkutular .solSutun > a:nth-child(2) > div {
    aspect-ratio: 4 / 3;
  }

  .urunkutular .sagSutun > a:last-child > div {
    height: auto !important;
    min-height: 0 !important;
    aspect-ratio: 16 / 7.2;
  }

  .urunkutular .solSutun > a:nth-child(3) > div {
    aspect-ratio: 16 / 7.2;
  }
}

/* Mobil: tek kolon, sabit ve dengeli ritim. */
@media only screen and (max-width: 600px) {
  .urunkutular {
    grid-template-columns: 1fr;
    grid-template-areas:
      "ev"
      "ticari"
      "vrf"
      "hava";
    gap: 12px;
    width: calc(100% - 24px) !important;
    margin: 26px auto 36px auto !important;
  }

  .urunkutular .solSutun > a:nth-child(1) > div,
  .urunkutular .solSutun > a:nth-child(2) > div,
  .urunkutular .solSutun > a:nth-child(3) > div,
  .urunkutular .sagSutun > a:last-child > div {
    height: auto !important;
    min-height: 0 !important;
    aspect-ratio: 16 / 9;
  }

  .urunkutular a {
    border-radius: 10px;
  }

  .urunkutular .Absolute-Center {
    left: 18px !important;
    right: 18px !important;
    bottom: 17px !important;
  }

  .urunkutular .Absolute-Center span {
    font-size: 1.05rem;
  }
}

@media (prefers-reduced-motion: reduce) {
  .urunkutular a,
  .urunkutular a > div::before {
    transition: none !important;
  }
}

</style>  


</head>

<body> 
             

                         

<!--#include file='ust.asp'-->




        
<section id="ana-hero-carousel" class="flexslider-container ozumCarousel ozumHeroCarousel" aria-label="Özüm Klima öne çıkan çözümleri">   
    <div class="flexslider ozumHeroSlider">
        <ul class="slides" >  
  
                      
<%
set object = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From flas "
object.open Sorgula,baglanti,1,3   

Do while not object.Eof                                                                                                                                                                                                                  
%>   
<li   onClick="window.location='<%=object("adres") %>';" style="cursor:pointer; background-image: url(slayt/<%=object("resim") %>);background-size:cover" ></li>  

<%  
object.MoveNExt
Loop
%>

                      
  </ul>
    </div>
</section>   
         
         
<script>
(function () {
  "use strict";

  function hazirlaCarousel(containerId, ayarlar) {
    var container = document.getElementById(containerId);

    if (!container) {
      return;
    }

    var liste = container.querySelector(".slides");
    var slaytlar = liste ? liste.children : [];
    var aktifIndex = 0;
    var zamanlayici = null;
    var noktalar = [];
    var dokunmaBaslangicX = null;
    var hareketAzalt = false;

    if (!liste || !slaytlar.length) {
      return;
    }

    if (container.getAttribute("data-carousel-ready") === "1") {
      return;
    }

    container.setAttribute("data-carousel-ready", "1");
    container.classList.add("vanillaCarouselHazir");

    var onceki = document.createElement("button");
    onceki.type = "button";
    onceki.className = "vanillaCarouselOk vanillaCarouselOnceki";
    onceki.setAttribute("aria-label", "Önceki görsel");

    var sonraki = document.createElement("button");
    sonraki.type = "button";
    sonraki.className = "vanillaCarouselOk vanillaCarouselSonraki";
    sonraki.setAttribute("aria-label", "Sonraki görsel");

    container.appendChild(onceki);
    container.appendChild(sonraki);

    if (ayarlar.noktalar && slaytlar.length > 1) {
      var noktaAlani = document.createElement("div");
      noktaAlani.className = "vanillaCarouselNoktalar";
      noktaAlani.setAttribute("aria-label", "Slayt seçimi");

      for (var n = 0; n < slaytlar.length; n++) {
        (function (index) {
          var nokta = document.createElement("button");
          nokta.type = "button";
          nokta.className = "vanillaCarouselNokta";
          nokta.setAttribute("aria-label", (index + 1) + ". görsele git");
          nokta.addEventListener("click", function () {
            goster(index, true);
          });
          noktalar.push(nokta);
          noktaAlani.appendChild(nokta);
        })(n);
      }

      container.appendChild(noktaAlani);
    }

    function arkaPlanUrl(slide) {
      var stil = window.getComputedStyle(slide);
      var deger = stil.backgroundImage || "";
      var eslesme = deger.match(/^url\(["']?(.*?)["']?\)$/);
      return eslesme ? eslesme[1] : "";
    }

    function carouselYukseklikAyarla(slide) {
      var url = arkaPlanUrl(slide);

      if (!url) {
        container.style.height = ayarlar.varsayilanYukseklik(window.innerWidth) + "px";
        return;
      }

      var img = new Image();

      img.onload = function () {
        var genislik = container.clientWidth;
        var oran = img.naturalHeight / img.naturalWidth;
        var yukseklik = Math.round(genislik * oran);
        var sinirlar = ayarlar.yukseklikSinirlari(window.innerWidth);

        yukseklik = Math.max(sinirlar.min, Math.min(yukseklik, sinirlar.max));
        container.style.height = yukseklik + "px";
      };

      img.onerror = function () {
        container.style.height = ayarlar.varsayilanYukseklik(window.innerWidth) + "px";
      };

      img.src = url;
    }

    function goster(index, kullaniciEtkilesimi) {
      if (index >= slaytlar.length) {
        index = 0;
      }

      if (index < 0) {
        index = slaytlar.length - 1;
      }

      for (var i = 0; i < slaytlar.length; i++) {
        var aktif = i === index;
        slaytlar[i].classList.toggle("vanillaAktif", aktif);
        slaytlar[i].setAttribute("aria-hidden", aktif ? "false" : "true");
      }

      for (var j = 0; j < noktalar.length; j++) {
        var noktaAktif = j === index;
        noktalar[j].classList.toggle("vanillaNoktaAktif", noktaAktif);
        noktalar[j].setAttribute("aria-current", noktaAktif ? "true" : "false");
      }

      aktifIndex = index;
      carouselYukseklikAyarla(slaytlar[aktifIndex]);

      if (kullaniciEtkilesimi) {
        yenidenBaslat();
      }
    }

    function sonrakiSlayt(kullaniciEtkilesimi) {
      goster(aktifIndex + 1, kullaniciEtkilesimi);
    }

    function oncekiSlayt(kullaniciEtkilesimi) {
      goster(aktifIndex - 1, kullaniciEtkilesimi);
    }

    function durdur() {
      if (zamanlayici !== null) {
        window.clearTimeout(zamanlayici);
        zamanlayici = null;
      }
    }

    function zamanla() {
      durdur();

      if (slaytlar.length < 2 || document.hidden) {
        return;
      }

      zamanlayici = window.setTimeout(function () {
        sonrakiSlayt(false);
        zamanla();
      }, 5000);
    }

    function baslat() {
      zamanla();
    }

    function yenidenBaslat() {
      zamanla();
    }

    onceki.addEventListener("click", function () {
      oncekiSlayt(true);
    });

    sonraki.addEventListener("click", function () {
      sonrakiSlayt(true);
    });

    container.addEventListener("keydown", function (event) {
      if (event.key === "ArrowLeft") {
        event.preventDefault();
        oncekiSlayt(true);
      } else if (event.key === "ArrowRight") {
        event.preventDefault();
        sonrakiSlayt(true);
      }
    });

    container.addEventListener("touchstart", function (event) {
      if (event.touches && event.touches.length === 1) {
        dokunmaBaslangicX = event.touches[0].clientX;
      }
    }, { passive: true });

    container.addEventListener("touchend", function (event) {
      if (dokunmaBaslangicX === null || !event.changedTouches || !event.changedTouches.length) {
        return;
      }

      var fark = event.changedTouches[0].clientX - dokunmaBaslangicX;
      dokunmaBaslangicX = null;

      if (Math.abs(fark) < 45) {
        return;
      }

      if (fark < 0) {
        sonrakiSlayt(true);
      } else {
        oncekiSlayt(true);
      }
    }, { passive: true });

    window.addEventListener("resize", function () {
      carouselYukseklikAyarla(slaytlar[aktifIndex]);
    });

    document.addEventListener("visibilitychange", function () {
      if (document.hidden) {
        durdur();
      } else {
        baslat();
      }
    });

    goster(0, false);
    baslat();
  }

  function baslatCarouseller() {
    hazirlaCarousel("ana-hero-carousel", {
      noktalar: true,
      yukseklikSinirlari: function (genislik) {
        if (genislik <= 760) {
          return { min: 190, max: 430 };
        }

        return { min: 390, max: 530 };
      },
      varsayilanYukseklik: function (genislik) {
        return genislik <= 760 ? 260 : 500;
      }
    });

    hazirlaCarousel("uygulama-galeri-carousel", {
      noktalar: false,
      yukseklikSinirlari: function (genislik) {
        if (genislik <= 480) {
          return { min: 220, max: 520 };
        }

        if (genislik <= 760) {
          return { min: 260, max: 560 };
        }

        if (genislik <= 1180) {
          return { min: 320, max: 720 };
        }

        return { min: 360, max: 760 };
      },
      varsayilanYukseklik: function (genislik) {
        if (genislik <= 480) {
          return 300;
        }

        if (genislik <= 760) {
          return 360;
        }

        return 520;
      }
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", baslatCarouseller);
  } else {
    baslatCarouseller();
  }

  window.addEventListener("load", function () {
    baslatCarouseller();
  });
})();
</script>     
            

<BR>
<!---------------------------------------------------------------------------------------------slider----------------------------------------------------------------------------------->

<Div class="slogan">
 <span class="slo2">     &#304;htiyac&#305;n&#305;za uygun konfor ve enerji tasarrufu sa&#287;layan klimalar  </span><BR>
<span class="slo1">D&ouml;rt mevsim profesyonel &#231;&ouml;z&uuml;mler</span>  
</Div>




<Div  class="kutular">

           <a href="klima.asp"> 
           <Div>  
                  
                 <div class="siyahlik" >&nbsp;</div>       
                <div class="iconTas" >  
             <i class="fas fa-asterisk"></i>        <BR>   <BR>   
Klima Modelleri   <BR>   <BR>  
             <span>
Y&uuml;ksek teknoloji, tasarruflu ve verimli klima modellerimizi incelemek i&#231;in t&#305;klay&#305;n.
</span>
                  </div>
               
           
           </Div>
           </a>          
           
           
           
           <a href="vrf_sistemler.asp"> 
           <Div>  
                  
                 <div class="siyahlik" >&nbsp;</div>       
                <div class="iconTas" >      
                <i class="fas fa-cubes"></i> <BR> <BR>  
   VRF Sistemler      <BR> <BR>  
                <span>
                Bir dış &uuml;nite ve i&#231; mekan s&#305;cakl&#305;&#287;&#305;n&#305; d&uuml;zenleyen birden fazla i&#231; &uuml;niteden olu&#351;maktad&#305;r. 
</span>
                  </div>
               
           
           </Div>
           </a>          
           
           
           <a href="kesfetteam.asp"> 
           <Div>  
                  
                 <div class="siyahlik" >&nbsp;</div>       
                <div class="iconTas" >   
<i class="fas fa-street-view"></i> <BR>   <BR>  
 Ke&#351;fetteam     <BR> <BR>     
              <span>
Yanl&#305;&#351; tercih tad&#305;n&#305;z&#305; ka&#231;&#305;r&#305;r. Ke&#351;fetteam'i &#231;a&#287;&#305;r&#305;n uzman ekibimizle birlikte ke&#351;fedelim.
</span>
                  </div>
               
           
           </Div>
           </a>          
           
           
           
           
           <a href="iletisim.asp"> 
           <Div >  
                  
                 <div class="siyahlik" >&nbsp;</div>       
                <div class="iconTas" >
    <i class="fas fa-street-view"></i><BR>   <BR>  
        Sat&#305;&#351; &amp; Servis <BR> <BR>  
                       <span>
&#304;klimlendirme sistemleri ile ilgili sat&#305;&#351; ve servis hizmetlerimizle sizinleyiz.
</span>
                  </div>
               
           
           </Div>
           </a>          
           
</Div>            

<section class="hizmetGovdesi" aria-labelledby="hizmet-baslik">
  <h1 id="hizmet-baslik">Antalya Klima Servisi, VRF ve Havaland&#305;rma &#199;&ouml;z&uuml;mleri</h1>
  <span class="hizmetAltBaslik">Ev, ofis ve ticari alanlar i&#231;in do&#287;ru ke&#351;if, verimli cihaz se&#231;imi, profesyonel montaj ve sat&#305;&#351; sonras&#305; teknik servis.</span>
  <div class="hizmetMetin">
    <p>Antalya'n&#305;n s&#305;cak ve nemli ikliminde klima se&#231;imi yaln&#305;zca serinlemek i&#231;in de&#287;il, konforu korumak ve enerji t&uuml;ketimini kontrol alt&#305;nda tutmak i&#231;in de &ouml;nemlidir. &Ouml;z&uuml;m Klima olarak ev, ofis, ma&#287;aza, restoran, otel ve ticari alanlarda ihtiyaca uygun iklimlendirme &#231;&ouml;z&uuml;mleri sunuyoruz. Amac&#305;m&#305;z; mek&acirc;n&#305;n yap&#305;s&#305;n&#305;, kullan&#305;m al&#305;&#351;kanl&#305;klar&#305;n&#305; ve uzun vadeli servis ihtiyac&#305;n&#305; birlikte de&#287;erlendirerek do&#287;ru sistemi kurmakt&#305;r.</p>
    <p>Antalya merkez, D&ouml;&#351;emealt&#305;, Kepez, Muratpa&#351;a ve Konyaalt&#305; gibi farkl&#305; b&ouml;lgelerde konut ve i&#351;letme ihtiya&#231;lar&#305; birbirinden ayr&#305;labilir. Bir evde sessizlik ve d&uuml;&#351;&uuml;k elektrik t&uuml;ketimi daha &ouml;nemliyken, bir i&#351;letmede s&uuml;rekli kullan&#305;m, servis eri&#351;imi, kapasite planlamas&#305; ve estetik montaj daha belirleyici olabilir. Bu nedenle hizmeti tek bir cihaz sat&#305;&#351;&#305; gibi de&#287;il; ke&#351;if, proje, montaj, bak&#305;m ve teknik destek ad&#305;mlar&#305;ndan olu&#351;an b&uuml;t&uuml;nl&uuml;kl&uuml; bir s&uuml;re&#231; olarak ele al&#305;yoruz.</p>
    <p>Klima al&#305;rken en s&#305;k yap&#305;lan hata sadece metrekareye bakarak cihaz se&#231;mektir. Oysa cephe y&ouml;n&uuml;, g&uuml;ne&#351; alma s&uuml;resi, tavan y&uuml;ksekli&#287;i, cam y&uuml;zeyi, ki&#351;i say&#305;s&#305; ve yal&#305;t&#305;m durumu cihaz performans&#305;n&#305; do&#287;rudan etkiler. Bu nedenle hizmet s&uuml;recimiz <a href="kesfetteam.asp" title="Antalya klima ke&#351;if hizmeti">ke&#351;if</a> ile ba&#351;lar; kapasite, cihaz tipi ve montaj noktas&#305; netle&#351;tikten sonra kullan&#305;c&#305;ya en uygun alternatifler sunulur.</p>
    <p><a href="klima.asp" title="Antalya Mitsubishi Electric klima modelleri">Mitsubishi Electric klima</a> modelleri, bireysel kullan&#305;mda sessiz &#231;al&#305;&#351;ma, enerji tasarrufu ve konfor beklentisine cevap verirken; ticari alanlarda daha y&uuml;ksek kapasite ve d&uuml;zenli performans ihtiyac&#305; &ouml;ne &#231;&#305;kar. Tekli klima uygulamalar&#305;ndan ticari tip cihazlara kadar, sat&#305;&#351; ve montaj s&uuml;recinde cihaz&#305;n sadece bug&uuml;nk&uuml; ihtiyac&#305; de&#287;il, uzun s&uuml;reli kullan&#305;m maliyeti de dikkate al&#305;n&#305;r. B&ouml;ylece gere&#287;inden b&uuml;y&uuml;k cihazla fazla enerji harcamak veya yetersiz kapasiteyle konfor kayb&#305; ya&#351;amak yerine dengeli bir se&#231;im yap&#305;l&#305;r.</p>
    <p>Otel, ofis, klinik, restoran, ma&#287;aza ve birden fazla odas&#305; bulunan yap&#305;larda <a href="vrf_sistemler.asp" title="Antalya VRF sistemleri">VRF sistemleri</a> daha esnek bir &#231;&ouml;z&uuml;m sa&#287;lar. Bir d&#305;&#351; &uuml;niteyle farkl&#305; i&#231; &uuml;nitelerin y&ouml;netilebilmesi, hem tasar&#305;m hem de kullan&#305;m kolayl&#305;&#287;&#305; a&#231;&#305;s&#305;ndan avantaj olu&#351;turur. Do&#287;ru projelendirme, borulama, drenaj ve devreye alma ad&#305;mlar&#305; sistemin verimini belirledi&#287;i i&#231;in uygulama s&uuml;recinde teknik detaylara dikkat edilir.</p>
    <p>Klima montaj&#305;, bak&#305;m ve teknik servis hizmetlerinde hedefimiz cihaz&#305;n ilk g&uuml;nk&uuml; performans&#305;n&#305; korumas&#305;d&#305;r. Hatal&#305; montaj veya geciken bak&#305;m, so&#287;utma verimini d&uuml;&#351;&uuml;rebilir ve ar&#305;za riskini art&#305;rabilir. Bu nedenle montaj sonras&#305; kullan&#305;m bilgilendirmesi, periyodik bak&#305;m ve ihtiya&#231; halinde teknik m&uuml;dahale s&uuml;recini planl&#305; &#351;ekilde y&uuml;r&uuml;t&uuml;yoruz.</p>
    <p>Kapal&#305; alanlarda temiz hava sirk&uuml;lasyonu ve i&#231; hava kalitesi de konforun par&#231;as&#305;d&#305;r. Bu y&uuml;zden <a href="havalandirma.asp" title="Antalya havaland&#305;rma sistemleri">havaland&#305;rma sistemleri</a> ve &#305;s&#305; pompas&#305; &#231;&ouml;z&uuml;mleriyle ya&#351;am alan&#305;n&#305; b&uuml;t&uuml;n olarak ele al&#305;yoruz. Antalya'da klima servisi, klima montaj&#305;, VRF sistemleri veya havaland&#305;rma hizmeti i&#231;in ihtiyac&#305;n&#305;za uygun &#231;&ouml;z&uuml;m&uuml; birlikte belirleyebiliriz.</p>
  </div>
  <div class="hizmetAksiyon"><a href="iletisim.asp" title="&Ouml;z&uuml;m Klima ileti&#351;im">Ke&#351;if ve servis talebi i&#231;in bizimle ileti&#351;ime ge&#231;in</a>.</div>
</section>

    


<section class="musteriYorumlari" aria-labelledby="musteri-yorumlari-baslik">
  <div class="musteriYorumlariBaslik">
    <span class="musteriYorumlariEtiket">Gerçek deneyimler</span>
    <h2 id="musteri-yorumlari-baslik">Müşterilerimiz ne söylüyor?</h2>
    <p>Antalya'da keşif, montaj ve teknik servis süreçlerimiz hakkında müşterilerimizin görüşleri.</p>
  </div>

  <div class="yorumSlider" id="yorumSlider">
<%
Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select top 20 * from yorumlar where onay=1"
Rs.Open rSQL, baglanti, 1, 3

Dim yorumSayisi
yorumSayisi = Rs.RecordCount

If yorumSayisi > 0 Then
  Do While Not Rs.Eof
%>
    <article class="yorumKart">
      <div class="yorumIkon" aria-hidden="true">
        <i class="fas fa-quote-left"></i>
      </div>
      <blockquote class="yorumMetni"><%=Rs("yorum") %></blockquote>
      <span class="yorumAltBilgi">Özüm Klima müşterisi</span>
    </article>
<%
    Rs.MoveNext
  Loop
Else
%>
    <div class="yorumBos">Henüz yayınlanmış müşteri yorumu bulunmuyor.</div>
<%
End If
%>

<% If yorumSayisi > 1 Then %>
    <button class="yorumYon yorumOnceki" type="button" aria-label="Önceki yorumu göster" onclick="yorumDegistir(-1)">
      <i class="fas fa-chevron-left" aria-hidden="true"></i>
    </button>
    <button class="yorumYon yorumSonraki" type="button" aria-label="Sonraki yorumu göster" onclick="yorumDegistir(1)">
      <i class="fas fa-chevron-right" aria-hidden="true"></i>
    </button>
<% End If %>
  </div>

<% If yorumSayisi > 1 Then %>
  <div class="yorumNoktalar" aria-label="Müşteri yorumu seçimi">
    <% For i = 1 To yorumSayisi %>
      <button class="yorumNokta" type="button" aria-label="<%=i%>. yorumu göster" onclick="yorumSec(<%=i - 1%>)"></button>
    <% Next %>
  </div>
<% End If %>

  <div class="yorumAltAlan">
    <div class="yorumAltMetin">
      <strong>Hizmet deneyiminizi paylaşın</strong>
      <span>Yorumlarınız hizmet kalitemizi geliştirmemize ve yeni müşterilerimizin doğru karar vermesine yardımcı olur.</span>
    </div>

    <div class="yorumAksiyonlar">
      <a class="yorumButon" href="kurumsal.asp?id=20">Tüm yorumlar ve değerlendirme</a>
      <a class="yorumSosyal" href="<%=ayars("face")%>" target="_blank" rel="noopener" aria-label="Facebook">
        <i class="fab fa-facebook-f" aria-hidden="true"></i>
      </a>
      <a class="yorumSosyal" href="<%=ayars("insta")%>" target="_blank" rel="noopener" aria-label="Instagram">
        <i class="fab fa-instagram" aria-hidden="true"></i>
      </a>
    </div>
  </div>
</section>

<script>
(function () {
  var aktifYorum = 0;
  var yorumlar = document.getElementsByClassName("yorumKart");
  var noktalar = document.getElementsByClassName("yorumNokta");
  var otomatikGecis;

  function yorumGoster(index) {
    var i;

    if (!yorumlar.length) {
      return;
    }

    if (index >= yorumlar.length) {
      aktifYorum = 0;
    } else if (index < 0) {
      aktifYorum = yorumlar.length - 1;
    } else {
      aktifYorum = index;
    }

    for (i = 0; i < yorumlar.length; i++) {
      yorumlar[i].className = yorumlar[i].className.replace(" yorumKartAktif", "");
    }

    for (i = 0; i < noktalar.length; i++) {
      noktalar[i].className = noktalar[i].className.replace(" yorumNoktaAktif", "");
      noktalar[i].setAttribute("aria-current", "false");
    }

    yorumlar[aktifYorum].className += " yorumKartAktif";

    if (noktalar.length && noktalar[aktifYorum]) {
      noktalar[aktifYorum].className += " yorumNoktaAktif";
      noktalar[aktifYorum].setAttribute("aria-current", "true");
    }
  }

  function otomatikBaslat() {
    if (yorumlar.length > 1) {
      otomatikGecis = window.setInterval(function () {
        yorumGoster(aktifYorum + 1);
      }, 9000);
    }
  }

  window.yorumDegistir = function (yon) {
    window.clearInterval(otomatikGecis);
    yorumGoster(aktifYorum + yon);
    otomatikBaslat();
  };

  window.yorumSec = function (index) {
    window.clearInterval(otomatikGecis);
    yorumGoster(index);
    otomatikBaslat();
  };

  yorumGoster(0);
  otomatikBaslat();
})();
</script>



<section id="uygulama-galeri-carousel" class="flexslider-container ozumCarousel ozumGaleriCarousel" aria-label="Özüm Klima uygulama galerisi">   
    <div class="flexslider ozumGaleriSlider">
        <ul class="slides" >  
     
                      
<%
set object = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From galery order by sira asc"
object.open Sorgula,baglanti,1,3   

Do while not object.Eof                                                                                                                                                                                                                  
%>   
<li  class="bizim" style="background-image: url('galeri/<%=object("foto")%>')"  onClick="window.location='galeri.asp';" ></li>  

<%  
object.MoveNExt
Loop
%>

                      
  </ul>
    </div>   
          
          
          

        
</section>

<section class="referansBolumu" aria-labelledby="referanslar-baslik">
  <header class="referansBaslik">
    <span class="referansEtiket">Kurumsal deneyim</span>
    <h2 id="referanslar-baslik">İklimlendirme çözümlerimizde bizi tercih eden işletmeler</h2>
    <p>Antalya'da konut, ofis ve ticari projelerde klima, VRF, havalandırma, montaj ve teknik servis hizmetleri sunduğumuz referanslarımızdan bazıları.</p>
  </header>

<%
Set icerisi = Server.CreateObject("ADODB.Recordset")
Sorgu = "Select * From refers"
icerisi.Open Sorgu, baglanti, 1, 3

Dim referansSayisi
referansSayisi = icerisi.RecordCount

If referansSayisi > 0 Then
%>
  <div class="referansAkis" role="region" aria-label="Özüm Klima referansları">
    <div class="referansTrack">

      <div class="referansGrup">
<%
  Do While Not icerisi.Eof
%>
       <%
Dim referansIsim, referansLink

referansIsim = Trim(CStr(icerisi("isim") & ""))
referansLink = Trim(CStr(icerisi("link") & ""))

If referansIsim = "" Then
  referansIsim = "Özüm Klima referans müşterisi"
End If

If referansLink <> "" Then
  If LCase(Left(referansLink, 7)) <> "http://" And LCase(Left(referansLink, 8)) <> "https://" Then
    referansLink = "http://" & referansLink
  End If
Else
  referansLink = "referanslar.asp"
End If
%>

<a class="referansLogo"
   href="<%=Server.HTMLEncode(referansLink)%>"
   title="<%=Server.HTMLEncode(referansIsim)%> referans çalışması"
   <% If referansLink <> "referanslar.asp" Then %>target="_blank" rel="noopener noreferrer"<% End If %>>

  <img src="ref/<%=Server.HTMLEncode(icerisi("resim"))%>"
       alt="<%=Server.HTMLEncode(referansIsim)%> - Özüm Klima iklimlendirme referansı"
       loading="lazy"
       decoding="async">
</a>
<%
    icerisi.MoveNext
  Loop
%>
      </div>

      <div class="referansGrup" aria-hidden="true">
<%
  icerisi.MoveFirst
  Do While Not icerisi.Eof
%>
        <span class="referansLogo">
          <img src="ref/<%=icerisi("resim") %>"
               alt=""
               loading="lazy"
               decoding="async">
        </span>
<%
    icerisi.MoveNext
  Loop
%>
      </div>

    </div>
  </div>
<%
Else
%>
  <div class="referansBos">Referanslarımız kısa süre içinde bu alanda yayınlanacaktır.</div>
<%
End If

icerisi.Close
Set icerisi = Nothing
%>

  <div class="referansAlt">
    <a href="referanslar.asp" title="Özüm Klima referans projeleri">
      Tüm referanslarımızı ve uygulama alanlarımızı inceleyin &raquo;
    </a>
  </div>
</section>


<DIV    class="urunkutular" >
       
                        
                        <Div class="solSutun">
                                 <a href="klima.asp?t=36" title="Antalya Mitsubishi Klima">
                                 <div style=" width:100%;max-widtH:411px ;  height:345px;background-image:url('images/evofis.jpg') " >
                                      <div class="Absolute-Center "><span>Ev Ofis Klimalar&#305;</span></div>
                                 </div>
                                 </a>   
                                 
                                 
                                 <a href="klima.asp?t=37" title="Antalya ticari klima">
                                 <div  style=" width:100%;max-widtH:411px ;  height:345px; background-image:url('images/ticari.jpg')  " >
                                      <div class="Absolute-Center "><span>Ticari Tip Klimalar</span></div>
                                 </div>
                                 </a>  
                                 
                                 
                                 <a href="vrf_sistemler.asp"  class="vrf">
                                 <div  style=" width:100%;max-widtH:842px ;  height:345px; background-image:url('images/vrf.jpg') " >
                                      <div class="Absolute-Center "><span>VRF Merkezi Sistemler</span></div>
                                 </div>
                                 </a>     
                   
                                 
                        
                        </Div>
              
                        
                         <Div  class="sagSutun" >
                                                                  
         <a href="vrf_sistemler.asp"  class="vrfm">
                                 <div  style=" width:100%;max-widtH:842px ;  height:345px; background-image:url('images/vrf.jpg') " >
                                      <div class="Absolute-Center "><span>VRF Merkezi Sistemler</span></div>
                                 </div>
                                 </a>                              
                         
                           <a href="havalandirma.asp">
                                 <div  style=" width:100%;max-widtH:410px ;  height:706px; background-image:url('images/hava.jpg')" >
                                      <div class="Absolute-Center "><span>Havaland&#305;rma Sistemleri</span></div>
                                 </div>
                                 </a>   
                                 
                </Div>     
                         
                         
                         
       
                              
</DIV> 


      
      
    <!--#include file='alt.asp'-->      

   

</body>
</html>