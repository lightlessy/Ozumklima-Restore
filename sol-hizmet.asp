<%
Function CleanText(ByVal value)
    If IsNull(value) Then
        CleanText = ""
    Else
        CleanText = Trim(CStr(value))
    End If
End Function

Function HtmlTrEncode(ByVal value)
    Dim text
    text = Server.HTMLEncode(CleanText(value))

    text = Replace(text, ChrW(304), "&#304;")
    text = Replace(text, ChrW(305), "&#305;")
    text = Replace(text, ChrW(286), "&#286;")
    text = Replace(text, ChrW(287), "&#287;")
    text = Replace(text, ChrW(350), "&#350;")
    text = Replace(text, ChrW(351), "&#351;")
    text = Replace(text, ChrW(214), "&Ouml;")
    text = Replace(text, ChrW(246), "&ouml;")
    text = Replace(text, ChrW(220), "&Uuml;")
    text = Replace(text, ChrW(252), "&uuml;")
    text = Replace(text, ChrW(199), "&Ccedil;")
    text = Replace(text, ChrW(231), "&ccedil;")

    HtmlTrEncode = text
End Function

Function MenuEtiket(ByVal value)
    Dim text
    text = CleanText(value)

    text = Replace(text, "Mitsubishi Elektrik", "Mitsubishi Electric")
    text = Replace(text, "VRF Sistemler", "VRF Sistemleri")
    text = Replace(text, "Sat" & ChrW(305) & ChrW(351) & " Montaj", "Sat" & ChrW(305) & ChrW(351) & " ve Montaj")
    text = Replace(text, "Satis Montaj", "Sat" & ChrW(305) & ChrW(351) & " ve Montaj")

    MenuEtiket = text
End Function
%>

<nav class="solMenu" aria-label="Hizmet men&uuml;s&uuml;">

    <span class="solMenuBas">H&#304;ZMETLER</span>

    <%
    Dim rsSayfalar, sqlSayfalar, menuIsim, menuLink

    Set rsSayfalar = Server.CreateObject("ADODB.Recordset")
    sqlSayfalar = "SELECT id, isim FROM sayfalar WHERE ana=2 ORDER BY sira ASC"
    rsSayfalar.Open sqlSayfalar, baglanti, 1, 1

    Do While Not rsSayfalar.EOF
        menuIsim = MenuEtiket(rsSayfalar("isim"))
        menuLink = "h.asp?h=" & Server.URLEncode(cevir(rsSayfalar("isim"))) & "&amp;id=" & rsSayfalar("id")
    %>

        <a href="<%=menuLink%>">
            <div><i class="fas fa-chevron-right" aria-hidden="true"></i> <%=HtmlTrEncode(menuIsim)%></div>
        </a>

    <%
        rsSayfalar.MoveNext
    Loop

    rsSayfalar.Close
    Set rsSayfalar = Nothing
    %>

    <span class="solMenuBas" style="margin-top:18px; display:block;">ANTALYA KL&#304;MA H&#304;ZMETLER&#304;</span>

    <a href="antalya-klima-servisi.asp">
        <div><i class="fas fa-chevron-right" aria-hidden="true"></i> Antalya Klima Servisi</div>
    </a>

    <a href="antalya-klima-montaji.asp">
        <div><i class="fas fa-chevron-right" aria-hidden="true"></i> Antalya Klima Montaj&#305;</div>
    </a>

    <a href="antalya-klima-bakimi.asp">
        <div><i class="fas fa-chevron-right" aria-hidden="true"></i> Antalya Klima Bak&#305;m&#305;</div>
    </a>

    <a href="antalya-mitsubishi-electric-klima.asp">
        <div><i class="fas fa-chevron-right" aria-hidden="true"></i> Antalya Mitsubishi Electric Klima</div>
    </a>

    <a href="antalya-vrf-sistemleri.asp">
        <div><i class="fas fa-chevron-right" aria-hidden="true"></i> Antalya VRF Sistemleri</div>
    </a>

    <a href="dosemealti-klima-servisi.asp">
        <div><i class="fas fa-chevron-right" aria-hidden="true"></i> D&ouml;&#351;emealt&#305; Klima Servisi</div>
    </a>

    <span class="solMenuBas" style="margin-top:18px; display:block;">KL&#304;MA REHBER&#304;</span>

    <a href="klima-rehberi.asp">
        <div><i class="fas fa-chevron-right" aria-hidden="true"></i> Klima Rehberi</div>
    </a>

</nav>