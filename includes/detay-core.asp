<%
Response.Buffer = True
Response.CodePage = 65001
Session.CodePage = 65001
Session.LCID = 1055
Response.Charset = "utf-8"

Const DP_AD_CMD_TEXT = 1
Const DP_AD_INTEGER = 3
Const DP_AD_PARAM_INPUT = 1

Function DpText(ByVal value)
    If IsNull(value) Or IsEmpty(value) Then
        DpText = ""
    Else
        DpText = CStr(value)
    End If
End Function

Function DpField(ByVal recordset, ByVal fieldName, ByVal fallbackValue)
    Dim fieldValue

    fieldValue = fallbackValue
    On Error Resume Next
    fieldValue = recordset.Fields(fieldName).Value
    If Err.Number <> 0 Or IsNull(fieldValue) Or IsEmpty(fieldValue) Then
        Err.Clear
        fieldValue = fallbackValue
    End If
    On Error GoTo 0

    DpField = fieldValue
End Function

Function DpNumber(ByVal value)
    If IsNumeric(value) Then
        DpNumber = CDbl(value)
    Else
        DpNumber = 0
    End If
End Function

Function DpPositiveLong(ByVal value)
    DpPositiveLong = 0
    If IsNumeric(value) Then
        If CDbl(value) > 0 Then DpPositiveLong = CLng(value)
    End If
End Function

Function DpPlainText(ByVal value)
    Dim textValue, tagStart, tagEnd

    textValue = DpText(value)
    textValue = Replace(textValue, vbCrLf, " ")
    textValue = Replace(textValue, vbCr, " ")
    textValue = Replace(textValue, vbLf, " ")
    textValue = Replace(textValue, "&nbsp;", " ", 1, -1, 1)
    textValue = Replace(textValue, "&amp;", "&", 1, -1, 1)
    textValue = Replace(textValue, "&quot;", Chr(34), 1, -1, 1)
    textValue = Replace(textValue, "&#39;", "'", 1, -1, 1)

    Do While InStr(textValue, "<") > 0
        tagStart = InStr(textValue, "<")
        tagEnd = InStr(tagStart, textValue, ">")
        If tagEnd <= tagStart Then Exit Do
        textValue = Left(textValue, tagStart - 1) & " " & Mid(textValue, tagEnd + 1)
    Loop

    textValue = Trim(textValue)
    Do While InStr(textValue, "  ") > 0
        textValue = Replace(textValue, "  ", " ")
    Loop

    DpPlainText = textValue
End Function

Function DpJsonText(ByVal value)
    Dim textValue

    textValue = DpText(value)
    textValue = Replace(textValue, "\", "\\")
    textValue = Replace(textValue, Chr(34), "\" & Chr(34))
    textValue = Replace(textValue, vbCrLf, "\n")
    textValue = Replace(textValue, vbCr, "\n")
    textValue = Replace(textValue, vbLf, "\n")
    textValue = Replace(textValue, "<", "\u003c")
    textValue = Replace(textValue, ">", "\u003e")
    textValue = Replace(textValue, "&", "\u0026")

    DpJsonText = textValue
End Function

Function DpJsonNumber(ByVal value)
    DpJsonNumber = Replace(CStr(CDbl(value)), ",", ".")
End Function

Function DpProductSlug(ByVal value)
    Dim slugValue, slugRegex

    slugValue = LCase(Trim(DpText(value)))
    slugValue = Replace(slugValue, "ç", "c")
    slugValue = Replace(slugValue, "ğ", "g")
    slugValue = Replace(slugValue, "ı", "i")
    slugValue = Replace(slugValue, "ö", "o")
    slugValue = Replace(slugValue, "ş", "s")
    slugValue = Replace(slugValue, "ü", "u")
    slugValue = Replace(slugValue, "Ç", "c")
    slugValue = Replace(slugValue, "Ğ", "g")
    slugValue = Replace(slugValue, "İ", "i")
    slugValue = Replace(slugValue, "I", "i")
    slugValue = Replace(slugValue, "Ö", "o")
    slugValue = Replace(slugValue, "Ş", "s")
    slugValue = Replace(slugValue, "Ü", "u")
    slugValue = Replace(slugValue, "ý", "i")
    slugValue = Replace(slugValue, "ð", "g")
    slugValue = Replace(slugValue, "þ", "s")

    Set slugRegex = New RegExp
    slugRegex.Global = True
    slugRegex.IgnoreCase = True
    slugRegex.Pattern = "[^a-z0-9]+"
    slugValue = slugRegex.Replace(slugValue, "-")

    Do While Len(slugValue) > 0 And Left(slugValue, 1) = "-"
        slugValue = Mid(slugValue, 2)
    Loop
    Do While Len(slugValue) > 0 And Right(slugValue, 1) = "-"
        slugValue = Left(slugValue, Len(slugValue) - 1)
    Loop

    If slugValue = "" Then slugValue = "urun"
    DpProductSlug = slugValue
End Function

Function DpIsHexColor(ByVal value)
    Dim colorValue, colorRegex

    colorValue = Replace(Trim(DpText(value)), "#", "")
    Set colorRegex = New RegExp
    colorRegex.Global = False
    colorRegex.IgnoreCase = True
    colorRegex.Pattern = "^([0-9a-f]{3}|[0-9a-f]{6})$"
    DpIsHexColor = colorRegex.Test(colorValue)
End Function

Function DpExtractProductId()
    Dim candidate, rawQuery, requestUrl, matchRegex, matches

    DpExtractProductId = 0
    candidate = Trim(Request.QueryString("id") & "")

    If candidate <> "" And IsNumeric(candidate) Then
        DpExtractProductId = DpPositiveLong(candidate)
        Exit Function
    End If

    rawQuery = Trim(Request.ServerVariables("QUERY_STRING") & "")
    If rawQuery <> "" Then candidate = rawQuery

    If candidate = "" Then
        requestUrl = Trim(Request.ServerVariables("URL") & "")
        If InStrRev(requestUrl, "/") > 0 Then
            candidate = Mid(requestUrl, InStrRev(requestUrl, "/") + 1)
        Else
            candidate = requestUrl
        End If
    End If

    If InStr(candidate, "&") > 0 Then candidate = Left(candidate, InStr(candidate, "&") - 1)

    Set matchRegex = New RegExp
    matchRegex.Global = False
    matchRegex.IgnoreCase = True
    matchRegex.Pattern = "([0-9]+)$"
    Set matches = matchRegex.Execute(candidate)

    If matches.Count = 1 Then
        DpExtractProductId = DpPositiveLong(matches(0).SubMatches(0))
    End If
End Function

Function DpIntegerQuery(ByVal connection, ByVal sqlText, ByVal idValue, ByRef errorText)
    Dim command, resultSet

    Set DpIntegerQuery = Nothing
    errorText = ""

    On Error Resume Next
    Set command = Server.CreateObject("ADODB.Command")
    Set command.ActiveConnection = connection
    command.CommandType = DP_AD_CMD_TEXT
    command.CommandText = sqlText
    command.Parameters.Append command.CreateParameter("@id", DP_AD_INTEGER, DP_AD_PARAM_INPUT, , CLng(idValue))
    Set resultSet = command.Execute

    If Err.Number <> 0 Then
        errorText = CStr(Err.Number) & ": " & Err.Description
        Err.Clear
    Else
        Set DpIntegerQuery = resultSet
    End If

    Set command = Nothing
    On Error GoTo 0
End Function

Sub DpIncrementHit(ByVal connection, ByVal productId)
    Dim command, affectedRows

    On Error Resume Next
    affectedRows = 0

    Set command = Server.CreateObject("ADODB.Command")
    Set command.ActiveConnection = connection
    command.CommandType = DP_AD_CMD_TEXT
    command.CommandText = "UPDATE [products] SET [hit]=1 WHERE [AffiliateID]=? AND [hit] Is Null"
    command.Parameters.Append command.CreateParameter("@id", DP_AD_INTEGER, DP_AD_PARAM_INPUT, , CLng(productId))
    command.Execute affectedRows
    Set command = Nothing

    If Err.Number = 0 And affectedRows = 0 Then
        Set command = Server.CreateObject("ADODB.Command")
        Set command.ActiveConnection = connection
        command.CommandType = DP_AD_CMD_TEXT
        command.CommandText = "UPDATE [products] SET [hit]=[hit]+1 WHERE [AffiliateID]=?"
        command.Parameters.Append command.CreateParameter("@id", DP_AD_INTEGER, DP_AD_PARAM_INPUT, , CLng(productId))
        command.Execute
        Set command = Nothing
    End If

    Err.Clear
    On Error GoTo 0
End Sub

Sub DpCloseRecordset(ByRef recordset)
    On Error Resume Next
    If Not (recordset Is Nothing) Then
        If recordset.State = 1 Then recordset.Close
        Set recordset = Nothing
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Sub DpCloseConnection(ByRef connection)
    On Error Resume Next
    If Not (connection Is Nothing) Then
        If connection.State = 1 Then connection.Close
        Set connection = Nothing
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Dim dpConn, dpProductRs, relatedProducts
Dim dpProductId, dpQueryError, dpRelatedError, dpFatalError, productFound, hasRelatedProducts
Dim productId, productTypeId, productCategoryId, productSurfaceId
Dim categoryName, typeName, surfaceName
Dim productName, productCode, productDescription, displayDescription
Dim productNotes, productVideo, manufacturerFile, documentFile
Dim canonicalUrl, seoTitle, seoDescription, absoluteImageUrl
Dim foto(5), colorValues(10), sizeValues(10)
Dim schemaImages, imageIndex, imageCount
Dim productSchema, breadcrumbSchema, categoryPath
Dim regularPrice, discountPrice, effectivePrice, stockCount, deliveryDays
Dim hasColors, hasSizes, optionIndex, optionValue, cleanColor
Dim relatedPrice, relatedRegularPrice, relatedDiscountPrice, relatedDiscountPercent
Dim relatedName, relatedImage, relatedUrl, relatedId

Set dpConn = Nothing
Set dpProductRs = Nothing
Set relatedProducts = Nothing

dpProductId = DpExtractProductId()
dpFatalError = False
productFound = False
hasRelatedProducts = False
dpQueryError = ""
dpRelatedError = ""

On Error Resume Next
Set dpConn = Server.CreateObject("ADODB.Connection")
dpConn.Open "driver={microsoft access driver (*.mdb)}; dbq=" & Server.MapPath("db/TRvd678577.mdb")
If Err.Number <> 0 Then
    dpFatalError = True
    dpQueryError = CStr(Err.Number) & ": " & Err.Description
    Err.Clear
End If
On Error GoTo 0

If Not dpFatalError And dpProductId > 0 Then
    Set dpProductRs = DpIntegerQuery(dpConn, _
        "SELECT p.*, ag.isim AS category_name, tp.isim AS type_name, sf.isim AS surface_name " & _
        "FROM (([products] AS p LEFT JOIN [anagrup] AS ag ON p.anagrup=ag.id) " & _
        "LEFT JOIN [tip] AS tp ON p.tip=tp.id) " & _
        "LEFT JOIN [surface] AS sf ON p.surface=sf.id " & _
        "WHERE p.AffiliateID=? AND p.yayin=1", _
        dpProductId, dpQueryError)

    If dpProductRs Is Nothing Then
        dpFatalError = True
    ElseIf Not dpProductRs.EOF Then
        productFound = True
    End If
End If

If productFound Then
    productId = DpPositiveLong(DpField(dpProductRs, "AffiliateID", dpProductId))
    productTypeId = DpPositiveLong(DpField(dpProductRs, "tip", 0))
    productCategoryId = DpPositiveLong(DpField(dpProductRs, "anagrup", 0))
    productSurfaceId = DpPositiveLong(DpField(dpProductRs, "surface", 0))

    categoryName = Trim(DpText(DpField(dpProductRs, "category_name", "")))
    typeName = Trim(DpText(DpField(dpProductRs, "type_name", "")))
    surfaceName = Trim(DpText(DpField(dpProductRs, "surface_name", "")))

    productName = Trim(DpText(DpField(dpProductRs, "isim", "")))
    productCode = Trim(DpText(DpField(dpProductRs, "kodu", "")))
    productDescription = DpPlainText(DpField(dpProductRs, "descr", ""))
    productNotes = DpText(DpField(dpProductRs, "notlar", ""))
    productVideo = DpText(DpField(dpProductRs, "video", ""))
    manufacturerFile = Trim(DpText(DpField(dpProductRs, "uretici", "")))
    documentFile = Trim(DpText(DpField(dpProductRs, "dokuman", "")))

    displayDescription = productDescription
    If displayDescription = "" Then
        displayDescription = productName & " için Antalya'da ücretsiz keşif, profesyonel montaj ve teknik servis desteği alın."
    End If

    seoDescription = displayDescription
    If Len(seoDescription) > 157 Then seoDescription = Left(seoDescription, 154) & "..."

    seoTitle = productName & " | Özüm Klima"
    If Len(seoTitle) > 65 Then seoTitle = Left(productName, 47) & "... | Özüm Klima"

    canonicalUrl = "https://ozumklima.com/" & DpProductSlug(productName) & "-" & productId

    schemaImages = ""
    imageCount = 0
    For imageIndex = 1 To 5
        foto(imageIndex) = Trim(DpText(DpField(dpProductRs, "foto" & imageIndex, "")))
        If foto(imageIndex) <> "" Then
            absoluteImageUrl = "https://ozumklima.com/urunler/" & foto(imageIndex)
            imageCount = imageCount + 1
            If schemaImages <> "" Then schemaImages = schemaImages & ","
            schemaImages = schemaImages & Chr(34) & DpJsonText(absoluteImageUrl) & Chr(34)
        End If
    Next

    regularPrice = DpNumber(DpField(dpProductRs, "fiyat", 0))
    discountPrice = DpNumber(DpField(dpProductRs, "indirim", 0))
    stockCount = DpNumber(DpField(dpProductRs, "stok", 0))
    deliveryDays = DpNumber(DpField(dpProductRs, "delivery", 0))
    effectivePrice = regularPrice
    If discountPrice > 0 Then effectivePrice = discountPrice

    hasColors = False
    hasSizes = False
    For optionIndex = 1 To 10
        colorValues(optionIndex) = Trim(DpText(DpField(dpProductRs, "renk" & optionIndex, "")))
        If colorValues(optionIndex) <> "" And colorValues(optionIndex) <> "0" Then hasColors = True

        sizeValues(optionIndex) = Trim(DpText(DpField(dpProductRs, "beden" & optionIndex, "")))
        If sizeValues(optionIndex) <> "" And sizeValues(optionIndex) <> "0" Then hasSizes = True
    Next

    categoryPath = categoryName
    If typeName <> "" Then
        If categoryPath <> "" Then categoryPath = categoryPath & " > "
        categoryPath = categoryPath & typeName
    End If
    If surfaceName <> "" Then
        If categoryPath <> "" Then categoryPath = categoryPath & " > "
        categoryPath = categoryPath & surfaceName
    End If

    productSchema = "{" & _
        Chr(34) & "@context" & Chr(34) & ":" & Chr(34) & "https://schema.org" & Chr(34) & "," & _
        Chr(34) & "@type" & Chr(34) & ":" & Chr(34) & "Product" & Chr(34) & "," & _
        Chr(34) & "@id" & Chr(34) & ":" & Chr(34) & DpJsonText(canonicalUrl & "#product") & Chr(34) & "," & _
        Chr(34) & "url" & Chr(34) & ":" & Chr(34) & DpJsonText(canonicalUrl) & Chr(34) & "," & _
        Chr(34) & "name" & Chr(34) & ":" & Chr(34) & DpJsonText(productName) & Chr(34) & "," & _
        Chr(34) & "description" & Chr(34) & ":" & Chr(34) & DpJsonText(displayDescription) & Chr(34)

    If schemaImages <> "" Then
        productSchema = productSchema & "," & Chr(34) & "image" & Chr(34) & ":[" & schemaImages & "]"
    End If
    If productCode <> "" And productCode <> "0" Then
        productSchema = productSchema & "," & Chr(34) & "sku" & Chr(34) & ":" & Chr(34) & DpJsonText(productCode) & Chr(34)
    End If
    If categoryPath <> "" Then
        productSchema = productSchema & "," & Chr(34) & "category" & Chr(34) & ":" & Chr(34) & DpJsonText(categoryPath) & Chr(34)
    End If

    If effectivePrice > 0 Then
        productSchema = productSchema & "," & Chr(34) & "offers" & Chr(34) & ":{" & _
            Chr(34) & "@type" & Chr(34) & ":" & Chr(34) & "Offer" & Chr(34) & "," & _
            Chr(34) & "url" & Chr(34) & ":" & Chr(34) & DpJsonText(canonicalUrl) & Chr(34) & "," & _
            Chr(34) & "priceCurrency" & Chr(34) & ":" & Chr(34) & "TRY" & Chr(34) & "," & _
            Chr(34) & "price" & Chr(34) & ":" & Chr(34) & DpJsonNumber(effectivePrice) & Chr(34) & "," & _
            Chr(34) & "availability" & Chr(34) & ":" & Chr(34)

        If stockCount > 0 Then
            productSchema = productSchema & "https://schema.org/InStock"
        Else
            productSchema = productSchema & "https://schema.org/OutOfStock"
        End If

        productSchema = productSchema & Chr(34) & "," & _
            Chr(34) & "seller" & Chr(34) & ":{" & _
                Chr(34) & "@type" & Chr(34) & ":" & Chr(34) & "Organization" & Chr(34) & "," & _
                Chr(34) & "name" & Chr(34) & ":" & Chr(34) & "Özüm Klima" & Chr(34) & _
            "}}"
    End If
    productSchema = productSchema & "}"

    breadcrumbSchema = "{" & _
        Chr(34) & "@context" & Chr(34) & ":" & Chr(34) & "https://schema.org" & Chr(34) & "," & _
        Chr(34) & "@type" & Chr(34) & ":" & Chr(34) & "BreadcrumbList" & Chr(34) & "," & _
        Chr(34) & "itemListElement" & Chr(34) & ":[" & _
            "{" & Chr(34) & "@type" & Chr(34) & ":" & Chr(34) & "ListItem" & Chr(34) & "," & Chr(34) & "position" & Chr(34) & ":1," & Chr(34) & "name" & Chr(34) & ":" & Chr(34) & "Ana Sayfa" & Chr(34) & "," & Chr(34) & "item" & Chr(34) & ":" & Chr(34) & "https://ozumklima.com/" & Chr(34) & "}," & _
            "{" & Chr(34) & "@type" & Chr(34) & ":" & Chr(34) & "ListItem" & Chr(34) & "," & Chr(34) & "position" & Chr(34) & ":2," & Chr(34) & "name" & Chr(34) & ":" & Chr(34) & DpJsonText(typeName) & Chr(34) & "," & Chr(34) & "item" & Chr(34) & ":" & Chr(34) & "https://ozumklima.com/klima.asp?t=" & productTypeId & Chr(34) & "}," & _
            "{" & Chr(34) & "@type" & Chr(34) & ":" & Chr(34) & "ListItem" & Chr(34) & "," & Chr(34) & "position" & Chr(34) & ":3," & Chr(34) & "name" & Chr(34) & ":" & Chr(34) & DpJsonText(productName) & Chr(34) & "," & Chr(34) & "item" & Chr(34) & ":" & Chr(34) & DpJsonText(canonicalUrl) & Chr(34) & "}" & _
        "]}"

    DpIncrementHit dpConn, productId

    Set relatedProducts = DpIntegerQuery(dpConn, _
        "SELECT TOP 8 p.* FROM [related] AS r " & _
        "INNER JOIN [products] AS p ON r.ilgili=p.AffiliateID " & _
        "WHERE r.urun=? AND p.yayin=1 " & _
        "ORDER BY p.sira DESC, p.AffiliateID DESC", _
        productId, dpRelatedError)

    If Not (relatedProducts Is Nothing) Then
        If Not relatedProducts.EOF Then hasRelatedProducts = True
    End If
End If

DpCloseRecordset dpProductRs

If dpFatalError Then
    Response.Status = "500 Internal Server Error"
    On Error Resume Next
    Response.AppendToLog " detail.asp database error"
    On Error GoTo 0
ElseIf Not productFound Then
    Response.Status = "404 Not Found"
End If
%>
