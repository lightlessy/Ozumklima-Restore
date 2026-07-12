' Özüm Klima read-only Access encoding audit
' Usage:
'   cscript //nologo tools\utf8-audit.vbs "C:\path\TRvd678577.mdb" > utf8-audit.csv
'
' This script DOES NOT update the database.
Option Explicit

Dim args, dbPath, conn, catalog, table, column, rs
Dim suspicious, value, sql, rowId, primaryKey, outputLine

Set args = WScript.Arguments
If args.Count < 1 Then
  WScript.Echo "Usage: cscript //nologo utf8-audit.vbs <database.mdb>"
  WScript.Quit 1
End If

dbPath = args(0)
Set conn = CreateObject("ADODB.Connection")

On Error Resume Next
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Persist Security Info=False;"
If Err.Number <> 0 Then
  Err.Clear
  conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath & ";Persist Security Info=False;"
End If
If Err.Number <> 0 Then
  WScript.Echo "ERROR,""Database could not be opened: " & CsvEscape(Err.Description) & """"
  WScript.Quit 2
End If
On Error GoTo 0

Set catalog = CreateObject("ADOX.Catalog")
Set catalog.ActiveConnection = conn

WScript.Echo "table,column,row_key,pattern,value"

For Each table In catalog.Tables
  If table.Type = "TABLE" Then
    primaryKey = FindPrimaryKey(table)

    For Each column In table.Columns
      If IsTextColumn(column.Type) Then
        sql = "SELECT * FROM [" & Replace(table.Name, "]", "]]" ) & "]"
        Set rs = CreateObject("ADODB.Recordset")

        On Error Resume Next
        rs.Open sql, conn, 0, 1
        If Err.Number = 0 Then
          Do Until rs.EOF
            value = ""
            If Not IsNull(rs.Fields(column.Name).Value) Then
              value = CStr(rs.Fields(column.Name).Value)
            End If

            suspicious = DetectPattern(value)
            If suspicious <> "" Then
              rowId = GetRowKey(rs, primaryKey)
              outputLine = CsvEscape(table.Name) & "," & _
                           CsvEscape(column.Name) & "," & _
                           CsvEscape(rowId) & "," & _
                           CsvEscape(suspicious) & "," & _
                           CsvEscape(value)
              WScript.Echo outputLine
            End If
            rs.MoveNext
          Loop
        Else
          WScript.Echo CsvEscape(table.Name) & ",,,READ_ERROR," & CsvEscape(Err.Description)
          Err.Clear
        End If
        On Error GoTo 0

        If rs.State = 1 Then rs.Close
        Set rs = Nothing
      End If
    Next
  End If
Next

conn.Close
Set catalog = Nothing
Set conn = Nothing

Function IsTextColumn(ByVal adoType)
  ' 8 = adBSTR, 129 = adChar, 130 = adWChar,
  ' 200 = adVarChar, 201 = adLongVarChar,
  ' 202 = adVarWChar, 203 = adLongVarWChar
  Select Case CLng(adoType)
    Case 8, 129, 130, 200, 201, 202, 203
      IsTextColumn = True
    Case Else
      IsTextColumn = False
  End Select
End Function

Function FindPrimaryKey(ByVal oneTable)
  Dim oneKey, oneColumn
  FindPrimaryKey = ""
  On Error Resume Next
  For Each oneKey In oneTable.Keys
    If oneKey.Type = 1 Then ' adKeyPrimary
      For Each oneColumn In oneKey.Columns
        FindPrimaryKey = oneColumn.Name
        Exit Function
      Next
    End If
  Next
  On Error GoTo 0
End Function

Function GetRowKey(ByVal oneRs, ByVal keyName)
  On Error Resume Next
  If keyName <> "" Then
    If Not IsNull(oneRs.Fields(keyName).Value) Then
      GetRowKey = CStr(oneRs.Fields(keyName).Value)
      If Err.Number = 0 Then Exit Function
    End If
  End If
  Err.Clear
  GetRowKey = "(no primary key)"
  On Error GoTo 0
End Function

Function DetectPattern(ByVal textValue)
  Dim patterns, labels, i
  patterns = Array("ý", "þ", "ð", "Ý", "Þ", "Ð", "Ã", "Â", "Ä", "Å", "â€", "&#")
  labels = Array("legacy-tr-mojibake", "legacy-tr-mojibake", "legacy-tr-mojibake", _
                 "legacy-tr-mojibake", "legacy-tr-mojibake", "legacy-tr-mojibake", _
                 "utf8-mojibake", "utf8-mojibake", "utf8-mojibake", "utf8-mojibake", _
                 "smart-punctuation-mojibake", "stored-html-entity")

  DetectPattern = ""
  For i = 0 To UBound(patterns)
    If InStr(1, textValue, patterns(i), vbBinaryCompare) > 0 Then
      If DetectPattern <> "" Then DetectPattern = DetectPattern & "+"
      DetectPattern = DetectPattern & labels(i)
    End If
  Next
End Function

Function CsvEscape(ByVal textValue)
  Dim safeValue
  safeValue = CStr(textValue & "")
  safeValue = Replace(safeValue, vbCrLf, " ")
  safeValue = Replace(safeValue, vbCr, " ")
  safeValue = Replace(safeValue, vbLf, " ")
  safeValue = Replace(safeValue, Chr(34), Chr(34) & Chr(34))
  CsvEscape = Chr(34) & safeValue & Chr(34)
End Function
