<%
' Staging include for the UTF-8 migration.
' Add <%@ Language="VBScript" CodePage="65001" %> as the FIRST line
' of each ASP entry page before switching that page to this include.
Response.Buffer = True
Response.CodePage = 65001
Session.CodePage = 65001
Session.LCID = 1055
Response.Charset = "utf-8"
%>
