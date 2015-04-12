<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_cpf_STRING
  MM_editTable = "tb_Acompanhamento"
  MM_editColumn = "cod_acompanhamento"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "filtro_acomp.asp"

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
  
End If
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim rs_exc_acomp__MMColParam
rs_exc_acomp__MMColParam = "1"
If (Request.QueryString("cod_acompanhamento") <> "") Then 
  rs_exc_acomp__MMColParam = Request.QueryString("cod_acompanhamento")
End If
%>
<%
Dim rs_exc_acomp
Dim rs_exc_acomp_numRows

Set rs_exc_acomp = Server.CreateObject("ADODB.Recordset")
rs_exc_acomp.ActiveConnection = MM_cpf_STRING
rs_exc_acomp.Source = "SELECT * FROM tb_Acompanhamento WHERE cod_acompanhamento = " + Replace(rs_exc_acomp__MMColParam, "'", "''") + ""
rs_exc_acomp.CursorType = 0
rs_exc_acomp.CursorLocation = 2
rs_exc_acomp.LockType = 1
rs_exc_acomp.Open()

rs_exc_acomp_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rs_exc_acomp_numRows = rs_exc_acomp_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style3 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
.style5 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; }
-->
</style>
</head>

<body>
<form id="form1" name="form1" method="POST" action="<%=MM_editAction%>">
  <input type="hidden" name="MM_delete" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= rs_exc_acomp.Fields.Item("cod_acompanhamento").Value %>">
  <label>
  <input type="submit" name="Submit" value="Excluir" />
</label>
</form>
<p>&nbsp;</p>

<div align="center">
  <table border="1">
    <tr bgcolor="#CCCCCC">
      <td><span class="style5">cod_acompanhamento</span></td>
      <td><span class="style5">PI</span></td>
      <td><span class="style5">Data do Registro</span></td>
      <td><span class="style5">Registro</span></td>
      <td><span class="style5">Respons&aacute;vel</span></td>
      <td><span class="style5">Previs&atilde;o</span></td>
      <td><span class="style5">t&eacute;rmino contratual</span></td>
      <td><span class="style5">cod_fiscal</span></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rs_exc_acomp.EOF)) %>
      <tr bgcolor="#CCCCCC">
        <td><span class="style3"><%=(rs_exc_acomp.Fields.Item("cod_acompanhamento").Value)%></span></td>
        <td><span class="style3"><%=(rs_exc_acomp.Fields.Item("PI").Value)%></span></td>
        <td><span class="style3"><%=(rs_exc_acomp.Fields.Item("Data do Registro").Value)%></span></td>
        <td><span class="style3"><%=(rs_exc_acomp.Fields.Item("Registro").Value)%></span></td>
        <td><span class="style3"><%=(rs_exc_acomp.Fields.Item("Responsável").Value)%></span></td>
        <td><span class="style3"><%=(rs_exc_acomp.Fields.Item("Previsão").Value)%></span></td>
        <td><span class="style3"><%=(rs_exc_acomp.Fields.Item("término contratual").Value)%></span></td>
        <td><span class="style3"><%=(rs_exc_acomp.Fields.Item("cod_fiscal").Value)%></span></td>
      </tr>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_exc_acomp.MoveNext()
Wend
%>
  </table>
</div>
</body>
</html>
<%
rs_exc_acomp.Close()
Set rs_exc_acomp = Nothing
%>
