<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="1,2"
MM_authFailedURL="erro.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (false Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_cpf_STRING
  MM_editTable = "tb_responsavel"
  MM_editRedirectUrl = "cad_resp.asp"
  MM_fieldsStr  = "Responsvel|value|cargo|value|email|value|telefone|value|cod_empresa|value|numero_crea|value"
  MM_columnsStr = "Responsável|',none,''|cargo|',none,''|email|',none,''|telefone|',none,''|cod_empresa|',none,''|numero_crea|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

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
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
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
Dim rs_resp
Dim rs_resp_numRows

Set rs_resp = Server.CreateObject("ADODB.Recordset")
rs_resp.ActiveConnection = MM_cpf_STRING
rs_resp.Source = "SELECT tb_responsavel.*, tb_Construtora.* FROM tb_responsavel INNER JOIN tb_Construtora ON tb_responsavel.cod_empresa = tb_Construtora.cod_construtora ORDER BY [tb_responsavel].[Responsável] ASC"
rs_resp.CursorType = 0
rs_resp.CursorLocation = 2
rs_resp.LockType = 1
rs_resp.Open()

rs_resp_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 20
Repeat1__index = 0
rs_resp_numRows = rs_resp_numRows + Repeat1__numRows
%>
<%
Dim rs_construtora
Dim rs_construtora_numRows

Set rs_construtora = Server.CreateObject("ADODB.Recordset")
rs_construtora.ActiveConnection = MM_cpf_STRING
rs_construtora.Source = "SELECT cod_construtora, Construtora FROM tb_Construtora ORDER BY Construtora ASC"
rs_construtora.CursorType = 0
rs_construtora.CursorLocation = 2
rs_construtora.LockType = 1
rs_construtora.Open()

rs_construtora_numRows = 0
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style5 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; }
.style7 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
.style17 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 16px;
}
.style22 {font-family: Arial, Helvetica, sans-serif; font-size: 9; }
.style23 {font-size: 9}
-->
</style>
</head>

<body>
<p align="center"><strong><span class="style17">Cadastro de Interessados </span></strong></p>
<form method="post" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style7"><span class="style22">Nome:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="Responsvel" value="" size="32">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style7"><span class="style22">Cargo:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="cargo" value="" size="18">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style7"><span class="style22">Email:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="email" value="" size="22">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style7"><span class="style22">Telefone:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="telefone" value="" size="18">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style7"><span class="style22">Núm. CREA:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="numero_crea" value="" size="18">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style7"><span class="style10">Empresa:</span></td>
      <td bgcolor="#CCCCCC">
	<select name="cod_empresa" class="style5">
	<option value=""></value>
        <%
While (NOT rs_construtora.EOF)
%>
        <option value="<%=(rs_construtora.Fields.Item("cod_construtora").Value)%>"><%=(rs_construtora.Fields.Item("Construtora").Value)%></option>
        <%
  rs_construtora.MoveNext()
Wend
If (rs_construtora.CursorType > 0) Then
  rs_construtora.MoveFirst
Else
  rs_construtora.Requery
End If
%>
      </select>      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style7"><span class="style23"></span></td>
      <td bgcolor="#CCCCCC"><input type="submit" value="Salvar">      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1">
</form>
<p>&nbsp;</p>

<div align="center">
  <table border="0">
    <tr bgcolor="#999999">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><span class="style7">Respons&aacute;vel</span></td>
      <td><span class="style7">Cargo</span></td>
      <td><span class="style7">Email</span></td>
      <td><span class="style7">Telefone</span></td>
      <td><span class="style7">Núm. CREA</span></td>
      <td><span class="style7">Empresa</span></td>
    </tr>
    <% While (NOT rs_resp.EOF) %>
      <tr bgcolor="#CCCCCC">
        <td><a href="altera_resp.asp?cod_fiscal=<%=(rs_resp.Fields.Item("cod_fiscal").Value)%>"><img src="const/imagens/edit.gif" width="16" height="15" border="0" /></a></td>
        <td><a href="del_resp.asp?cod_fiscal=<%=(rs_resp.Fields.Item("cod_fiscal").Value)%>"><img src="const/imagens/delete.gif" width="16" height="15" border="0" /></a></td>
        <td><span class="style5"><%=(rs_resp.Fields.Item("Responsável").Value)%></span></td>
        <td><span class="style5"><%=(rs_resp.Fields.Item("cargo").Value)%></span></td>
        <td><span class="style5"><%=(rs_resp.Fields.Item("email").Value)%></span></td>
        <td><span class="style5"><%=(rs_resp.Fields.Item("telefone").Value)%></span></td>
        <td><span class="style5"><%=(rs_resp.Fields.Item("numero_crea").Value)%></span></td>
        <td><span class="style5"><%=(rs_resp.Fields.Item("Construtora").Value)%></span></td>
      </tr>
      <% 
  rs_resp.MoveNext()
Wend
%>
  </table>
</div>
</body>
</html>
<%
rs_resp.Close()
Set rs_resp = Nothing
%>
