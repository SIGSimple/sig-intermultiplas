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
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_cpf_STRING
  MM_editTable = "tb_responsavel"
  MM_editColumn = "cod_fiscal"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
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
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
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
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
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
Dim rs_resp__MMColParam
rs_resp__MMColParam = "1"
If (Request.QueryString("cod_fiscal") <> "") Then 
  rs_resp__MMColParam = Request.QueryString("cod_fiscal")
End If
%>
<%
Dim rs_resp
Dim rs_resp_numRows

Set rs_resp = Server.CreateObject("ADODB.Recordset")
rs_resp.ActiveConnection = MM_cpf_STRING
rs_resp.Source = "SELECT * FROM tb_responsavel WHERE cod_fiscal = " + Replace(rs_resp__MMColParam, "'", "''") + " ORDER BY Responsável ASC"
rs_resp.CursorType = 0
rs_resp.CursorLocation = 2
rs_resp.LockType = 1
rs_resp.Open()

rs_resp_numRows = 0
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
<%
Dim rs_usuario
Dim rs_usuario_numRows

Set rs_usuario = Server.CreateObject("ADODB.Recordset")
rs_usuario.ActiveConnection = MM_cpf_STRING
rs_usuario.Source = "SELECT idusuario, nome FROM login ORDER BY nome ASC"
rs_usuario.CursorType = 0
rs_usuario.CursorLocation = 2
rs_usuario.LockType = 1
rs_usuario.Open()

rs_usuario_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style17 {	font-family: Arial, Helvetica, sans-serif;
	font-size: 16px;
}
.style44 {font-family: Arial, Helvetica, sans-serif; font-size: 9; font-weight: bold; }
.style45 {font-size: 9}
.style22 {font-family: Arial, Helvetica, sans-serif; font-size: 9; }
.style7 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
-->
</style>
</head>

<body>
<p align="center"><strong><span class="style17">Altera&ccedil;&atilde;o de Fiscais </span></strong></p>
<form method="post" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style7"><span class="style44">Nome:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="Responsvel" value="<%=(rs_resp.Fields.Item("Responsável").Value)%>" size="32">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style7"><span class="style44">Cargo:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="cargo" value="<%=(rs_resp.Fields.Item("cargo").Value)%>" size="18">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style7"><span class="style44">Email:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="email" value="<%=(rs_resp.Fields.Item("email").Value)%>" size="22">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style7"><span class="style44">Telefone:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="telefone" value="<%=(rs_resp.Fields.Item("telefone").Value)%>" size="18">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style7"><span class="style44">Núm. CREA:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="numero_crea" value="<%=(rs_resp.Fields.Item("numero_crea").Value)%>" size="18">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style7"><span class="style10">Empresa:</span></td>
      <td bgcolor="#CCCCCC"><select name="cod_empresa" class="style5">
      <option value=""></option>
        <%
While (NOT rs_construtora.EOF)
  If Trim(rs_construtora.Fields.Item("Construtora").Value) <> "" Then
    Response.Write "      <OPTION value='" & (rs_construtora.Fields.Item("cod_construtora").Value) & "'"
    If Lcase(rs_construtora.Fields.Item("cod_construtora").Value) = Lcase(rs_resp.Fields.Item("cod_empresa").Value) then
      Response.Write "selected"
    End If
    Response.Write ">" & (rs_construtora.Fields.Item("Construtora").Value) & "</OPTION>"
  End If

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
      <td align="right" nowrap bgcolor="#CCCCCC" class="style7"><span class="style10">Login no Sistema:</span></td>
      <td bgcolor="#CCCCCC"><select name="cod_usuario" class="style5">
        <option value=""></option>
        <%
While (NOT rs_usuario.EOF)
  If Trim(rs_usuario.Fields.Item("nome").Value) <> "" Then
    Response.Write "      <OPTION value='" & (rs_usuario.Fields.Item("idusuario").Value) & "'"
    If Lcase(rs_usuario.Fields.Item("idusuario").Value) = Lcase(rs_resp.Fields.Item("cod_usuario").Value) then
      Response.Write "selected"
    End If
    Response.Write ">" & (rs_usuario.Fields.Item("nome").Value) & "</OPTION>"
  End If

  rs_usuario.MoveNext()
Wend
If (rs_usuario.CursorType > 0) Then
  rs_usuario.MoveFirst
Else
  rs_usuario.Requery
End If
%>
      </select>      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style7"><span class="style45"></span></td>
      <td bgcolor="#CCCCCC"><input type="submit" value="Salvar">      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= rs_resp.Fields.Item("cod_fiscal").Value %>">
</form>
<p>&nbsp;</p>
</body>
</html>
<%
rs_resp.Close()
Set rs_resp = Nothing
%>
