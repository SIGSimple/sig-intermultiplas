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
  MM_editTable = "login"
  MM_editRedirectUrl = "cad_usuario.asp"
  MM_fieldsStr  = "nome|value|senha|value|nivel|value"
  MM_columnsStr = "nome|',none,''|senha|',none,''|nivel|none,none,NULL"

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
Dim rs_lista
Dim rs_lista_numRows

Set rs_lista = Server.CreateObject("ADODB.Recordset")
rs_lista.ActiveConnection = MM_cpf_STRING
rs_lista.Source = "SELECT login.*, tb_nivel.desc_nivel  FROM login INNER JOIN tb_nivel ON login.nivel = tb_nivel.id_nivel;  "
rs_lista.CursorType = 0
rs_lista.CursorLocation = 2
rs_lista.LockType = 1
rs_lista.Open()

rs_lista_numRows = 0
%>
<%
Dim rs_cadusuario
Dim rs_cadusuario_numRows

Set rs_cadusuario = Server.CreateObject("ADODB.Recordset")
rs_cadusuario.ActiveConnection = MM_cpf_STRING
rs_cadusuario.Source = "SELECT * FROM login"
rs_cadusuario.CursorType = 0
rs_cadusuario.CursorLocation = 2
rs_cadusuario.LockType = 1
rs_cadusuario.Open()

rs_cadusuario_numRows = 0
%>
<%
Dim rs_nivel
Dim rs_nivel_numRows

Set rs_nivel = Server.CreateObject("ADODB.Recordset")
rs_nivel.ActiveConnection = MM_cpf_STRING
rs_nivel.Source = "SELECT * FROM tb_nivel"
rs_nivel.CursorType = 0
rs_nivel.CursorLocation = 2
rs_nivel.LockType = 1
rs_nivel.Open()

rs_nivel_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rs_lista_numRows = rs_lista_numRows + Repeat1__numRows
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Cadastro de Usu&aacute;rios</title>
<style type="text/css">
<!--
.style3 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; }
.style5 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
.style13 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
.style20 {font-family: Arial, Helvetica, sans-serif; font-size: 11px; font-weight: bold; }
.style21 {font-size: 11px}
-->
</style>
</head>

<body>
<form method="POST" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style20">Nome:</span></td>
      <td bgcolor="#CCCCCC"><input name="nome" type="text" value="" size="32">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style20">Senha:</span></td>
      <td bgcolor="#CCCCCC"><input name="senha" type="password" value="000000" size="15">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style20">Nivel de Acesso:</span></td>
      <td bgcolor="#CCCCCC"><span class="style13">
        <select name="nivel">
          <%
While (NOT rs_nivel.EOF)
%>
          <option value="<%=(rs_nivel.Fields.Item("id_nivel").Value)%>"><%=(rs_nivel.Fields.Item("desc_nivel").Value)%></option>
          <%
  rs_nivel.MoveNext()
Wend
If (rs_nivel.CursorType > 0) Then
  rs_nivel.MoveFirst
Else
  rs_nivel.Requery
End If
%>
        </select>
      </span> </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC">&nbsp;</td>
      <td bgcolor="#CCCCCC"><input type="submit" value="Salvar">      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1">
</form>
<div align="center">
  <table border="0">
    <tr bgcolor="#CCCCCC">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><span class="style5">idusuario</span></td>
      <td><span class="style5">Nome</span></td>
      <td><span class="style5">nivel</span></td>
      <td><span class="style5">N&iacute;vel de Acesso </span></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rs_lista.EOF)) %>
      <tr bgcolor="#CCCCCC">
        <td><a href="altera_usuario.asp?idusuario=<%=(rs_lista.Fields.Item("idusuario").Value)%>"><img src="const/imagens/edit.gif" width="16" height="15" border="0" /></a></td>
        <td><a href="delete_usuario.asp?idusuario=<%=(rs_lista.Fields.Item("idusuario").Value)%>"><img src="imagens/delete.gif" width="16" height="15" border="0" /></a></td>
        <td><span class="style3"><%=(rs_lista.Fields.Item("idusuario").Value)%></span></td>
        <td><span class="style3"><%=(rs_lista.Fields.Item("nome").Value)%></span></td>
        <td><span class="style3"><%=(rs_lista.Fields.Item("nivel").Value)%></span></td>
        <td><span class="style3"><%=(rs_lista.Fields.Item("desc_nivel").Value)%></span></td>
      </tr>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_lista.MoveNext()
Wend
%>
  </table>
</div>
</body>
</html>
<%
rs_lista.Close()
Set rs_lista = Nothing
%>
<%
rs_cadusuario.Close()
Set rs_cadusuario = Nothing
%>
<%
rs_nivel.Close()
Set rs_nivel = Nothing
%>
