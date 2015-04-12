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
  MM_editTable = "login"
  MM_editColumn = "idusuario"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "cad_usuario.asp"
  MM_fieldsStr  = "nome|value|nivel|value|senha|value"
  MM_columnsStr = "nome|',none,''|nivel|none,none,NULL|senha|',none,''"

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
Dim rs_altera_usuario__MMColParam
rs_altera_usuario__MMColParam = "1"
If (Request.QueryString("idusuario") <> "") Then 
  rs_altera_usuario__MMColParam = Request.QueryString("idusuario")
End If
%>
<%
Dim rs_altera_usuario
Dim rs_altera_usuario_numRows

Set rs_altera_usuario = Server.CreateObject("ADODB.Recordset")
rs_altera_usuario.ActiveConnection = MM_cpf_STRING
rs_altera_usuario.Source = "SELECT * FROM login WHERE idusuario = " + Replace(rs_altera_usuario__MMColParam, "'", "''") + ""
rs_altera_usuario.CursorType = 0
rs_altera_usuario.CursorLocation = 2
rs_altera_usuario.LockType = 1
rs_altera_usuario.Open()

rs_altera_usuario_numRows = 0
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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style5 {font-family: Arial, Helvetica, sans-serif; font-size: 12; }
.style6 {font-size: 12}
-->
</style>
</head>

<body>
<form method="post" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style5">Nome:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="nome" value="<%=(rs_altera_usuario.Fields.Item("nome").Value)%>" size="32">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style5">Nivel:</span></td>
      <td bgcolor="#CCCCCC"><select name="nivel">
        <%
While (NOT rs_nivel.EOF)
%>
        <option value="<%=(rs_nivel.Fields.Item("id_nivel").Value)%>" <%If (Not isNull((rs_altera_usuario.Fields.Item("nivel").Value))) Then If (CStr(rs_nivel.Fields.Item("id_nivel").Value) = CStr((rs_altera_usuario.Fields.Item("nivel").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rs_nivel.Fields.Item("desc_nivel").Value)%></option>
        <%
  rs_nivel.MoveNext()
Wend
If (rs_nivel.CursorType > 0) Then
  rs_nivel.MoveFirst
Else
  rs_nivel.Requery
End If
%>
      </select>      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style5">Senha:</span></td>
      <td bgcolor="#CCCCCC"><input type="password" name="senha" value="" size="15">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style6"></span></td>
      <td bgcolor="#CCCCCC"><input type="submit" value="Salvar">      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= rs_altera_usuario.Fields.Item("idusuario").Value %>">
</form>
<p>&nbsp;</p>
</body>
</html>
<%
rs_altera_usuario.Close()
Set rs_altera_usuario = Nothing
%>
<%
rs_nivel.Close()
Set rs_nivel = Nothing
%>
