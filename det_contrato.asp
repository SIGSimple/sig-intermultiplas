<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/cpf.asp" -->
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
  MM_editTable = "tb_datas"
  MM_editColumn = "cod_predio"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "det_contrato.asp"
  MM_fieldsStr  = "cod_predio|value|dt_abertura|value|dt_assinatura|value|dt_base|value|dt_CI|value|dt_impressao_ois|value|dt_termino|value|PI|value"
  MM_columnsStr = "cod_predio|none,none,NULL|dt_abertura|',none,''|dt_assinatura|',none,''|dt_base|',none,''|dt_CI|',none,''|dt_impressao_ois|',none,''|dt_termino|',none,''|PI|none,none,NULL"

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
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_cpf_STRING
Recordset1.Source = "SELECT cod_predio, dt_abertura, dt_assinatura, dt_base, dt_CI, dt_impressao_ois, dt_termino, PI FROM tb_datas WHERE PI = " + Replace(Recordset1__MMColParam, "'", "''") + ""
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
</head>

<body>
<form method="post" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td nowrap align="right">Cod_predio:</td>
      <td><input type="text" name="cod_predio" value="<%=(Recordset1.Fields.Item("cod_predio").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Dt_abertura:</td>
      <td><input type="text" name="dt_abertura" value="<%=(Recordset1.Fields.Item("dt_abertura").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Dt_assinatura:</td>
      <td><input type="text" name="dt_assinatura" value="<%=(Recordset1.Fields.Item("dt_assinatura").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Dt_base:</td>
      <td><input type="text" name="dt_base" value="<%=(Recordset1.Fields.Item("dt_base").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Dt_CI:</td>
      <td><input type="text" name="dt_CI" value="<%=(Recordset1.Fields.Item("dt_CI").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Dt_impressao_ois:</td>
      <td><input type="text" name="dt_impressao_ois" value="<%=(Recordset1.Fields.Item("dt_impressao_ois").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Dt_termino:</td>
      <td><input type="text" name="dt_termino" value="<%=(Recordset1.Fields.Item("dt_termino").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">PI:</td>
      <td><input type="text" name="PI" value="<%=(Recordset1.Fields.Item("PI").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">&nbsp;</td>
      <td><input type="submit" value="salvar">
      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("cod_predio").Value %>">
</form>
<p>&nbsp;</p>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>