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
  MM_editTable = "tb_pi"
  MM_editColumn = "PI"
  MM_recordId = "'" + Request.Form("MM_recordId") + "'"
  MM_editRedirectUrl = "contratos.asp"
  MM_fieldsStr  = "Foi_solicitado_Aditamento|value|Valor_do_Aditamento|value|Prazo_do_Aditamento|value"
  MM_columnsStr = "[Foi solicitado Aditamento?]|none,1,0|[Valor do Aditamento]|',none,''|[Prazo do Aditamento]|none,none,NULL"

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
Dim rs_contratos__MMColParam
rs_contratos__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  rs_contratos__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim rs_contratos
Dim rs_contratos_cmd
Dim rs_contratos_numRows

Set rs_contratos_cmd = Server.CreateObject ("ADODB.Command")
rs_contratos_cmd.ActiveConnection = MM_cpf_STRING
rs_contratos_cmd.CommandText = "SELECT tb_pi.*, [Redução]/100 AS reduc FROM tb_pi WHERE PI = ?" 
rs_contratos_cmd.Prepared = true
rs_contratos_cmd.Parameters.Append rs_contratos_cmd.CreateParameter("param1", 200, 1, 255, rs_contratos__MMColParam) ' adVarChar

Set rs_contratos = rs_contratos_cmd.Execute
rs_contratos_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style7 {font-family: Arial, Helvetica, sans-serif; font-size: 12; }
.style8 {font-size: 12}
.style11 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; }
.style17 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; }
-->
</style>
</head>

<body>
<div align="center">
  <table width="865" border="1">
    <tr bgcolor="#999999">
      <td width="97"><span class="style17">PI</span></td>
      <td width="124"><span class="style17">Crit&eacute;rio de C&aacute;lculo</span></td>
      <td width="125"><span class="style17">Crit&eacute;rio de Reajuste </span></td>
      <td width="57"><span class="style17">Data Assinatura </span></td>
      <td width="21"><span class="style17">Foi Solic.Adit? </span></td>
      <td width="18"><span class="style17">Prazo Contrato </span></td>
      <td width="18"><span class="style17">Prazo Aditamento </span></td>
      <td width="18"><span class="style17">Or&ccedil;amento FDE </span></td>
      <td width="130"><span class="style17">Redu&ccedil;&atilde;o</span></td>
      <td width="31"><span class="style17">Valor Contrato </span></td>
      <td width="36"><span class="style17">Valor Aditamento </span></td>
      <td width="114"><span class="style17">&Oacute;rg&atilde;o</span></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td><span class="style11"><%=(rs_contratos.Fields.Item("PI").Value)%></span></td>
      <td><span class="style11"><%=(rs_contratos.Fields.Item("Critério de Cálculo").Value)%></span></td>
      <td><span class="style11"><%=(rs_contratos.Fields.Item("Critério de Reajuste").Value)%></span></td>
      <td><span class="style11"><%=(rs_contratos.Fields.Item("Data da Assinatura").Value)%></span></td>
      <td><span class="style11"><%=(rs_contratos.Fields.Item("Foi solicitado Aditamento?").Value)%></span></td>
      <td><span class="style11"><%=(rs_contratos.Fields.Item("Prazo do Contrato").Value)%></span></td>
      <td><span class="style11"><%=(rs_contratos.Fields.Item("Prazo do Aditamento").Value)%></span></td>
      <td><span class="style11"><%= FormatNumber((rs_contratos.Fields.Item("Orçamento FDE").Value), 2, -2, -2, -2) %></span></td>
      <td class="style11"><%= FormatPercent((rs_contratos.Fields.Item("reduc").Value), 3, -2, -2, -2) %></td>
      <td><span class="style11"><%= FormatNumber((rs_contratos.Fields.Item("Valor do Contrato").Value), 2, -2, -2, -2) %></span></td>
      <td><span class="style11"><%= FormatNumber((rs_contratos.Fields.Item("Valor do Aditamento").Value), 2, -2, -2, -2) %></span></td>
      <td><span class="style11"><%=(rs_contratos.Fields.Item("Órgão").Value)%></span></td>
    </tr>
  </table>
</div>
<p>&nbsp;</p>
<form method="POST" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style7">Foi solicitado Aditamento?:</span></td>
      <td bgcolor="#CCCCCC"><input <%If (CStr((rs_contratos.Fields.Item("Foi solicitado Aditamento?").Value)) = CStr("True")) Then Response.Write("checked=""checked""") : Response.Write("")%> type="checkbox" name="Foi_solicitado_Aditamento" value=1 >      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style7">Valor do Aditamento:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="Valor_do_Aditamento" value="<%=(rs_contratos.Fields.Item("Valor do Aditamento").Value)%>" size="20">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style7">Prazo do Aditamento:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="Prazo_do_Aditamento" value="<%=(rs_contratos.Fields.Item("Prazo do Aditamento").Value)%>" size="15">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style8"></span></td>
      <td bgcolor="#CCCCCC"><input type="submit" value="Salvar">      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= rs_contratos.Fields.Item("PI").Value %>">
</form>
<p>&nbsp;</p>
</body>
</html>
<%
rs_contratos.Close()
Set rs_contratos = Nothing
%>
