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
  MM_editRedirectUrl = "Os_alteracao.asp"
  MM_fieldsStr  = "PI|value|Data_da_Impresso_da_OIS|value|Data_da_CI|value|Data_da_Abertura|value|Descrio_da_Interveno_Gerenciadora|value|Fator_de_Reduo|value|rea_gerenciada|value"
  MM_columnsStr = "PI|',none,''|[Data da Impressão da OIS]|',none,NULL|[Data da CI]|',none,NULL|[Data da Abertura]|',none,NULL|[Descrição da Intervenção Gerenciadora]|',none,''|[Fator de Redução]|',none,''|[Área gerenciada]|none,none,NULL"

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
Dim rs_altera_os__MMColParam
rs_altera_os__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  rs_altera_os__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim rs_altera_os
Dim rs_altera_os_numRows

Set rs_altera_os = Server.CreateObject("ADODB.Recordset")
rs_altera_os.ActiveConnection = MM_cpf_STRING
rs_altera_os.Source = "SELECT * FROM tb_pi WHERE PI = '" + Replace(rs_altera_os__MMColParam, "'", "''") + "'"
rs_altera_os.CursorType = 0
rs_altera_os.CursorLocation = 2
rs_altera_os.LockType = 1
rs_altera_os.Open()

rs_altera_os_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rs_altera_os_numRows = rs_altera_os_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Alter&ccedil;&atilde;o de OS</title>
<style type="text/css">
<!--
.style25 {font-family: Arial, Helvetica, sans-serif; font-size: 10; }
.style26 {font-size: 10}
.style27 {
	font-size: 24px;
	font-weight: bold;
}
.style30 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; }
.style34 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; color: #FFFFFF; }
-->
</style>
</head>

<body>
<p align="center" class="style25 style27">Dados da OS</p>
<form method="POST" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style25">PI:</span></td>
      <td bgcolor="#CCCCCC"><input name="PI" type="text" value="<%=(rs_altera_os.Fields.Item("PI").Value)%>" size="32" readonly="true">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style25">Data da Impressão da OIS:</span></td>
      <td bgcolor="#CCCCCC"><input name="Data_da_Impresso_da_OIS" type="text" value="<%=(rs_altera_os.Fields.Item("Data da Impressão da OIS").Value)%>" size="15" readonly="true">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style25">Data da CI:</span></td>
      <td bgcolor="#CCCCCC"><input name="Data_da_CI" type="text" value="<%=(rs_altera_os.Fields.Item("Data da CI").Value)%>" size="15" readonly="true">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style25">Data da Abertura:</span></td>
      <td bgcolor="#CCCCCC"><input name="Data_da_Abertura" type="text" value="<%=(rs_altera_os.Fields.Item("Data da Abertura").Value)%>" size="15" readonly="true">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style25">Descrição da Intervenção Gerenciadora:</span></td>
      <td bgcolor="#CCCCCC"><textarea name="Descrio_da_Interveno_Gerenciadora" cols="32" rows="" disabled="disabled"><%=(rs_altera_os.Fields.Item("Descrição da Intervenção Gerenciadora").Value)%></textarea>      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style25">Fator de Redução:</span></td>
      <td bgcolor="#CCCCCC"><input name="Fator_de_Reduo" type="text" value="<%=(rs_altera_os.Fields.Item("Fator de Redução").Value)%>" size="12" readonly="true">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style25">Área gerenciada:</span></td>
      <td bgcolor="#CCCCCC"><input name="rea_gerenciada" type="text" value="<%=(rs_altera_os.Fields.Item("Área gerenciada").Value)%>" size="12" readonly="true">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style26"></span></td>
      <td bgcolor="#CCCCCC">&nbsp;</td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= rs_altera_os.Fields.Item("PI").Value %>">
</form>
<form id="form2" name="form2" method="post" action="">
  <label for="select"></label>
  <div align="center">
    <label for="Submit"></label>
  </div>
</form>
<div align="center">
  <table border="0">
    <tr bgcolor="#666666">
      <td><span class="style34">Data da Impress&atilde;o da OIS</span></td>
      <td><span class="style34">Data da CI</span></td>
      <td><span class="style34">Data da Abertura</span></td>
      <td><span class="style34">Descri&ccedil;&atilde;o da Interven&ccedil;&atilde;o Gerenciadora</span></td>
      <td><span class="style34">Fator de Redu&ccedil;&atilde;o</span></td>
      <td><span class="style34">&Aacute;rea gerenciada</span></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rs_altera_os.EOF)) %>
      <tr bgcolor="#CCCCCC">
        <td><span class="style30"><%=(rs_altera_os.Fields.Item("Data da Impressão da OIS").Value)%></span></td>
        <td><span class="style30"><%=(rs_altera_os.Fields.Item("Data da CI").Value)%></span></td>
        <td><span class="style30"><%=(rs_altera_os.Fields.Item("Data da Abertura").Value)%></span></td>
        <td><span class="style30"><%=(rs_altera_os.Fields.Item("Descrição da Intervenção Gerenciadora").Value)%></span></td>
        <td><span class="style30"><%=(rs_altera_os.Fields.Item("Fator de Redução").Value)%></span></td>
        <td><span class="style30"><%=(rs_altera_os.Fields.Item("Área gerenciada").Value)%></span></td>
      </tr>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_altera_os.MoveNext()
Wend
%>
  </table>
</div>
</body>
</html>
<%
rs_altera_os.Close()
Set rs_altera_os = Nothing
%>