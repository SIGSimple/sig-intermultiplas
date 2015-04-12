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
  MM_editTable = "tb_Medicao_Gerenciadora"
  MM_editColumn = "PI"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "medicoes_gerenciadora.asp"

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
Dim rs_exclui_med_ger__MMColParam
rs_exclui_med_ger__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  rs_exclui_med_ger__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim rs_exclui_med_ger
Dim rs_exclui_med_ger_numRows

Set rs_exclui_med_ger = Server.CreateObject("ADODB.Recordset")
rs_exclui_med_ger.ActiveConnection = MM_cpf_STRING
rs_exclui_med_ger.Source = "SELECT * FROM tb_Medicao_Gerenciadora WHERE PI = " + Replace(rs_exclui_med_ger__MMColParam, "'", "''") + ""
rs_exclui_med_ger.CursorType = 0
rs_exclui_med_ger.CursorLocation = 2
rs_exclui_med_ger.LockType = 1
rs_exclui_med_ger.Open()

rs_exclui_med_ger_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rs_exclui_med_ger_numRows = rs_exclui_med_ger_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style3 {font-family: Arial, Helvetica, sans-serif; font-size: 11px; }
-->
</style>
</head>

<body>
<form id="form1" name="form1" method="POST" action="<%=MM_editAction%>">
  <div align="center">
    <input type="hidden" name="MM_delete" value="form1">
    <input type="hidden" name="MM_recordId" value="<%= rs_exclui_med_ger.Fields.Item("PI").Value %>">
  </div>
  <label for="Submit"></label>
  <div align="center">
    <input type="submit" name="Submit" value="Excluir" id="Submit" />
  </div>
</form>
<p>&nbsp;</p>

<div align="center">
  <table border="0">
    <tr bgcolor="#999999">
      <td><span class="style3">PI</span></td>
      <td><span class="style3">C&oacute;digo do Pr&eacute;dio</span></td>
      <td><span class="style3">Nome da Unidade</span></td>
      <td><span class="style3">Gerenciadora Mede ?</span></td>
      <td><span class="style3">N&uacute;mero da OS</span></td>
      <td><span class="style3">campo de trabalho</span></td>
      <td><span class="style3">Data de in&iacute;cio</span></td>
      <td><span class="style3">Dura&ccedil;&atilde;o da OS</span></td>
      <td><span class="style3">Data de T&eacute;rmino</span></td>
      <td><span class="style3">Data de T&eacute;rmino Contratual</span></td>
      <td><span class="style3">Justificativa</span></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rs_exclui_med_ger.EOF)) %>
      <tr bgcolor="#CCCCCC">
        <td><span class="style3"><%=(rs_exclui_med_ger.Fields.Item("PI").Value)%></span></td>
        <td><span class="style3"><%=(rs_exclui_med_ger.Fields.Item("Código do Prédio").Value)%></span></td>
        <td><span class="style3"><%=(rs_exclui_med_ger.Fields.Item("Nome da Unidade").Value)%></span></td>
        <td><span class="style3"><%=(rs_exclui_med_ger.Fields.Item("Gerenciadora Mede ?").Value)%></span></td>
        <td><span class="style3"><%=(rs_exclui_med_ger.Fields.Item("Número da OS").Value)%></span></td>
        <td><span class="style3"><%=(rs_exclui_med_ger.Fields.Item("campo de trabalho").Value)%></span></td>
        <td><span class="style3"><%=(rs_exclui_med_ger.Fields.Item("Data de início").Value)%></span></td>
        <td><span class="style3"><%=(rs_exclui_med_ger.Fields.Item("Duração da OS").Value)%></span></td>
        <td><span class="style3"><%=(rs_exclui_med_ger.Fields.Item("Data de Término").Value)%></span></td>
        <td><span class="style3"><%=(rs_exclui_med_ger.Fields.Item("Data de Término Contratual").Value)%></span></td>
        <td><span class="style3"><%=(rs_exclui_med_ger.Fields.Item("Justificativa").Value)%></span></td>
      </tr>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_exclui_med_ger.MoveNext()
Wend
%>
  </table>
</div>
</body>
</html>
<%
rs_exclui_med_ger.Close()
Set rs_exclui_med_ger = Nothing
%>
