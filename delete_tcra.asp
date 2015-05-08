<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
' *** Edit Operations: declare variables

Response.CharSet = "UTF-8"

Dim objCon
Set objCon = Server.CreateObject("ADODB.Connection")
    objCon.Open MM_cpf_STRING

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
  MM_editTable = "tb_tcra"
  MM_editColumn = "id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "cad_tcra.asp"

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
Dim rs_delete_us__MMColParam
rs_delete_us__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  rs_delete_us__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim rs_delete_us
Dim rs_delete_us_numRows

strQ = "select * from tb_tcra where id = " + Replace(rs_delete_us__MMColParam, "'", "''") + ""

Set rs_delete_us = Server.CreateObject("ADODB.Recordset")
rs_delete_us.CursorLocation = 3
rs_delete_us.CursorType = 3
rs_delete_us.LockType = 1
rs_delete_us.Open strQ, objCon, , , &H0001

rs_delete_us_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rs_delete_us_numRows = rs_delete_us_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style3 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; }
.style5 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
-->
</style>
</head>

<body>
<form id="form1" name="form1" method="POST" action="<%=MM_editAction%>">
  <input type="hidden" name="MM_delete" value="form1">
  <input type="hidden" name="MM_recordId" value="<%=(Request.QueryString("id"))%>">
  <label for="Submit"></label>
  <div align="center">
    <input type="submit" name="Submit" value="Excluir" id="Submit" />
  </div>
</form>
<p>&nbsp;</p>
<div align="center">
  <table border="0">
    <tr bgcolor="#CCCCCC">
      <td>
        <span class="style5">Núm. Outorga</span>
      </td>
      <td>
        <span class="style5">Data de Concessão</span>
      </td>
      <td>
        <span class="style5">Data de Vencimento</span>
      </td>
      <td>
        <span class="style5">Observações</span>
      </td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rs_delete_us.EOF)) %>
      <tr bgcolor="#CCCCCC">
        <td>
          <span class="style3"><%=(rs_delete_us.Fields.Item("cod_tcra").Value)%></span>
        </td>
        <td>
          <span class="style3"><%=(rs_delete_us.Fields.Item("dta_concessao").Value)%></span>
        </td>
        <td>
          <span class="style3"><%=(rs_delete_us.Fields.Item("dta_vencimento").Value)%></span>
        </td>
        <td>
          <span class="style3"><%=(rs_delete_us.Fields.Item("dsc_observacoes").Value)%></span>
        </td>
      </tr>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_delete_us.MoveNext()
Wend
%>
  </table>
</div>
</body>
</html>
<%
rs_delete_us.Close()
Set rs_delete_us = Nothing
%>
