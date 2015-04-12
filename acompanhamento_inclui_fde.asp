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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_cpf_STRING
  MM_editTable = "tb_Acompanhamento"
  MM_editRedirectUrl = "acompanhamento_inclui.asp"
  MM_fieldsStr  = "PI|value|Data_do_Registro|value|Responsvel|value|Previso|value|Registro|value"
  MM_columnsStr = "PI|none,none,NULL|[Data do Registro]|',none,NULL|Responsável|none,none,NULL|Previsão|',none,''|Registro|',none,''"

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
Dim rs_pi__MMColParam
rs_pi__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  rs_pi__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim rs_pi
Dim rs_pi_numRows

Set rs_pi = Server.CreateObject("ADODB.Recordset")
rs_pi.ActiveConnection = MM_cpf_STRING
rs_pi.Source = "SELECT tb_pi.PI, [tb_predio].[cod_predio] & ' - ' & [tb_predio].[Nome_Unidade] AS Expr1, c_Semaforico1.[Avanço Físico Atual]  FROM c_Semaforico1 INNER JOIN (tb_predio INNER JOIN tb_pi ON tb_predio.cod_predio = tb_pi.cod_predio) ON c_Semaforico1.[PI-item] = tb_pi.PI  WHERE PI = '" + Replace(rs_pi__MMColParam, "'", "''") + "'  ORDER BY tb_pi.PI;"
rs_pi.CursorType = 0
rs_pi.CursorLocation = 2
rs_pi.LockType = 1
rs_pi.Open()

rs_pi_numRows = 0
%>
<%
Dim rs_fiscal
Dim rs_fiscal_numRows

Set rs_fiscal = Server.CreateObject("ADODB.Recordset")
rs_fiscal.ActiveConnection = MM_cpf_STRING
rs_fiscal.Source = "SELECT *  FROM tb_responsavel  ORDER BY Responsável ASC"
rs_fiscal.CursorType = 0
rs_fiscal.CursorLocation = 2
rs_fiscal.LockType = 1
rs_fiscal.Open()

rs_fiscal_numRows = 0
%>
<%
Dim rs_acomp__MMColParam
rs_acomp__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  rs_acomp__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim rs_acomp
Dim rs_acomp_numRows

Set rs_acomp = Server.CreateObject("ADODB.Recordset")
rs_acomp.ActiveConnection = MM_cpf_STRING
rs_acomp.Source = "SELECT *  FROM tb_Acompanhamento  WHERE PI = '" + Replace(rs_acomp__MMColParam, "'", "''") + "'"
rs_acomp.CursorType = 0
rs_acomp.CursorLocation = 2
rs_acomp.LockType = 1
rs_acomp.Open()

rs_acomp_numRows = 0
%>
<%
Dim rs_lista_acomp__MMColParam
rs_lista_acomp__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  rs_lista_acomp__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim rs_lista_acomp
Dim rs_lista_acomp_numRows

Set rs_lista_acomp = Server.CreateObject("ADODB.Recordset")
rs_lista_acomp.ActiveConnection = MM_cpf_STRING
rs_lista_acomp.Source = "SELECT *  FROM cLista_acomp  WHERE PI = '" + Replace(rs_lista_acomp__MMColParam, "'", "''") + "'"
rs_lista_acomp.CursorType = 0
rs_lista_acomp.CursorLocation = 2
rs_lista_acomp.LockType = 1
rs_lista_acomp.Open()

rs_lista_acomp_numRows = 0
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
Recordset1.Source = "SELECT tb_pi.PI, tb_pi.[Descrição da Intervenção FDE]  FROM tb_pi  WHERE PI = '" + Replace(Recordset1__MMColParam, "'", "''") + "'  GROUP BY tb_pi.PI, tb_pi.[Descrição da Intervenção FDE];    "
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rs_acomp_numRows = rs_acomp_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = 1000
Repeat2__index = 0
rs_lista_acomp_numRows = rs_lista_acomp_numRows + Repeat2__numRows
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
.style19 {font-family: Arial, Helvetica, sans-serif; font-size: 14px; font-weight: bold; }
.style22 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
.style26 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
.style27 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
}
-->
</style>
</head>

<body>
<div align="center">
  <table width="823" border="0">
    <tr bgcolor="#CCCCCC">
      <td><span class="style19"><%=(rs_pi.Fields.Item("PI").Value)%></span></td>
      <td colspan="2"><div align="left"><span class="style19"><%=(rs_pi.Fields.Item("Expr1").Value)%></span></div></td>
      <td><span class="style27">Avan&ccedil;o F&iacute;sico Atual: <%= FormatPercent((rs_pi.Fields.Item("Avanço Físico Atual").Value), 2, -2, -2, -2) %></span></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="147" bgcolor="#CCCCCC" class="style27">Interven&ccedil;&atilde;o FDE </td>
      <td width="326" bgcolor="#CCCCCC"><div align="left" class="style27"><%=(Recordset1.Fields.Item("Descrição da Intervenção FDE").Value)%></div></td>
      <td width="110" bgcolor="#CCCCCC" class="style27">T&eacute;rmino contratual</td>
      <td width="222" bgcolor="#CCCCCC"><span class="style26"><%=(rs_lista_acomp.Fields.Item("termino_contratual").Value)%></span></td>
    </tr>
  </table>
  <p>&nbsp;</p>
</div>
<div align="center">
  <table border="0">
    <tr bgcolor="#999999">
      <td><span class="style26">Data do Registro</span></td>
      <td><span class="style26">Registro</span></td>
      <td><span class="style26">Previs&atilde;o</span></td>
      <td><span class="style26">Respons&aacute;vel</span></td>
    </tr>
    <% While ((Repeat2__numRows <> 0) AND (NOT rs_lista_acomp.EOF)) %>
      <tr bgcolor="#CCCCCC">
        <td><span class="style22"><%=(rs_lista_acomp.Fields.Item("Data do Registro").Value)%></span></td>
        <td><div align="left"><span class="style22"><%=(rs_lista_acomp.Fields.Item("Registro").Value)%></span></div></td>
        <td><span class="style22"><%=(rs_lista_acomp.Fields.Item("Previsão").Value)%></span></td>
        <td><span class="style22"><%=(rs_lista_acomp.Fields.Item("Responsável").Value)%></span></td>
      </tr>
      <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rs_lista_acomp.MoveNext()
Wend
%>
  </table>
</div>
</body>
</html>
<%
rs_pi.Close()
Set rs_pi = Nothing
%>
<%
rs_fiscal.Close()
Set rs_fiscal = Nothing
%>
<%
rs_acomp.Close()
Set rs_acomp = Nothing
%>
<%
rs_lista_acomp.Close()
Set rs_lista_acomp = Nothing
%>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
