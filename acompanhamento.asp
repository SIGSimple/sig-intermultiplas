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
  MM_editRedirectUrl = "filtro_acomp.asp"
  MM_fieldsStr  = "cod_predio|value|PI|value|Data_do_Registro|value|Previso|value|Registro|value|cod_fiscal|value|trmino_contratual|value"
  MM_columnsStr = "cod_predio|',none,''|PI|none,none,NULL|[Data do Registro]|',none,NULL|Previsão|',none,''|Registro|',none,''|cod_fiscal|none,none,NULL|[término contratual]|',none,''"

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
Dim rs_acomp__MMColParam
rs_acomp__MMColParam = "1"
If (Request.Form("cod_predio") <> "") Then 
  rs_acomp__MMColParam = Request.Form("cod_predio")
End If
%>
<%
Dim rs_acomp
Dim rs_acomp_numRows

Set rs_acomp = Server.CreateObject("ADODB.Recordset")
rs_acomp.ActiveConnection = MM_cpf_STRING
rs_acomp.Source = "SELECT tb_predio.cod_predio, tb_pi.PI, tb_predio.Nome_Unidade, tb_pi.[Descrição da Intervenção FDE], tb_responsavel.Responsável  FROM (tb_predio INNER JOIN tb_pi ON tb_predio.cod_predio = tb_pi.cod_predio) INNER JOIN tb_responsavel ON tb_pi.cod_fiscal = tb_responsavel.cod_fiscal  WHERE tb_predio.cod_predio = '" + Replace(rs_acomp__MMColParam, "'", "''") + "'  ORDER BY tb_predio.cod_predio, tb_pi.PI;    "
rs_acomp.CursorType = 0
rs_acomp.CursorLocation = 2
rs_acomp.LockType = 1
rs_acomp.Open()

rs_acomp_numRows = 0
%>
<%
Dim rs_pi__MMColParam
rs_pi__MMColParam = "1"
If (Request.Form("cod_predio") <> "") Then 
  rs_pi__MMColParam = Request.Form("cod_predio")
End If
%>
<%
Dim rs_pi
Dim rs_pi_numRows

Set rs_pi = Server.CreateObject("ADODB.Recordset")
rs_pi.ActiveConnection = MM_cpf_STRING
rs_pi.Source = "SELECT tb_pi.PI, [tb_predio].[cod_predio] & ' - ' & [Nome_Unidade] AS Expr1  FROM tb_predio INNER JOIN tb_pi ON tb_predio.cod_predio = tb_pi.cod_predio  WHERE tb_predio.cod_predio = '" + Replace(rs_pi__MMColParam, "'", "''") + "'  ORDER BY tb_pi.PI;  "
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
rs_fiscal.Source = "SELECT * FROM tb_responsavel ORDER BY Responsável ASC"
rs_fiscal.CursorType = 0
rs_fiscal.CursorLocation = 2
rs_fiscal.LockType = 1
rs_fiscal.Open()

rs_fiscal_numRows = 0
%>
<%
Dim rs_lista_acomp
Dim rs_lista_acomp_numRows

Set rs_lista_acomp = Server.CreateObject("ADODB.Recordset")
rs_lista_acomp.ActiveConnection = MM_cpf_STRING
rs_lista_acomp.Source = "SELECT tb_Acompanhamento.PI, tb_Acompanhamento.[Data do Registro], tb_predio.cod_predio, tb_predio.Nome_Unidade, tb_responsavel.Responsável, tb_Acompanhamento.Registro, tb_Acompanhamento.Previsão, tb_Acompanhamento.[término contratual]  FROM tb_responsavel INNER JOIN ((tb_Acompanhamento INNER JOIN tb_predio ON tb_Acompanhamento.cod_predio = tb_predio.cod_predio) INNER JOIN tb_pi ON tb_predio.cod_predio = tb_pi.cod_predio) ON tb_responsavel.cod_fiscal = tb_pi.cod_fiscal  ORDER BY tb_Acompanhamento.PI, tb_Acompanhamento.[Data do Registro];  "
rs_lista_acomp.CursorType = 0
rs_lista_acomp.CursorLocation = 2
rs_lista_acomp.LockType = 1
rs_lista_acomp.Open()

rs_lista_acomp_numRows = 0
%>
<%
Dim rs_predio__MMColParam
rs_predio__MMColParam = "1"
If (Request.Form("cod_predio") <> "") Then 
  rs_predio__MMColParam = Request.Form("cod_predio")
End If
%>
<%
Dim rs_predio
Dim rs_predio_numRows

Set rs_predio = Server.CreateObject("ADODB.Recordset")
rs_predio.ActiveConnection = MM_cpf_STRING
rs_predio.Source = "SELECT * FROM tb_predio WHERE cod_predio = '" + Replace(rs_predio__MMColParam, "'", "''") + "'"
rs_predio.CursorType = 0
rs_predio.CursorLocation = 2
rs_predio.LockType = 1
rs_predio.Open()

rs_predio_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rs_acomp_numRows = rs_acomp_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Acompanhamento</title>
<style type="text/css">
<!--
.style5 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
.style19 {font-family: Arial, Helvetica, sans-serif; font-size: 12; font-weight: bold; }
.style21 {font-family: Arial, Helvetica, sans-serif; font-size: 12; }
-->
</style>
</head>

<body>
<table border="0">
  <tr bgcolor="#999999" class="style5">
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><span class="style19">PI</span></td>
    <td><span class="style19">cod_predio</span></td>
    <td><span class="style19">Nome da Unidade</span></td>
    <td><span class="style19">Descri&ccedil;&atilde;o da Interven&ccedil;&atilde;o FDE</span></td>
    <td><span class="style19">Respons&aacute;vel</span></td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT rs_acomp.EOF)) %>
    <tr bgcolor="#CCCCCC" class="style5">
      <td><a href="delete_acomp.asp?cod_acompanhamento=<%=(rs_acomp.Fields.Item("cod_acompanhamento").Value)%>"></a></td>
      <td><a href="delete_acomp.asp?pi=<%=(rs_acomp.Fields.Item("PI").Value)%>"></a><span class="style21"><%=(rs_acomp.Fields.Item("cod_acompanhamento").Value)%></span></td>
      <td class="style21"><a href="acompanhamento_inclui.asp?pi=<%=(rs_acomp.Fields.Item("PI").Value)%>"><%=(rs_acomp.Fields.Item("PI").Value)%></a></td>
      <td class="style21"><%=(rs_acomp.Fields.Item("cod_predio").Value)%></td>
      <td class="style21"><%=(rs_acomp.Fields.Item("Nome_Unidade").Value)%></td>
      <td class="style21"><%=(rs_acomp.Fields.Item("Descrição da Intervenção FDE").Value)%></td>
      <td><%=(rs_acomp.Fields.Item("Responsável").Value)%></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_acomp.MoveNext()
Wend
%>
</table>
</body>
</html>
<%
rs_acomp.Close()
Set rs_acomp = Nothing
%>
<%
rs_pi.Close()
Set rs_pi = Nothing
%>
<%
rs_fiscal.Close()
Set rs_fiscal = Nothing
%>
<%
rs_lista_acomp.Close()
Set rs_lista_acomp = Nothing
%>
<%
rs_predio.Close()
Set rs_predio = Nothing
%>
