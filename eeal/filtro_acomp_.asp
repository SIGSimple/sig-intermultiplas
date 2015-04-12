<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim rs_lista_predio_usua__MMColParam
rs_lista_predio_usua__MMColParam = "1"
If (Request.Form("nome") <> "") Then 
  rs_lista_predio_usua__MMColParam = Request.Form("nome")
End If
%>
<%
Dim rs_lista_predio_usua
Dim rs_lista_predio_usua_numRows

Set rs_lista_predio_usua = Server.CreateObject("ADODB.Recordset")
rs_lista_predio_usua.ActiveConnection = MM_cpf_STRING
rs_lista_predio_usua.Source = "SELECT *  FROM cLista_predio_usuario  WHERE nome = '" + Replace(rs_lista_predio_usua__MMColParam, "'", "''") + "'"
rs_lista_predio_usua.CursorType = 0
rs_lista_predio_usua.CursorLocation = 2
rs_lista_predio_usua.LockType = 1
rs_lista_predio_usua.Open()

rs_lista_predio_usua_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rs_lista_predio_usua_numRows = rs_lista_predio_usua_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style3 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
.style7 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; color: #FFFFFF; }
-->
</style>
</head>

<body>
<table border="0">
  <tr bgcolor="#666666">
    <td width="158"><span class="style7">cod_predio</span></td>
    <td width="312"><span class="style7">Predio</span></td>
    <td width="260"><span class="style7">nome</span></td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT rs_lista_predio_usua.EOF)) %>
    <tr bgcolor="#CCCCCC">
      <td><span class="style3"><%=(rs_lista_predio_usua.Fields.Item("cod_predio").Value)%></span></td>
      <td><span class="style3"><%=(rs_lista_predio_usua.Fields.Item("Expr1").Value)%></span></td>
      <td><span class="style3"><%=(rs_lista_predio_usua.Fields.Item("nome").Value)%></span></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_lista_predio_usua.MoveNext()
Wend
%>
</table>
</body>
</html>
<%
rs_lista_predio_usua.Close()
Set rs_lista_predio_usua = Nothing
%>
