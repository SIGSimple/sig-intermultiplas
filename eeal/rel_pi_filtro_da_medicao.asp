<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_cpf_STRING
Recordset1.Source = "SELECT tb_predio.cod_predio, [tb_predio].[cod_predio] & ' - ' & [tb_predio].[Nome_Unidade] AS Expr1  FROM tb_predio inner JOIN tb_PI ON tb_predio.cod_predio = tb_PI.cod_predio  GROUP BY tb_predio.cod_predio, [tb_predio].[cod_predio] & ' - ' & [tb_predio].[Nome_Unidade]  ORDER BY [tb_predio].[cod_predio] & ' - ' & [tb_predio].[Nome_Unidade];"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim Recordset2__MMColParam
Recordset2__MMColParam = "1"
If (Request.Form("cod_predio") <> "") Then 
  Recordset2__MMColParam = Request.Form("cod_predio")
End If
%>
<%
Dim Recordset2__MMColParam
Recordset2__MMColParam = "1"
If (Request.Form("cod_predio") <> "") Then 
  Recordset2__MMColParam = Request.Form("cod_predio")
End If
%>
<%
Dim Recordset2
Dim Recordset2_numRows

Set Recordset2 = Server.CreateObject("ADODB.Recordset")
Recordset2.ActiveConnection = MM_cpf_STRING
Recordset2.Source = "SELECT tb_pi.cod_predio, tb_pi.PI, tb_responsavel.Responsável  FROM tb_responsavel RIGHT JOIN (tb_predio INNER JOIN tb_pi ON tb_predio.cod_predio = tb_pi.cod_predio) ON tb_responsavel.cod_fiscal = tb_pi.cod_fiscal  WHERE tb_predio.cod_predio = '" + Replace(Recordset2__MMColParam, "'", "''") + "'  GROUP BY tb_pi.cod_predio, tb_pi.PI, tb_responsavel.Responsável  ORDER BY tb_pi.cod_predio, tb_pi.PI;    "
Recordset2.CursorType = 0
Recordset2.CursorLocation = 2
Recordset2.LockType = 1
Recordset2.Open()

Recordset2_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
Recordset2_numRows = Recordset2_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style3 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; }
.style4 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
.style6 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; color: #003399; }
.style8 {font-family: Arial, Helvetica, sans-serif; font-size: 18px; font-weight: bold; color: #003399; }
-->
</style>
</head>

<body>
<form id="form1" name="form1" method="post" action="">
  <label for="select"></label>
  <select name="select" id="select">
    <option value=""></option>
    <%
While (NOT Recordset1.EOF)
%>
    <option value="<%=(Recordset1.Fields.Item("cod_predio").Value)%>"><%=(Recordset1.Fields.Item("Expr1").Value)%></option>
    <%
  Recordset1.MoveNext()
Wend
If (Recordset1.CursorType > 0) Then
  Recordset1.MoveFirst
Else
  Recordset1.Requery
End If
%>
  </select>
  <label for="Submit"></label>
  <input type="submit" name="Submit" value="Buscar" id="Submit" />
</form>
<table border="0">
  <tr bgcolor="#CCCCCC">
    <td width="140"><span class="style4">cod_predio</span></td>
    <td width="89"><span class="style4">PI</span></td>
    <td width="151"><span class="style4">Respons&aacute;vel</span></td>
    <td colspan="2"><div align="center" class="style4"><span class="style8">A&Ccedil;&Otilde;ES</span></div>      <div align="center" class="style4"></div></td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT Recordset2.EOF)) %>
    <tr bgcolor="#CCCCCC">
      <td><span class="style3"><%=(Recordset2.Fields.Item("cod_predio").Value)%></span></td>
      <td><span class="style3"><%=(Recordset2.Fields.Item("PI").Value)%></span></td>
      <td><span class="style3"><%=(Recordset2.Fields.Item("Responsável").Value)%></span></td>
      <td width="133"><div align="center"><span class="style6"><a href="atualiza_PI_fiscal.asp?PI=<%=(Recordset2.Fields.Item("PI").Value)%>">PI</a></span></div></td>
      <td width="230"><div align="center"><span class="style6"><a href="acompanhamento_inclui.asp?PI=<%=(Recordset2.Fields.Item("PI").Value)%>">Acompanhamento</a></span></div></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset2.MoveNext()
Wend
%>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>
