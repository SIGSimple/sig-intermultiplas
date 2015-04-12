<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim rs_responsavel
Dim rs_responsavel_numRows

Set rs_responsavel = Server.CreateObject("ADODB.Recordset")
rs_responsavel.ActiveConnection = MM_cpf_STRING
rs_responsavel.Source = "SELECT tb_responsavel.cod_fiscal, tb_responsavel.Responsável  FROM tb_responsavel INNER JOIN tb_pi ON tb_responsavel.cod_fiscal = tb_pi.cod_fiscal  GROUP BY tb_responsavel.cod_fiscal, tb_responsavel.Responsável  ORDER BY tb_responsavel.Responsável;  "
rs_responsavel.CursorType = 0
rs_responsavel.CursorLocation = 2
rs_responsavel.LockType = 1
rs_responsavel.Open()

rs_responsavel_numRows = 0
%>
<%
Dim rs_PIS__MMColParam
rs_PIS__MMColParam = "1"
If (Request.Form("cod_fiscal") <> "") Then 
  rs_PIS__MMColParam = Request.Form("cod_fiscal")
End If
%>
<%
Dim rs_PIS
Dim rs_PIS_numRows

Set rs_PIS = Server.CreateObject("ADODB.Recordset")
rs_PIS.ActiveConnection = MM_cpf_STRING
rs_PIS.Source = "SELECT tb_pi.cod_predio, tb_pi.PI, tb_responsavel.Responsável, tb_predio.Nome_Unidade  FROM tb_predio INNER JOIN (tb_responsavel INNER JOIN tb_pi ON tb_responsavel.cod_fiscal = tb_pi.cod_fiscal) ON tb_predio.cod_predio = tb_pi.cod_predio  WHERE tb_pi.cod_fiscal = " + Replace(rs_PIS__MMColParam, "'", "''") + "  ORDER BY tb_pi.cod_predio, tb_pi.PI;    "
rs_PIS.CursorType = 0
rs_PIS.CursorLocation = 2
rs_PIS.LockType = 1
rs_PIS.Open()

rs_PIS_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rs_PIS_numRows = rs_PIS_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style5 {font-family: Arial, Helvetica, sans-serif; font-size: 14px; font-weight: bold; }
.style7 {font-family: Arial, Helvetica, sans-serif; font-size: 14px; font-weight: bold; color: #FFFFFF; }
.style8 {
	color: #660033;
	font-weight: bold;
}
.style9 {font-family: Arial, Helvetica, sans-serif; font-size: 18px; font-weight: bold; }
.style10 {font-size: 14px; font-family: Arial, Helvetica, sans-serif;}
.style12 {font-family: Arial, Helvetica, sans-serif; font-size: 18px; font-weight: bold; color: #003399; }
-->
</style>
</head>

<body>
<form id="form1" name="form1" method="post" action="">
  <label for="select"></label>
  <select name="Cod_fiscal" class="style5" id="Cod_fiscal">
    <option value=""></option>
    <%
While (NOT rs_responsavel.EOF)
%>
    <option value="<%=(rs_responsavel.Fields.Item("cod_fiscal").Value)%>"><%=(rs_responsavel.Fields.Item("Responsável").Value)%></option>
    <%
  rs_responsavel.MoveNext()
Wend
If (rs_responsavel.CursorType > 0) Then
  rs_responsavel.MoveFirst
Else
  rs_responsavel.Requery
End If
%>
  </select>
  <label for="Submit"></label>
  <input type="submit" name="Submit" value="Buscar" id="Submit" />
  <span class="style12"></span>
  <span class="style9"></span>
</form>

<table border="0">
  <tr bgcolor="#999999">
    <td width="126"><span class="style7">cod_predio</span></td>
    <td width="358"><span class="style7">Nome Unidade </span></td>
    <td width="182"><span class="style7">PI</span></td>
    <td colspan="3"><div align="center" class="style9">A&Ccedil;&Otilde;ES</div></td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT rs_PIS.EOF)) %>
    <tr bgcolor="#EFEFEF">
      <td class="style10"><%=(rs_PIS.Fields.Item("cod_predio").Value)%></td>
      <td class="style10"><%=(rs_PIS.Fields.Item("Nome_Unidade").Value)%></td>
      <td class="style10"><%=(rs_PIS.Fields.Item("PI").Value)%></td>
      <td width="67" bgcolor="#FFEFE8" class="style5"><div align="center" class="style8"><a href="atualiza_PI_adm.asp?pi=<%=(rs_PIS.Fields.Item("PI").Value)%>" target="_blank">PI</a></div></td>
      <td width="125" bgcolor="#FFEFE8" class="style5"><div align="center"><strong class="style8"><a href="acompanhamento_inclui_adm.asp?pi=<%=(rs_PIS.Fields.Item("PI").Value)%>" target="_blank">Acompanhamento</a></strong></div></td>
      <td width="139" bgcolor="#FFEFE8" class="style5"><div align="center"><strong class="style8"><a href="med_constr_inclui.asp?pi=<%=(rs_PIS.Fields.Item("PI").Value)%>" target="_blank">Medi&ccedil;&otilde;es</a></strong></div></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_PIS.MoveNext()
Wend
%>
</table>
</body>
</html>
<%
rs_responsavel.Close()
Set rs_responsavel = Nothing
%>
<%
rs_PIS.Close()
Set rs_PIS = Nothing
%>