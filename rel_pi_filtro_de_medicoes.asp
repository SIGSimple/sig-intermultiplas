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
Dim rs_PIS__MMColParam1
rs_PIS__MMColParam1 = "1"
If (Request.Form("cod_situacao") <> "") Then 
  rs_PIS__MMColParam1 = Request.Form("cod_situacao")
End If
%>
<%
Dim rs_PIS
Dim rs_PIS_cmd
Dim rs_PIS_numRows

Set rs_PIS_cmd = Server.CreateObject ("ADODB.Command")
rs_PIS_cmd.ActiveConnection = MM_cpf_STRING
rs_PIS_cmd.CommandText = "SELECT tb_situacao_pi.desc_situacao, tb_pi.cod_predio, tb_pi.PI, tb_responsavel.Responsável, tb_predio.Nome_Unidade FROM tb_situacao_pi INNER JOIN (tb_responsavel INNER JOIN (tb_predio INNER JOIN tb_pi ON tb_predio.cod_predio = tb_pi.cod_predio) ON tb_responsavel.cod_fiscal = tb_pi.cod_fiscal) ON tb_situacao_pi.cod_situacao = tb_pi.cod_situacao WHERE tb_pi.cod_fiscal = ? and tb_situacao_pi.cod_situacao = ? ORDER BY tb_pi.cod_predio;" 
rs_PIS_cmd.Prepared = true
rs_PIS_cmd.Parameters.Append rs_PIS_cmd.CreateParameter("param1", 5, 1, -1, rs_PIS__MMColParam) ' adDouble
rs_PIS_cmd.Parameters.Append rs_PIS_cmd.CreateParameter("param2", 5, 1, -1, rs_PIS__MMColParam1) ' adDouble

Set rs_PIS = rs_PIS_cmd.Execute
rs_PIS_numRows = 0
%>
<%
Dim rs_situacao
Dim rs_situacao_cmd
Dim rs_situacao_numRows

Set rs_situacao_cmd = Server.CreateObject ("ADODB.Command")
rs_situacao_cmd.ActiveConnection = MM_cpf_STRING
rs_situacao_cmd.CommandText = "SELECT tb_situacao_pi.cod_situacao, tb_situacao_pi.desc_situacao FROM tb_situacao_pi ORDER BY tb_situacao_pi.desc_situacao; " 
rs_situacao_cmd.Prepared = true

Set rs_situacao = rs_situacao_cmd.Execute
rs_situacao_numRows = 0
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
.style13 {
	font-size: 12px
}
-->
</style>
</head>

<body>
<form id="form1" name="form1" method="post" action="">
  <label for="select"></label>
  <span class="style10">Fiscal</span>
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
  <span class="style10">Situa&ccedil;&atilde;o</span>
  <select name="cod_situacao" id="cod_situacao">
    <option value=""></option>
    <%
While (NOT rs_situacao.EOF)
%>
    <option value="<%=(rs_situacao.Fields.Item("cod_situacao").Value)%>"><%=(rs_situacao.Fields.Item("desc_situacao").Value)%></option>
    <%
  rs_situacao.MoveNext()
Wend
If (rs_situacao.CursorType > 0) Then
  rs_situacao.MoveFirst
Else
  rs_situacao.Requery
End If
%>
  </select>
  <input type="submit" name="Submit" value="Buscar" id="Submit" />
  <span class="style12"></span>
  <span class="style9"></span>
</form>

<p class="style10 style13"><a href="rel_pi_filtrofiscal.asp" target="_blank"><U>RELAT&Oacute;RIO DE PLANEJAMENTO E CONTROLE DAS MEDI&Ccedil;&Otilde;ES</U></a></p>
<table border="0">
  <tr bgcolor="#999999">
    <td width="181"><span class="style7">cod_predio</span></td>
    <td width="168"><span class="style7">PI</span></td>
    <td colspan="3"><div align="center" class="style9">A&Ccedil;&Otilde;ES</div></td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT rs_PIS.EOF)) %>
    <tr bgcolor="#EFEFEF">
      <td class="style10"><%=(rs_PIS.Fields.Item("cod_predio").Value)%></td>
      <td class="style10"><%=(rs_PIS.Fields.Item("PI").Value)%></td>
      <td width="56" bgcolor="#FFEFE8" class="style5"><div align="center" class="style8"><a href="atualiza_PI_adm.asp?pi=<%=(rs_PIS.Fields.Item("PI").Value)%>" target="_blank">PI</a></div></td>
      <td width="302" bgcolor="#FFEFE8" class="style5"><div align="center"><strong class="style8"><a href="acompanhamento_inclui_adm.asp?pi=<%=(rs_PIS.Fields.Item("PI").Value)%>" target="_blank">Acompanhamento</a></strong></div></td>
      <td width="142" bgcolor="#FFEFE8" class="style5"><div align="center"><strong class="style8"><a href="med_constr_inclui.asp?pi=<%=(rs_PIS.Fields.Item("PI").Value)%>" target="_blank">Medi&ccedil;&otilde;es</a></strong></div></td>
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
<%
rs_situacao.Close()
Set rs_situacao = Nothing
%>