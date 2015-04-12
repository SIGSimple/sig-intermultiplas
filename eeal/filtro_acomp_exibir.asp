<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim rs_lista_acomp__MMColParam
rs_lista_acomp__MMColParam = "1"
If (Request.Form("cod_predio") <> "") Then 
  rs_lista_acomp__MMColParam = Request.Form("cod_predio")
End If
%>
<%
Dim rs_lista_acomp
Dim rs_lista_acomp_numRows

Set rs_lista_acomp = Server.CreateObject("ADODB.Recordset")
rs_lista_acomp.ActiveConnection = MM_cpf_STRING
rs_lista_acomp.Source = "SELECT tb_predio.cod_predio, tb_predio.Nome_Unidade, tb_pi.PI, tb_pi.[Descrição da Intervenção FDE], tb_Acompanhamento.[Data do Registro], tb_Acompanhamento.Registro, tb_responsavel.Responsável, tb_Acompanhamento.Previsão, tb_Acompanhamento.[término contratual], tb_Acompanhamento.cod_acompanhamento  FROM (tb_responsavel INNER JOIN (tb_predio INNER JOIN tb_pi ON tb_predio.cod_predio = tb_pi.cod_predio) ON tb_responsavel.cod_fiscal = tb_pi.cod_fiscal) INNER JOIN tb_Acompanhamento ON tb_pi.PI = tb_Acompanhamento.PI  WHERE [tb_predio].[cod_predio] = '" + Replace(rs_lista_acomp__MMColParam, "'", "''") + "'  ORDER BY tb_pi.PI, tb_Acompanhamento.[Data do Registro] DESC;      "
rs_lista_acomp.CursorType = 0
rs_lista_acomp.CursorLocation = 2
rs_lista_acomp.LockType = 1
rs_lista_acomp.Open()

rs_lista_acomp_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 20
Repeat1__index = 0
rs_lista_acomp_numRows = rs_lista_acomp_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style9 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; color: #FFFFFF; }
.style11 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
.style18 {font-family: Arial, Helvetica, sans-serif; font-weight: bold; font-size: 14px; }
.style20 {
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
	font-size: 18px;
	color: #333333;
}
-->
</style>
</head>

<body>
<p align="center" class="style20">ACOMPANHAMENTO</p>
<table width="784" border="1">
  <tr bgcolor="#CCCCCC">
    <td width="200"><span class="style18"><%=(rs_lista_acomp.Fields.Item("cod_predio").Value)%></span></td>
    <td width="537"><span class="style18"><%=(rs_lista_acomp.Fields.Item("Nome_Unidade").Value)%></span></td>
  </tr>
</table>
<p>&nbsp;</p>
<div align="center">
  <table border="0">
    <tr bgcolor="#666666">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><span class="style9">PI</span></td>
      <td><span class="style9">Data do Registro</span></td>
      <td><span class="style9">cod_predio</span></td>
      <td><span class="style9">Nome_Unidade</span></td>
      <td><span class="style9">Respons&aacute;vel</span></td>
      <td><span class="style9">Registro</span></td>
      <td><span class="style9">Previs&atilde;o</span></td>
      <td><span class="style9">t&eacute;rmino contratual</span></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rs_lista_acomp.EOF)) %>
      <tr bgcolor="#CCCCCC">
        <td><a href="altera_acomp.asp?cod_acompanhamento=<%=(rs_lista_acomp.Fields.Item("cod_acompanhamento").Value)%>"><img src="depto/imagens/edit.gif" width="16" height="15" border="0" /></a></td>
        <td><a href="delete_acomp.asp?cod_acompanhamento=<%=(rs_lista_acomp.Fields.Item("cod_acompanhamento").Value)%>"><img src="depto/imagens/delete.gif" width="16" height="15" border="0" /></a></td>
        <td><span class="style11"><%=(rs_lista_acomp.Fields.Item("PI").Value)%></span></td>
        <td><span class="style11"><%=(rs_lista_acomp.Fields.Item("Data do Registro").Value)%></span></td>
        <td><span class="style11"><%=(rs_lista_acomp.Fields.Item("cod_predio").Value)%></span></td>
        <td><span class="style11"><%=(rs_lista_acomp.Fields.Item("Nome_Unidade").Value)%></span></td>
        <td><span class="style11"><%=(rs_lista_acomp.Fields.Item("Responsável").Value)%></span></td>
        <td><span class="style11"><%=(rs_lista_acomp.Fields.Item("Registro").Value)%></span></td>
        <td><span class="style11"><%=(rs_lista_acomp.Fields.Item("Previsão").Value)%></span></td>
        <td><span class="style11"><%=(rs_lista_acomp.Fields.Item("término contratual").Value)%></span></td>
      </tr>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_lista_acomp.MoveNext()
Wend
%>
  </table>
</div>
</body>
</html>
<%
rs_lista_acomp.Close()
Set rs_lista_acomp = Nothing
%>
