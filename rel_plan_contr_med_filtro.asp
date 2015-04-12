<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.Form("cod_fiscal") <> "") Then 
  Recordset1__MMColParam = Request.Form("cod_fiscal")
End If
%>
<%
Dim Recordset1__MMColParam1
Recordset1__MMColParam1 = "1"
If (Request.Form("cod_situacao") <> "") Then 
  Recordset1__MMColParam1 = Request.Form("cod_situacao")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_cpf_STRING
Recordset1_cmd.CommandText = "SELECT * FROM cRelPlanmedicao WHERE cod_fiscal = ? and cod_situacao = ? order by cod_predio,pi" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 5, 1, -1, Recordset1__MMColParam) ' adDouble
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param2", 5, 1, -1, Recordset1__MMColParam1) ' adDouble

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 9px;
}
.style11 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10px;
	color: #FFFFFF;
}
.style13 {
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style15 {font-family: Arial, Helvetica, sans-serif; font-size: 8px; }
.style17 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; color: #FFFFFF; }
.style19 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
-->
</style>
</head>

<body>
<p align="center" class="style13"><U>RELATÓRIO DE PLANEJAMENTO E CONTROLE DAS MEDIÇÕES</U></p>
<table border="0">
  <tr bgcolor="#999999">
    <td><span class="style17">PI</span></td>
    <td><span class="style17">cod_predio</span></td>
    <td><span class="style17">Nome_Unidade</span></td>
    <td><span class="style17">Municipio</span></td>
    <td class="style11">Tipo de obra gerenciadora</td>
    <td><span class="style17">Órgão</span></td>
    <td><span class="style17">Fiscal</span></td>
    <td><span class="style17">Data de Abertura</span></td>
    <td><span class="style17">Situação</span></td>
    <td><span class="style17">Última Medição</span></td>
    <td><span class="style17">Término Contratual</span></td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) %>
    <tr bgcolor="#F3F3F3" class="style19">
      <td><span class="style19"><%=(Recordset1.Fields.Item("PI").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("cod_predio").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("Nome_Unidade").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("Municipios").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("Descrição da Intervenção Gerenciadora").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("Órgão").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("Responsável").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("Data da Abertura").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("desc_situacao").Value)%></span></td>
      <td><%=(Recordset1.Fields.Item("ultima_medicao").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Término Contratual").Value)%></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
