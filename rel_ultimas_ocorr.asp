<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.Form("dt_registro") <> "") Then 
  Recordset1__MMColParam = Request.Form("dt_registro")
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_cpf_STRING
Recordset1.Source = "SELECT cultimasocorrencias.*, IIf([data da abertura] Is Null,0,IIf([data da abertura]<>0,Now()-[data da abertura],0)/[prazo do contrato]) AS [Dias Corridos X Prazo Contratual], tb_Acompanhamento.Registro, tb_Acompanhamento.n_LO, tb_Acompanhamento.dt_vistoria  FROM cultimasocorrencias INNER JOIN tb_Acompanhamento ON (cultimasocorrencias.[PI-item] = tb_Acompanhamento.PI) AND (cultimasocorrencias.dt_registro = tb_Acompanhamento.[Data do Registro])  WHERE dt_registro = " + Replace(Recordset1__MMColParam, "'", "''") + " or dt_registro like '%" + Replace(Recordset1__MMColParam, "'", "''") + "%'"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim Recordset2
Dim Recordset2_numRows

Set Recordset2 = Server.CreateObject("ADODB.Recordset")
Recordset2.ActiveConnection = MM_cpf_STRING
Recordset2.Source = "SELECT tb_Acompanhamento.[Data do Registro]  FROM tb_Acompanhamento  GROUP BY tb_Acompanhamento.[Data do Registro]  ORDER BY tb_Acompanhamento.[Data do Registro] DESC;"
Recordset2.CursorType = 0
Recordset2.CursorLocation = 2
Recordset2.LockType = 1
Recordset2.Open()

Recordset2_numRows = 0
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
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style19 {font-size: 18px; color: #333366; font-family: Arial, Helvetica, sans-serif; font-weight: bold; }
.style22 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
.style26 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; color: #FFFFFF; }
.style6 {	font-family: Arial, Helvetica, sans-serif;
	font-size: 14px;
}
.style28 {
	color: #FF0000;
	font-size: 24;
}
-->
</style>
</head>

<body>
<div align="center"><span class="style19">RELAT&Oacute;RIO DAS &Uacute;LTIMAS OCORR&Ecirc;NCIAS</span>
  <span class="style19">EM <span class="style28"><%=(Recordset1.Fields.Item("dt_registro").Value)%></span></span></div>

<form id="form1" name="form1" method="post" action="">
  <label for="dt_registro"></label>
  <span class="style6">Selecione a Data </span>
  <label></label>
  <label for="Submit"></label>
  <select name="dt_registro" id="dt_registro">
    <option value="">Todos</option>
    <%
While (NOT Recordset2.EOF)
%>
    <option value="<%=(Recordset2.Fields.Item("Data do Registro").Value)%>"><%=(Recordset2.Fields.Item("Data do Registro").Value)%></option>
    <%
  Recordset2.MoveNext()
Wend
If (Recordset2.CursorType > 0) Then
  Recordset2.MoveFirst
Else
  Recordset2.Requery
End If
%>
  </select>
  <input type="submit" name="Submit" value="Buscar" id="Submit" />
  <label for="select"></label>
</form>
<table border="0">
  <tr bgcolor="#999999">
    <td width="34"><span class="style26">PI-item</span></td>
    <td width="114"><span class="style26">cod_predio</span></td>
    <td width="135"><span class="style26">Nome_Unidade</span></td>
    <td width="120"><span class="style26">Respons&aacute;vel</span></td>
    <td width="85"><span class="style26">Fiscal</span></td>
    <td width="90"><span class="style26">&Oacute;rg&atilde;o</span></td>
    <td width="110"><span class="style26">Data</span></td>
    <td width="99"><span class="style26">A&ccedil;&atilde;o</span></td>
    <td width="127"><span class="style26">Situa&ccedil;&atilde;o</span></td>
    <td width="186"><span class="style26">Previs&atilde;o de T&eacute;rmino </span></td>
    <td width="122"><span class="style26">N&ordm; do LO </span></td>
    <td width="119"><span class="style26">Data da Vistoria </span></td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) %>
    <tr bgcolor="#F4F4F4">
      <td><span class="style22"><%=(Recordset1.Fields.Item("PI-item").Value)%></span></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("cod_predio").Value)%></span></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("Nome_Unidade").Value)%></span></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("Responsável").Value)%></span></td>
      <td class="style22"><%=(Recordset1.Fields.Item("fiscal").Value)%></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("Órgão").Value)%></span></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("dt_registro").Value)%></span></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("Registro").Value)%></span></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("desc_situacao").Value)%></span></td>
      <td class="style22"><%=(Recordset1.Fields.Item("Data Prevista para o Término").Value)%></td>
      <td class="style22"><%=(Recordset1.Fields.Item("n_LO").Value)%></td>
      <td class="style22"><%=(Recordset1.Fields.Item("dt_vistoria").Value)%></td>
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
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>
