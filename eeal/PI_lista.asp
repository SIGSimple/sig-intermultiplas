<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_cpf_STRING
Recordset1.Source = "SELECT tb_predio.cod_predio, tb_predio.Nome_Unidade, tb_Municipios.Municipios, tb_PI.PI, tb_PI.[Descrição da Intervenção FDE], tb_Construtora.Construtora, tb_PI.[Número do Contrato], tb_PI.[Dígito do Contrato], tb_PI.Órgão, tb_PI.[Data do TRP], tb_PI.[Data do TRD]  FROM (tb_diretoria INNER JOIN tb_predio ON tb_diretoria.cod_diretoria = tb_predio.cod_diretoria) INNER JOIN (tb_Construtora INNER JOIN (tb_PI INNER JOIN tb_Municipios ON tb_PI.cod_mun = tb_Municipios.cod_mun) ON tb_Construtora.cod_construtora = tb_PI.cod_construtora) ON tb_predio.cod_predio = tb_PI.cod_predio  ORDER BY tb_predio.cod_predio, tb_predio.Nome_Unidade;"
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
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
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
-->
</style>
</head>

<body>
<table border="0">
  <tr bgcolor="#999999">
    <td><span class="style9">cod_predio</span></td>
    <td><span class="style9">Nome_Unidade</span></td>
    <td><span class="style9">Municipios</span></td>
    <td><span class="style9">PI</span></td>
    <td><span class="style9">Descri&ccedil;&atilde;o da Interven&ccedil;&atilde;o FDE</span></td>
    <td><span class="style9">Construtora</span></td>
    <td><span class="style9">N&uacute;mero do Contrato</span></td>
    <td><span class="style9">D&iacute;gito do Contrato</span></td>
    <td><span class="style9">&Oacute;rg&atilde;o</span></td>
    <td><span class="style9">Data do TRP</span></td>
    <td><span class="style9">Data do TRD</span></td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) %>
    <tr bgcolor="#CCCCCC">
      <td><span class="style11"><%=(Recordset1.Fields.Item("cod_predio").Value)%></span></td>
      <td><span class="style11"><%=(Recordset1.Fields.Item("Nome_Unidade").Value)%></span></td>
      <td><span class="style11"><%=(Recordset1.Fields.Item("Municipios").Value)%></span></td>
      <td><span class="style11"><a href="contratos.asp?PI=<%=(Recordset1.Fields.Item("PI").Value)%>"><%=(Recordset1.Fields.Item("PI").Value)%></a></span></td>
      <td><span class="style11"><%=(Recordset1.Fields.Item("Descrição da Intervenção FDE").Value)%></span></td>
      <td><span class="style11"><%=(Recordset1.Fields.Item("Construtora").Value)%></span></td>
      <td><span class="style11"><%=(Recordset1.Fields.Item("Número do Contrato").Value)%></span></td>
      <td><span class="style11"><%=(Recordset1.Fields.Item("Dígito do Contrato").Value)%></span></td>
      <td><span class="style11"><%=(Recordset1.Fields.Item("Órgão").Value)%></span></td>
      <td><span class="style11"><%=(Recordset1.Fields.Item("Data do TRP").Value)%></span></td>
      <td><span class="style11"><%=(Recordset1.Fields.Item("Data do TRD").Value)%></span></td>
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
