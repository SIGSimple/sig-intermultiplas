<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.Form("cod_situacao") <> "") Then 
  Recordset1__MMColParam = Request.Form("cod_situacao")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_cpf_STRING
Recordset1_cmd.CommandText = "SELECT * FROM c_PiSituacao WHERE cod_situacao = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 5, 1, -1, Recordset1__MMColParam) ' adDouble

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Recordset2
Dim Recordset2_numRows

Set Recordset2 = Server.CreateObject("ADODB.Recordset")
Recordset2.ActiveConnection = MM_cpf_STRING
Recordset2.Source = "SELECT tb_situacao_pi.cod_situacao, tb_situacao_pi.desc_situacao  FROM tb_situacao_pi  ORDER BY tb_situacao_pi.desc_situacao;  "
Recordset2.CursorType = 0
Recordset2.CursorLocation = 2
Recordset2.LockType = 1
Recordset2.Open()

Recordset2_numRows = 0
%>
<%
Dim Recordset3__MMColParam
Recordset3__MMColParam = "1"
If (Request.Form("cod_situacao") <> "") Then 
  Recordset3__MMColParam = Request.Form("cod_situacao")
End If
%>
<%
Dim Recordset3
Dim Recordset3_numRows

Set Recordset3 = Server.CreateObject("ADODB.Recordset")
Recordset3.ActiveConnection = MM_cpf_STRING
Recordset3.Source = "SELECT Count(tb_pi.PI) AS ContarDePI, tb_situacao_pi.cod_situacao, tb_situacao_pi.desc_situacao  FROM ((((tb_pi LEFT JOIN tb_predio ON tb_pi.cod_predio = tb_predio.cod_predio) LEFT JOIN tb_Construtora ON tb_pi.cod_construtora = tb_Construtora.cod_construtora) LEFT JOIN tb_Municipios ON tb_pi.cod_mun = tb_Municipios.cod_mun) LEFT JOIN tb_diretoria ON tb_pi.cod_diretoria = tb_diretoria.id) LEFT JOIN tb_situacao_pi ON tb_pi.cod_situacao = tb_situacao_pi.cod_situacao  WHERE tb_pi.cod_situacao = " + Replace(Recordset3__MMColParam, "'", "''") + "  GROUP BY tb_situacao_pi.cod_situacao, tb_situacao_pi.desc_situacao  ORDER BY Count(tb_pi.PI);  "
Recordset3.CursorType = 0
Recordset3.CursorLocation = 2
Recordset3.LockType = 1
Recordset3.Open()

Recordset3_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = 10
Repeat2__index = 0
Recordset3_numRows = Recordset3_numRows + Repeat2__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style9 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; color: #FFFFFF; }
.style15 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; color: #003399; }
.style19 {font-family: Arial, Helvetica, sans-serif; font-size: 9px; }
.style12 {font-family: Arial, Helvetica, sans-serif; font-size: 9px; color: #FFFFFF; }
-->
</style>
</head>

<body>
<form id="form1" name="form1" method="post" action="">
  <label for="select"></label>
  <select name="cod_situacao" id="cod_situacao">
    <option value=""></option>
    <%
While (NOT Recordset2.EOF)
%><option value="<%=(Recordset2.Fields.Item("cod_situacao").Value)%>"><%=(Recordset2.Fields.Item("desc_situacao").Value)%></option>
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
  <label for="Submit"></label>
  <input type="submit" name="Submit" value="Buscar" id="Submit" />
  <a href="PI_situacao_todos.asp" target="_blank" class="style15">Listar Todos</a>
</form>

<table width="453" border="0">
  <tr bgcolor="#CCCCCC">
    <td><span class="style15"><%=(Recordset3.Fields.Item("ContarDePI").Value)%></span></td>
    <td><span class="style15">PIs</span></td>
    <td><span class="style15">No est&aacute;gio </span></td>
    <td><span class="style15"><%=(Recordset3.Fields.Item("desc_situacao").Value)%></span></td>
  </tr>
</table>
<p>&nbsp;</p>
<table width="3208" border="0">
  <tr bgcolor="#999999" class="style19">
    <td width="30"><span class="style9">PI</span></td>
    <td width="30"><span class="style9">C&oacute;digo do Pr&eacute;dio</span></td>
    <td width="30"><span class="style9">Nome_Unidade</span></td>
    <td width="30"><span class="style9">Diretoria de Ensino </span></td>
    <td width="30"><span class="style9">Munic&iacute;pio</span></td>
    <td width="30"><span class="style9">Interven&ccedil;&atilde;o FDE</span></td>
    <td width="30"><span class="style9">Crit&eacute;rio de C&aacute;culo</span></td>
    <td width="30"><span class="style9">Crit&eacute;rio de Reajuste</span></td>
    <td width="30"><span class="style9">N&ordm; do Contrato</span></td>
    <td width="30"><span class="style9">D&iacute;gito do Contrato</span></td>
    <td width="30"><span class="style9">Data Base</span></td>
    <td width="30"><span class="style9">Data Assinatura</span></td>
    <td width="17"><span class="style9">Data OIS</span></td>
    <td width="17"><span class="style9">Data CI</span></td>
    <td width="17"><span class="style9">Data Abertura</span></td>
    <td width="17"><span class="style9">Foi Solicitado Aditamento?</span></td>
    <td width="17"><span class="style9">Prazo Contrato</span></td>
    <td width="17"><span class="style9">Prazo Aditamento</span></td>
    <td width="17"><span class="style9">Or&ccedil;amento FDE</span></td>
    <td width="30"><span class="style9">Redu&ccedil;&atilde;o</span></td>
    <td width="17"><span class="style9">Valor Contrato</span></td>
    <td width="20"><span class="style9">Valor Aditamento</span></td>
    <td width="30"><span class="style9">&Oacute;rg&atilde;o</span></td>
    <td width="20"><span class="style9">Gerenciadora Mede</span></td>
    <td width="20"><span class="style9">Interven&ccedil;&atilde;o Gerenciadora<br />
    </span></td>
    <td width="4"><span class="style9">Fator de Redu&ccedil;&atilde;o</span></td>
    <td width="4"><span class="style9">&Aacute;rea Gerenciada</span></td>
    <td width="4"><span class="style9">Data TRP</span></td>
    <td width="4"><span class="style9">&Eacute; Medi&ccedil;&atilde;o Final</span></td>
    <td width="30"><span class="style9">Data TRD</span></td>
    <td width="30"><span class="style9">Inf. Placa Obra</span></td>
    <td width="30"><span class="style9">Respons&aacute;vel</span></td>
    <td width="30"><span class="style9">Construtora</span></td>
    <td width="30"><span class="style9">Situa&ccedil;&atilde;o</span></td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) %>
    <tr bgcolor="#CCCCCC" class="style19">
      <td><span class="style19"><%=(Recordset1.Fields.Item("PI").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("cod_predio").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("Nome_Unidade").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("desc_diretoria").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("Municipios").Value)%></span></td>
      <td class="style19"><%=(Recordset1.Fields.Item("Descrição da Intervenção FDE").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Critério de Cálculo").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Critério de Reajuste").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Número do Contrato").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Dígito do Contrato").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Data Base").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Data da Assinatura").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Data da Impressão da OIS").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Data da CI").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Data da Abertura").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Foi solicitado Aditamento?").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Prazo do Contrato").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Prazo do Aditamento").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Orçamento FDE").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Redução").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Valor do Contrato").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Valor do Aditamento").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Órgão").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Gerenciadora Mede ?").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Descrição da Intervenção Gerenciadora").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Fator de Redução").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Área gerenciada").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Data do TRP").Value)%></td>
      <td><%=(Recordset1.Fields.Item("É Medição Final ?").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Data do TRD").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Informação da Placa de Obra").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Responsável").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Construtora").Value)%></td>
      <td><%=(Recordset1.Fields.Item("desc_situacao").Value)%></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
</table>

<p>&nbsp;</p>
<p>&nbsp;</p>
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
<%
Recordset3.Close()
Set Recordset3 = Nothing
%>
