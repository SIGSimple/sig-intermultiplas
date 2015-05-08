<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%

Dim Recordset1__MMColParam

If (Request.QueryString("data") <> "" And Request.QueryString("data") <> "0") Then 
  Recordset1__MMColParam = Request.QueryString("data")
  dta = Split(Recordset1__MMColParam,"/")
  sql = "SELECT * FROM c_lista_rel_ultimas_ocorrencias WHERE [Data do Registro] = #" & dta(1) & "/" & dta(0) & "/" & dta(2) & "#"
End If

Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_cpf_STRING
Recordset1.Source = sql
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
<script src="//code.jquery.com/jquery-1.11.2.min.js"></script>
<script type="text/javascript">
  $(function() {
    $("form#form1").on("submit", function(e){
      if(e.currentTarget.data.value != "0")
        return true;
      else
        return false;
    });
  });
</script>
</head>

<body>
<div align="center"><span class="style19">RELAT&Oacute;RIO DAS &Uacute;LTIMAS OCORR&Ecirc;NCIAS</span>
  <span class="style19">EM <span class="style28"><%=(Recordset1__MMColParam)%></span></span></div>

<form id="form1" name="form1" method="get" action="">
  <label for="data"></label>
  <span class="style6">Selecione a Data </span>
  <label for="Submit"></label>
  <select name="data" id="data">
    <option value="0">selecione</option>
    <%
      While (NOT Recordset2.EOF)
        If Trim(Recordset2.Fields.Item("Data do Registro").Value) <> "" Then
          Response.Write "      <OPTION value='" & (Recordset2.Fields.Item("Data do Registro").Value) & "'"
          If Lcase(Recordset2.Fields.Item("Data do Registro").Value) = Lcase(Recordset1__MMColParam) then
            Response.Write "selected"
          End If
          Response.Write ">" & (Recordset2.Fields.Item("Data do Registro").Value) & "</OPTION>"
        End If
        Recordset2.MoveNext()
      Wend
      If (Recordset2.CursorType > 0) Then
        Recordset2.MoveFirst
      Else
        Recordset2.Requery
      End If
    %>
  </select>
  <input type="submit" value="Buscar" id="Submit" />
</form>
<table border="0">
  <tr bgcolor="#999999">
    <%
      If Session("MM_UserAuthorization") = 1 or Session("MM_UserAuthorization") = 4 Then
    %>
    <td><span class="style26">Editar</span></td>
    <%
      End If
    %>
    <td><span class="style26">Ver RDO</span></td>
    <td width="150" align="center"><span class="style26">Município</span></td>
    <td width="150" align="center"><span class="style26">Localidade</span></td>
    <td width="100" align="center"><span class="style26">Nº Autos</span></td>
    <td width="250" align="center"><span class="style26">Respons&aacute;vel</span></td>
    <td width="100" align="center"><span class="style26">Data</span></td>
    <td width="400" align="center"><span class="style26">A&ccedil;&atilde;o</span></td>
    <td width="200" align="center"><span class="style26">Situa&ccedil;&atilde;o Obra</span></td>
    <td width="200" align="center"><span class="style26">Situa&ccedil;&atilde;o MSST</span></td>
    <td width="200" align="center"><span class="style26">Foi Relizada a Vistoria?</span></td>
    <td width="200" align="center"><span class="style26">Data da Vistoria</span></td>
    <td width="200" align="center"><span class="style26">É Pendência?</span></td>
    <td width="200" align="center"><span class="style26">Tipo de Pendência</span></td>
    <td width="200" align="center"><span class="style26">Descrição</span></td>
    <td width="200" align="center"><span class="style26">Tipo de Registro</span></td>
  </tr>
  <% While (NOT Recordset1.EOF) %>
    <tr bgcolor="#F4F4F4">
      <%
        If Session("MM_UserAuthorization") = 1 or Session("MM_UserAuthorization") = 4 Then
      %>
      <td><a target="_blank" href="altera_acomp.asp?cod_acompanhamento=<%=(Recordset1.Fields.Item("cod_acompanhamento").Value)%>"><img src="img/edit.gif"></a></td>
      <%
        End If
      %>
      <td><a target="_blank" href="rel_rdo.asp?cod_empreendimento=<%=(Recordset1.Fields.Item("num_autos").Value)%>&data=<%=(Recordset1.Fields.Item("Data do Registro").Value)%>"><img src="img/doc.png"></a></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("municipio").Value)%></span></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("nome_empreendimento").Value)%></span></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("num_autos").Value)%></span></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("nme_responsavel").Value)%></span></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("Data do Registro").Value)%></span></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("Registro").Value)%></span></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("desc_situacao").Value)%></span></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("dsc_situacao_sso").Value)%></span></td>
      <td align="center"><span class="style22"><%=(Recordset1.Fields.Item("e_vistoria").Value)%></span></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("dt_vistoria").Value)%></span></td>
      <td align="center"><span class="style22"><%=(Recordset1.Fields.Item("e_pendencia").Value)%></span></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("dsc_tipo_pendencia").Value)%></span></td>
      <td><span class="style22"><%=(Recordset1.Fields.Item("dsc_pendencia").Value)%></span></td>
      <td align="center"><span class="style22"><%=(Recordset1.Fields.Item("dsc_tipo_registro").Value)%></span></td>
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
