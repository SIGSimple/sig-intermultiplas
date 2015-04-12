<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_cpf_STRING
Recordset1_cmd.CommandText = "SELECT * FROM tb_predio" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
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
	font-weight: bold;
	font-size: small;
}
.style2 {
	color: #CC0000;
	font-weight: bold;
	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
}
.style3 {color: #CC6600}
-->
</style>
</head>

<body>
<form id="form1" name="form1" method="post" action="rel_pi_filtro_cod.asp">
  <label>  <span class="style2">Localizar PI pelo código</span> <span class="style2">- <span class="style3">Informe o código do PI</span> ou parte dele</span><br />
  <input name="PI" type="text" id="PI" size="15" />
  </label>
  <label>
  <input type="submit" name="button" id="button" value="Buscar" />
  </label>
  <a href="http://www.cep-escolas.com.br/busca_pi_unidade.asp" class="style1"></a>
</form>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
