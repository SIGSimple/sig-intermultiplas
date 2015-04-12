<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_cpf_STRING
Recordset1.Source = "SELECT * FROM tb_mensagem"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style7 {font-family: Arial, Helvetica, sans-serif; font-size: 18px; font-weight: bold; }
-->
</style>
</head>

<body>
<div align="center">
  <table border="0">
    <tr bgcolor="#F0F0F0">
      <td width="441"><div align="center"><span class="style7"><a href="altera_mensagem.asp?cod_mensagem=<%=(Recordset1.Fields.Item("cod_mensagem").Value)%>">Alterar Mensagem da P&aacute;gina Inicial </a></span></div></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
