<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim rs_localiza
Dim rs_localiza_numRows

Set rs_localiza = Server.CreateObject("ADODB.Recordset")
rs_localiza.ActiveConnection = MM_cpf_STRING
rs_localiza.Source = "SELECT * FROM tb_predio"
rs_localiza.CursorType = 0
rs_localiza.CursorLocation = 2
rs_localiza.LockType = 1
rs_localiza.Open()

rs_localiza_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
</head>

<body>
<form id="form1" name="form1" method="post" action="PI_lista.asp?Nome da Unidade=<%=(rs_localiza.Fields.Item("Nome_Unidade").Value)%>">
  <input name="Nome da Unidade" type="text" id="Nome da Unidade" />
  <label>
  <input type="submit" name="Submit" value="Buscar" />
</label>
</form>
</body>
</html>
<%
rs_localiza.Close()
Set rs_localiza = Nothing
%>
