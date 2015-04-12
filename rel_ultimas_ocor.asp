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
Recordset1.Source = "SELECT *  FROM cultimas_ocorrencias  WHERE dt_registro = " + Replace(Recordset1__MMColParam, "'", "''") + ""
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
<div align="center"><span class="style19">RELAT&Oacute;RIO DAS &Uacute;LTIMAS OCORR&Ecirc;NCIAS</span></div>

<form action="rel_ultimas_ocorr.asp?dt_registro=<%=(Recordset2.Fields.Item("Data do Registro").Value)%>" method="post" name="form1" target="_blank" id="form1">
  <label for="dt_registro"></label>
  <span class="style6">Selecione a Data </span>
  <label>
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
  </label>
  <label for="Submit"></label>
  <input type="submit" name="Submit" value="Buscar" id="Submit" />
</form>
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
