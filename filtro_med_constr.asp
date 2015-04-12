<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim Recordset2
Dim Recordset2_numRows

Set Recordset2 = Server.CreateObject("ADODB.Recordset")
Recordset2.ActiveConnection = MM_cpf_STRING
Recordset2.Source = "SELECT tb_predio.cod_predio, [tb_predio].[cod_predio] & ' - ' & [tb_predio].[Nome_Unidade] AS Expr1  FROM tb_predio INNER JOIN tb_PI ON tb_predio.cod_predio = tb_PI.cod_predio  GROUP BY tb_predio.cod_predio, [tb_predio].[cod_predio] & ' - ' & [tb_predio].[Nome_Unidade]  ORDER BY [tb_predio].[cod_predio] & ' - ' & [tb_predio].[Nome_Unidade];  "
Recordset2.CursorType = 0
Recordset2.CursorLocation = 2
Recordset2.LockType = 1
Recordset2.Open()

Recordset2_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<!--mstheme--><link rel="stylesheet" href="spri1011-28591.css">
<meta name="Microsoft Theme" content="spring 1011">
</head>

<body>
<form id="form1" name="form1" method="post" action="filtro_exibir_med_constr.asp?cod_predio=<%=(Recordset2.Fields.Item("cod_predio").Value)%>">
  <label>
  <select name="cod_predio" id="cod_predio" style="font-size: 10px" size="1">
    <%
While (NOT Recordset2.EOF)
%>
    <option value="<%=(Recordset2.Fields.Item("cod_predio").Value)%>"><%=(Recordset2.Fields.Item("Expr1").Value)%></option>
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
  <label>
  <input type="submit" name="Submit" value="Buscar" />
</label>
</form>
</body>
</html>
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>