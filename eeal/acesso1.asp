<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
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
If (Request.Form("cod_situacao")  <> "") Then 
  Recordset1__MMColParam1 = Request.Form("cod_situacao") 
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_cpf_STRING
Recordset1.Source = "SELECT *  FROM tb_pi  WHERE cod_fiscal = " + Replace(Recordset1__MMColParam, "'", "''") + " and cod_situacao = " + Replace(Recordset1__MMColParam1, "'", "''") + ""
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
Recordset2.Source = "SELECT * FROM tb_responsavel"
Recordset2.CursorType = 0
Recordset2.CursorLocation = 2
Recordset2.LockType = 1
Recordset2.Open()

Recordset2_numRows = 0
%>
<%
Dim Recordset3
Dim Recordset3_numRows

Set Recordset3 = Server.CreateObject("ADODB.Recordset")
Recordset3.ActiveConnection = MM_cpf_STRING
Recordset3.Source = "SELECT *  FROM tb_situacao_pi  ORDER BY 2"
Recordset3.CursorType = 0
Recordset3.CursorLocation = 2
Recordset3.LockType = 1
Recordset3.Open()

Recordset3_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10000
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
.style3 {font-family: Arial, Helvetica, sans-serif; font-size: 9px; }
-->
</style>
</head>

<body>
<form id="form1" name="form1" method="post" action="">
  <label for="select"></label>
  <label for="label"></label>
  <label for="Submit"></label>
  <select name="cod_fiscal" id="cod_fiscal">
    <option value=""></option>
    <%
While (NOT Recordset2.EOF)
%>
    <option value="<%=(Recordset2.Fields.Item("cod_fiscal").Value)%>"><%=(Recordset2.Fields.Item("Responsável").Value)%></option>
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
  <p>
    <label>
    <input <%If (CStr((Recordset3.Fields.Item("cod_situacao").Value)) = CStr("1")) Then Response.Write("checked=""checked""") : Response.Write("")%> type="radio" name="cod_situacao" value="1" />
    <span class="style3">AGUARDA ABERTURA</span></label>
    <span class="style3"><br />
    <label>
<input <%If (CStr((Recordset3.Fields.Item("cod_situacao").Value)) = CStr("3")) Then Response.Write("checked=""checked""") : Response.Write("")%> type="radio" name="cod_situacao" value="3" />      
EM EXECUÇÃO</label>
    <br />
    <label>
      <input <%If (CStr((Recordset3.Fields.Item("cod_situacao").Value)) = CStr("4")) Then Response.Write("checked=""checked""") : Response.Write("")%> type="radio" name="cod_situacao" value="4" />
      CONCLUIDA</label>
    <br />
    <label>
      <input <%If (CStr((Recordset3.Fields.Item("cod_situacao").Value)) = CStr("5")) Then Response.Write("checked=""checked""") : Response.Write("")%> type="radio" name="cod_situacao" value="5" />
      DEVOLVIDA</label>
    <br />
    <label>
      <input <%If (CStr((Recordset3.Fields.Item("cod_situacao").Value)) = CStr("6")) Then Response.Write("checked=""checked""") : Response.Write("")%> type="radio" name="cod_situacao" value="6" />
      RESCINDIDA</label>
    <br />
    <label>
      <input <%If (CStr((Recordset3.Fields.Item("cod_situacao").Value)) = CStr("0")) Then Response.Write("checked=""checked""") : Response.Write("")%> type="radio" name="cod_situacao" value="0" />
      TODAS</label>
    </span><br />
  </p>
</form>
<table border="0">
  <tr bgcolor="#CCCCCC">
    <td><span class="style3">C&oacute;digo</span></td>
    <td><span class="style3">PI</span></td>
    <td><span class="style3">cod_construtora</span></td>
    <td><span class="style3">cod_mun</span></td>
    <td><span class="style3">cod_predio</span></td>
    <td><span class="style3">cod_fiscal</span></td>
    <td><span class="style3">cod_situacao</span></td>
    <td><span class="style3">novo contrato</span></td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) %>
    <tr bgcolor="#CCCCCC">
      <td><span class="style3"><%=(Recordset1.Fields.Item("Código").Value)%></span></td>
      <td><span class="style3"><%=(Recordset1.Fields.Item("PI").Value)%></span></td>
      <td><span class="style3"><%=(Recordset1.Fields.Item("cod_construtora").Value)%></span></td>
      <td><span class="style3"><%=(Recordset1.Fields.Item("cod_mun").Value)%></span></td>
      <td><span class="style3"><%=(Recordset1.Fields.Item("cod_predio").Value)%></span></td>
      <td><span class="style3"><%=(Recordset1.Fields.Item("cod_fiscal").Value)%></span></td>
      <td><span class="style3"><%=(Recordset1.Fields.Item("cod_situacao").Value)%></span></td>
      <td><span class="style3"><%=(Recordset1.Fields.Item("novo contrato").Value)%></span></td>
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
<%
Recordset3.Close()
Set Recordset3 = Nothing
%>
