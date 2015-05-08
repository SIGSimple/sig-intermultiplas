<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
	Response.CharSet = "UTF-8"
	
	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	If Not IsEmpty(Request.Form) Then
		strQ = "SELECT * FROM tb_contrato_item Where 1 <> 1"

		Set rs_update = Server.CreateObject("ADODB.Recordset")
			rs_update.CursorLocation = 3
			rs_update.CursorType = 0
			rs_update.LockType = 3
			rs_update.Open strQ, objCon, , , &H0001
			rs_update.Addnew()
			
			' INÍCIO CAMPOS
			rs_update("cod_contrato") 	= Trim(Request.Form("cod_contrato"))
			rs_update("cod_item") 		= Trim(Request.Form("cod_item"))
			rs_update("dsc_item") 		= Trim(Request.Form("dsc_item"))
			rs_update("flg_reajuste") 	= Request.Form("flg_reajuste")
			' FIM CAMPOS
			
			rs_update.Update
	End If

	Dim rs
	Dim rs_numRows

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.ActiveConnection = MM_cpf_STRING
	rs.Source = "SELECT * FROM tb_contrato_item WHERE cod_contrato = " & Request.QueryString("cod_contrato")
	rs.CursorType = 0
	rs.CursorLocation = 2
	rs.LockType = 1
	rs.Open()

	rs_numRows = 0
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
		<title>Untitled Document</title>
		<style type="text/css">
			<!--
				.style5 {font-family: Arial, Helvetica, sans-serif; font-size: 12px;}
				.style7 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold;}
				.style17 {font-family: Arial, Helvetica, sans-serif; font-size: 16px;}
				.style22 {font-family: Arial, Helvetica, sans-serif; font-size: 9;}
				.style23 {font-size: 9}
			-->
		</style>
		<link rel="stylesheet" href="//code.jquery.com/ui/1.11.3/themes/smoothness/jquery-ui.css">
		<script src="//code.jquery.com/jquery-1.11.2.min.js"></script>
		<script type="text/javascript" src="//code.jquery.com/ui/1.11.4/jquery-ui.min.js"></script>
		<script type="text/javascript" src="js/datepicker-pt-BR.js"></script>
		<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.9.0/moment.min.js"></script>
		<script type="text/javascript" src="js/jquery.number.min.js"></script>
	</head>

	<body>
		<p align="center">
			<strong>
				<span class="style17">Itens do Contrato</span>
			</strong>
		</p>

		<form method="post" name="form1">
			<input type="hidden" name="cod_contrato" value="<%=(Request.QueryString("cod_contrato"))%>">
			<table align="center">
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Nº Autos - Contrato:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<span class="style22"><%=(Request.QueryString("num_autos"))%></span>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Código do Item:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="cod_item" value="" size="32">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Descrição do Item:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="dsc_item" value="" size="32">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Item de Reajuste?</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input name="flg_reajuste" type="checkbox" value="1" />
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">&nbsp;</td>
					<td bgcolor="#CCCCCC">
						<input type="submit" value="Salvar">
					</td>
				</tr>
			</table>
		</form>

		<div align="center">
			<table border="0">
				<tr bgcolor="#999999">
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>
						<span class="style7">Código do Item</span>
					</td>
					<td>
						<span class="style7">Descrição do Item</span>
					</td>
					<td>
						<span class="style7">Item de Reajuste?</span>
					</td>
				</tr>
				<%
					cod_contrato = Request.QueryString("cod_contrato")
					strQ = "SELECT *, IIf([flg_reajuste]=True,'Sim','Não') AS reajuste FROM tb_contrato_item WHERE cod_contrato = " & cod_contrato

					Set rs_lista = Server.CreateObject("ADODB.Recordset")
						rs_lista.CursorLocation = 3
						rs_lista.CursorType = 3
						rs_lista.LockType = 1
						rs_lista.Open strQ, objCon, , , &H0001

					If Not rs_lista.EOF Then
						While Not rs_lista.EOF
				%>
				<tr bgcolor="#CCCCCC">
					<td>
						<a href="altera_contrato_item.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>&cod_contrato=<%=(cod_contrato)%>">
							<img src="const/imagens/edit.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<a href="delete_contrato_item.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>&cod_contrato=<%=(cod_contrato)%>">
							<img src="const/imagens/delete.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("cod_item").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dsc_item").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("reajuste").Value)%></span>
					</td>
				</tr>
				<%
							rs_lista.MoveNext
						Wend
					End If
				%>
			</table>
		</div>
	</body>
</html>