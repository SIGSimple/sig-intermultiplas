<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
	Response.CharSet = "UTF-8"
	
	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	If Not IsEmpty(Request.Form) Then
		strQ = "SELECT * FROM tb_pi_contrato Where 1 <> 1"

		Set rs_update = Server.CreateObject("ADODB.Recordset")
			rs_update.CursorLocation = 3
			rs_update.CursorType = 0
			rs_update.LockType = 3
			rs_update.Open strQ, objCon, , , &H0001
			rs_update.Addnew()
			
			' INÍCIO CAMPOS
			rs_update("cod_empreendimento") 	= Trim(Request.Form("cod_empreendimento"))
			rs_update("cod_contrato") 			= Trim(Request.Form("cod_contrato"))
			' FIM CAMPOS
			
			rs_update.Update
	End If
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
		<script type="text/javascript" src="js/jquery.number.min.js"></script>
		<script type="text/javascript" src="js/datepicker-pt-BR.js"></script>
		<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.9.0/moment.min.js"></script>
	</head>

	<body>
		<p align="center">
			<strong>
				<span class="style17">Associação Empreendimento x Contrato</span>
			</strong>
		</p>

		<form method="post" name="form1">
			<table align="center">
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Empreendimento:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_empreendimento">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM c_lista_dados_obras "

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("nome_empreendimento").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("Código").Value)%>"><%=(rs_combo.Fields.Item("municipio").Value)%> - <%=(rs_combo.Fields.Item("nome_empreendimento").Value)%></option>
							<%
										End If
										rs_combo.MoveNext
									Wend
								End If
							%>
						</select>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Contrato:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_contrato">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM c_lista_contrato "

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("num_contrato").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("id").Value)%>"><%=(rs_combo.Fields.Item("num_autos").Value)%> - <%=(rs_combo.Fields.Item("num_contrato").Value)%></option>
							<%
										End If
										rs_combo.MoveNext
									Wend
								End If
							%>
						</select>
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
					<td>
						<span class="style7">Municipio</span>
					</td>
					<td>
						<span class="style7">Localidade</span>
					</td>
					<td>
						<span class="style7">Nº Autos Empreendimento</span>
					</td>
					<td>
						<span class="style7">Nº Autos Contrato</span>
					</td>
					<td>
						<span class="style7">Nº Contrato</span>
					</td>
				</tr>
				<%
					strQ = "SELECT * FROM c_lista_contrato_empreendimento"

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
						<a href="delete_contrato_empreendimento.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>">
							<img src="const/imagens/delete.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("municipio").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("nome_empreendimento").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("PI").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("num_autos").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("num_contrato").Value)%></span>
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