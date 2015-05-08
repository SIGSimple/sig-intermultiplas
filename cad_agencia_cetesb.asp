<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%	
	Response.CharSet = "UTF-8"

	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	If Not IsEmpty(Request.Form) Then

		strQ = "SELECT * FROM tb_agencia_cetesb Where 1 <> 1"

		Set rs_update = Server.CreateObject("ADODB.Recordset")
			rs_update.CursorLocation = 3
			rs_update.CursorType = 0
			rs_update.LockType = 3
			rs_update.Open strQ, objCon, , , &H0001
			rs_update.Addnew()
			
			' INÍCIO CAMPOS
			rs_update("nme_agencia") 	= Trim(Request.Form("nme_agencia"))
			rs_update("dsc_endereco") 	= Trim(Request.Form("dsc_endereco"))
			rs_update("dsc_bairro") 	= Trim(Request.Form("dsc_bairro"))
			
			If(Request.Form("cod_municipio") <> "") Then
				rs_update("cod_municipio") 	= Request.Form("cod_municipio")
			End If

			rs_update("num_cep") 		= Trim(Request.Form("num_cep"))
			rs_update("num_telefone") 	= Trim(Request.Form("num_telefone"))
			rs_update("num_fax") 		= Trim(Request.Form("num_fax"))
			rs_update("end_email") 		= Trim(Request.Form("end_email"))
			' FIM CAMPOS
			
			rs_update.Update
	End If
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
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
	</head>

	<body>
		<p align="center">
			<strong>
				<span class="style17">Cadastro de Agências CETESB</span>
			</strong>
		</p>

		<form method="post" name="form1">
			<table align="center">
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Agência Ambiental de(o):</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="nme_agencia" value="" size="40">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Endereço:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="dsc_endereco" value="" size="40">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Bairro:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="dsc_bairro" value="" size="40">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Cidade/Município:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_municipio">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM c_lista_municipios"

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("nme_municipio").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("cod_municipio").Value)%>"><%=(rs_combo.Fields.Item("nme_municipio").Value)%></option>
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
						<span class="style22">CEP:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="num_cep" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Telefone:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="num_telefone" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Fax:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="num_fax" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">E-Mail:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="end_email" value="" size="10">
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
						<span class="style7">Agência Ambiental de(o)</span>
					</td>
					<td>
						<span class="style7">Endereço</span>
					</td>
					<td>
						<span class="style7">Bairro</span>
					</td>
					<td>
						<span class="style7">Cidade/Município</span>
					</td>
					<td>
						<span class="style7">CEP</span>
					</td>
					<td>
						<span class="style7">Telefone</span>
					</td>
					<td>
						<span class="style7">Fax</span>
					</td>
					<td>
						<span class="style7">E-Mail</span>
					</td>
				</tr>
				<%
					strQ = "select * from c_lista_agencias_cetesb"

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
						<a href="altera_agencia_cetesb.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>">
							<img src="const/imagens/edit.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<a href="delete_agencia_cetesb.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>">
							<img src="const/imagens/delete.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("nme_agencia").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dsc_endereco").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dsc_bairro").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("nme_municipio").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("num_cep").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("num_telefone").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("num_fax").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("end_email").Value)%></span>
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