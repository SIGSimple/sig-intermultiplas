<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
	Response.CharSet = "UTF-8"
	
	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	If Not IsEmpty(Request.Form) Then
		strQ = "SELECT * FROM tb_licenca_ambiental Where 1 <> 1"

		Set rs_update = Server.CreateObject("ADODB.Recordset")
			rs_update.CursorLocation = 3
			rs_update.CursorType = 0
			rs_update.LockType = 3
			rs_update.Open strQ, objCon, , , &H0001
			rs_update.Addnew()
			
			' INÍCIO CAMPOS
			rs_update("cod_empreendimento") = Trim(Request.Form("cod_empreendimento"))
			rs_update("num_licenca") 		= Trim(Request.Form("num_licenca"))
			rs_update("cod_tipo_licenca") 	= Trim(Request.Form("cod_tipo_licenca"))
			rs_update("dta_concessao") 		= Trim(Request.Form("dta_concessao"))
			rs_update("dta_vencimento") 	= Trim(Request.Form("dta_vencimento"))
			rs_update("dsc_observacoes") 	= Trim(Request.Form("dsc_observacoes"))
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
		<script type="text/javascript" src="js/datepicker-pt-BR.js"></script>
		<script type="text/javascript">
			$(function() {
				$(".datepicker").datepicker($.datepicker.regional["pt-BR"]);
			});
		</script>
	</head>

	<body>
		<p align="center">
			<strong>
				<span class="style17">Cadastro de Licenças Ambientais</span>
			</strong>
		</p>

		<form method="post" name="form1">
			<input type="hidden" name="cod_empreendimento" value="<%=(Request.QueryString("cod_empreendimento"))%>"/>
			<table align="center">
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Município:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<%=(Request.QueryString("nme_municipio"))%>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Empreendimento:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<%=(Request.QueryString("nme_empreendimento"))%>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Núm. Licença:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="num_licenca" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Tipo de Licença:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_tipo_licenca">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM tb_tipo_licenca "

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("dsc_tipo_licenca").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("id").Value)%>"><%=(rs_combo.Fields.Item("dsc_tipo_licenca").Value)%></option>
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
						<span class="style22">Data de Concessão:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_concessao" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data de Vencimento:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_vencimento" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Observações:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<textarea name="dsc_observacoes" cols="25"></textarea>
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
					<!-- <td>&nbsp;</td>
					<td>&nbsp;</td> -->
					<td>
						<span class="style7">Núm. Licença</span>
					</td>
					<td>
						<span class="style7">Tipo de Licença</span>
					</td>
					<td>
						<span class="style7">Data de Concessão</span>
					</td>
					<td>
						<span class="style7">Data de Vencimento</span>
					</td>
					<td>
						<span class="style7">Observações</span>
					</td>
					<td>&nbsp;</td>
				</tr>
				<%
					cod_empreendimento = Request.QueryString("cod_empreendimento")
					strQ = "select * from tb_licenca_ambiental INNER JOIN tb_tipo_licenca ON tb_licenca_ambiental.cod_tipo_licenca = tb_tipo_licenca.id where cod_empreendimento = " & cod_empreendimento

					Set rs_lista = Server.CreateObject("ADODB.Recordset")
						rs_lista.CursorLocation = 3
						rs_lista.CursorType = 3
						rs_lista.LockType = 1
						rs_lista.Open strQ, objCon, , , &H0001

					If Not rs_lista.EOF Then
						While Not rs_lista.EOF
				%>
				<tr bgcolor="#CCCCCC">
					<!-- <td>
						<a href="altera_convenio.asp?cod_convenio=cod_convenio">
							<img src="const/imagens/edit.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<a href="del_convenio.asp?cod_convenio=cod_convenio">
							<img src="const/imagens/delete.gif" width="16" height="15" border="0" />
						</a>
					</td> -->
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("num_licenca").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dsc_tipo_licenca").Value)%></span>
					</td>
					<td align="center">
						<span class="style5"><%=(rs_lista.Fields.Item("dta_concessao").Value)%></span>
					</td>
					<td align="center">
						<span class="style5"><%=(rs_lista.Fields.Item("dta_vencimento").Value)%></span>
					</td>
					<td align="center">
						<span class="style5"><%=(rs_lista.Fields.Item("dsc_observacoes").Value)%></span>
					</td>
					<td>
						<form id="form-upload" method="post" enctype="multipart/form-data"
							action="novo_upload.asp?id=<%=(rs_lista.Fields.Item("tb_licenca_ambiental.id").Value)%>&folder=LICENCA&retUrl=<%=(Request.ServerVariables("URL"))%>?<%=(Request.QueryString)%>">
							<input type="file" name="blob">
							<br/>
							<input type="submit" id="btnSubmit" value="Upload">
						</form>
						
						<%
							cod_convenio = rs_lista.Fields.Item("tb_licenca_ambiental.id").Value
							strF = "SELECT * FROM tb_licenca_ambiental_arquivo WHERE cod_referencia = " & cod_convenio

							Set rs_files = Server.CreateObject("ADODB.Recordset")
								rs_files.CursorLocation = 3
								rs_files.CursorType = 3
								rs_files.LockType = 1
								rs_files.Open strF, objCon, , , &H0001

							If Not rs_files.EOF Then
								While Not rs_files.EOF
						%>
							<ul>
								<li>
									<a href="download.asp?path=<%=(rs_files.Fields.Item("pth_arquivo").Value)%>&filename=<%=(rs_lista.Fields.Item("tb_licenca_ambiental.id").Value)%>_<%=(rs_files.Fields.Item("nme_arquivo").Value)%>">
										<%=(rs_files.Fields.Item("nme_arquivo").Value)%>
									</a>
								</li>
							</ul>
						<%
									rs_files.MoveNext
								Wend
							End If
						%>
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