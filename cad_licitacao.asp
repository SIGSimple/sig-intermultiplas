<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
	Response.CharSet = "UTF-8"
	
	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	If Not IsEmpty(Request.Form) Then
		strQ = "SELECT * FROM tb_licitacao Where 1 <> 1"


		Set rs_update = Server.CreateObject("ADODB.Recordset")
			rs_update.CursorLocation = 3
			rs_update.CursorType = 0
			rs_update.LockType = 3
			rs_update.Open strQ, objCon, , , &H0001
			rs_update.Addnew()
			
			' INÍCIO CAMPOS
			rs_update("num_autos") 				= Trim(Request.Form("num_autos"))
			
			If Request.Form("cod_tipo_contratacao") <> "" Then
				rs_update("cod_tipo_contratacao") 	= Request.Form("cod_tipo_contratacao")
			End If
			
			If Request.Form("cod_financiador") <> "" Then
			rs_update("cod_financiador") 		= Request.Form("cod_financiador")
			End If

			If Request.Form("cod_modalidade") <> "" Then
				rs_update("cod_modalidade") 		= Request.Form("cod_modalidade")
			End If

			rs_update("num_edital") 			= Trim(Request.Form("num_edital"))

			If Request.Form("dta_publicacao_doe") <> "" Then
				rs_update("dta_publicacao_doe") = Request.Form("dta_publicacao_doe")
			End If
			
			If Request.Form("dta_licitacao") <> "" Then
				rs_update("dta_licitacao") 		= Request.Form("dta_licitacao")
			End If

			If Request.Form("cod_situacao_licitacao") <> "" Then
				rs_update("cod_situacao_licitacao") = Request.Form("cod_situacao_licitacao")
			End If

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
				<span class="style17">Cadastro de Licitações</span>
			</strong>
		</p>

		<form method="post" name="form1">
			<table align="center">
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Núm. Autos Licitação:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="num_autos" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Tipo de Contratação:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_tipo_contratacao">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM tb_tipo_contratacao "

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("dsc_tipo_contratacao").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("id").Value)%>"><%=(rs_combo.Fields.Item("dsc_tipo_contratacao").Value)%></option>
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
						<span class="style22">Financiador:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_financiador">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM tb_financiador_licitacao "

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("dsc_financiador").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("id").Value)%>"><%=(rs_combo.Fields.Item("dsc_financiador").Value)%></option>
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
						<span class="style22">Modalidade de Contratação:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_modalidade">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM tb_modalidade_licitacao"

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("dsc_modalidade").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("id").Value)%>"><%=(rs_combo.Fields.Item("dsc_modalidade").Value)%></option>
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
						<span class="style22">Núm. Edital:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="num_edital" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data de Publicação D.O.E.:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_publicacao_doe" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data da Licitação:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_licitacao" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Status/Situação:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_situacao_licitacao">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM tb_situacao_licitacao "

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("dsc_situacao").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("id").Value)%>"><%=(rs_combo.Fields.Item("dsc_situacao").Value)%></option>
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
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>
						<span class="style7">Núm. Autos</span>
					</td>
					<td>
						<span class="style7">Tipo de Contratação</span>
					</td>
					<td>
						<span class="style7">Financiador</span>
					</td>
					<td>
						<span class="style7">Modalidade de Contratação</span>
					</td>
					<td>
						<span class="style7">Núm. Edital</span>
					</td>
					<td>
						<span class="style7">Data de Publicação D.O.E</span>
					</td>
					<td>
						<span class="style7">Data da Licitação</span>
					</td>
					<td>
						<span class="style7">Status/Situação</span>
					</td>
					<td>&nbsp;</td>
				</tr>
				<%
					cod_empreendimento = Request.QueryString("cod_empreendimento")
					strQ = "SELECT * from c_lista_licitacoes ORDER BY id ASC"

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
						<a href="altera_licitacao.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>">
							<img src="const/imagens/edit.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<a href="delete_licitacao.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>">
							<img src="const/imagens/delete.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("id"))%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("num_autos"))%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dsc_tipo_contratacao").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dsc_financiador").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dsc_modalidade").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("num_edital").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_publicacao_doe").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_licitacao").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dsc_situacao").Value)%></span>
					</td>
					<td>
						<form id="form-upload" method="post" enctype="multipart/form-data"
							action="novo_upload.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>&folder=LICITACAO&retUrl=<%=(Request.ServerVariables("URL"))%>?<%=(Request.QueryString)%>">
							<input type="file" name="blob">
							<br/>
							<input type="submit" id="btnSubmit" value="Upload">
						</form>

						<%
							cod_convenio = rs_lista.Fields.Item("id").Value
							strF = "SELECT * FROM tb_licitacao_arquivo WHERE cod_referencia = " & cod_convenio

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
									<a href="download.asp?path=/ARQUIVOS/LICITACAO&filename=<%=(rs_lista.Fields.Item("id").Value)%>_<%=(rs_files.Fields.Item("nme_arquivo").Value)%>">
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