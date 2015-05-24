<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
	Response.CharSet = "UTF-8"
	
	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	If Not IsEmpty(Request.Form) Then
		strQ = "SELECT * FROM tb_convenio_aditamento Where 1 <> 1"

		Set rs_update = Server.CreateObject("ADODB.Recordset")
			rs_update.CursorLocation = 3
			rs_update.CursorType = 0
			rs_update.LockType = 3
			rs_update.Open strQ, objCon, , , &H0001
			rs_update.Addnew()
			
			flg_vazio = True

			' INÍCIO CAMPOS
			If Request.Form("cod_convenio") <> "" Then
				flg_vazio = False
				rs_update("cod_convenio")		 		= Trim(Request.Form("cod_convenio"))
			End If
			If Request.Form("num_termo_aditamento") <> "" Then
				flg_vazio = False
				rs_update("num_termo_aditamento")		= Trim(Request.Form("num_termo_aditamento"))
			End If
			If Request.Form("cod_tipo_aditamento") <> "" Then
				flg_vazio = False
				rs_update("cod_tipo_aditamento") 		= Trim(Request.Form("cod_tipo_aditamento"))
			End If
			If Request.Form("dta_assinatura") <> "" Then
				flg_vazio = False
				rs_update("dta_assinatura") 			= Trim(Request.Form("dta_assinatura"))
			End If
			If Request.Form("vlr_aditamento") <> "" Then
				flg_vazio = False
				rs_update("vlr_aditamento") 			= Replace(Trim(Request.Form("vlr_aditamento")), ",", ".")
			End If
			If Request.Form("prz_meses") <> "" Then
				flg_vazio = False
				rs_update("prz_meses") 					= Trim(Request.Form("prz_meses"))
			End If
			If Request.Form("dta_vigencia") <> "" Then
				flg_vazio = False
				rs_update("dta_vigencia") 				= Trim(Request.Form("dta_vigencia"))
			End If
			' FIM CAMPOS
			
			Response.Write flg_vazio

			If Not flg_vazio Then
				Response.Write "Entrou"
				rs_update.Update
			End If
	End If

	Dim rs
	Dim rs_numRows

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.ActiveConnection = MM_cpf_STRING
	rs.Source = "SELECT * FROM c_lista_convenios WHERE id = " + Request.QueryString("cod_convenio")
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
		<script type="text/javascript">
			$(document).ready(function() {
				$(".datepicker").datepicker($.datepicker.regional["pt-BR"]);

				$("input[name=prz_meses]").on("blur", function() {
					dta_vigencia_original 	= moment($("input[name=dta_vigencia_original]").val(), "DD/MM/YYYY");
					prz_meses 				= $("input[name=prz_meses]").val()
					$("input[name=dta_vigencia]").val(dta_vigencia_original.add(prz_meses, "M").format("DD/MM/YYYY"));
				});

				var vlr_lines = $(".vlr");
				$.each(vlr_lines, function(i, item){
					$(item).html($.number($(item).html(), 2, ",", "."));
				});
			});
		</script>
	</head>

	<body>
		<p align="center">
			<strong>
				<span class="style17">Aditamentos de Convênio</span>
			</strong>
		</p>

		<form method="post" name="form1">
			<input type="hidden" name="cod_convenio" value="<%=(Request.QueryString("cod_convenio"))%>">
			<input type="hidden" name="dta_vigencia_original" value="<%=(rs.Fields.Item("dta_vigencia_total").Value)%>">
			<table align="center">
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Convênio:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<span class="style22"><%=(rs.Fields.Item("num_convenio").Value)%></span>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Vigência Até:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<span class="style22"><%=(rs.Fields.Item("dta_vigencia_total").Value)%></span>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Núm. Termo:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="num_termo_aditamento" value="" size="32">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Tipo de Aditamento:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_tipo_aditamento">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM tb_tipo_aditamento "

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("dsc_tipo_aditamento").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("id").Value)%>"><%=(rs_combo.Fields.Item("dsc_tipo_aditamento").Value)%></option>
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
						<span class="style22">Data Assinatura:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_assinatura" value="" size="32">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Valor Aditamento:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="vlr_aditamento" value="" size="32">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Prazo (Meses):</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="prz_meses" value="" size="32">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Vigência até:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_vigencia" value="" size="32">
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
					<!-- <td>&nbsp;</td> -->
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>
						<span class="style7">Núm. Termo</span>
					</td>
					<td>
						<span class="style7">Tipo de Aditamento</span>
					</td>
					<td>
						<span class="style7">Data Assinatura</span>
					</td>
					<td>
						<span class="style7">Valor Aditamento</span>
					</td>
					<td>
						<span class="style7">Prazo (Meses)</span>
					</td>
					<td>
						<span class="style7">Vigência até</span>
					</td>
					<td align="center">
						<span class="style7">Upload de Arquivos (Máx. 2MB)</span>
					</td>
				</tr>
				<%
					cod_convenio = Request.QueryString("cod_convenio")
					strQ = "SELECT tb_convenio_aditamento.*, tb_tipo_aditamento.dsc_tipo_aditamento FROM tb_tipo_aditamento RIGHT JOIN tb_convenio_aditamento ON tb_tipo_aditamento.id = tb_convenio_aditamento.cod_tipo_aditamento WHERE tb_convenio_aditamento.cod_convenio = " & cod_convenio

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
					</td>-->
					<td>
						<a href="delete_convenio_aditamento.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>&cod_convenio=<%=(cod_convenio)%>">
							<img src="const/imagens/delete.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<a href="cad_convenio_aditamento_nota.asp?cod_convenio_aditamento=<%=(rs_lista.Fields.Item("id").Value)%>&num_termo_aditamento=<%=(rs_lista.Fields.Item("num_termo_aditamento").Value)%>">
							<span class="style5">
								Notas
							</span>
						</a>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("num_termo_aditamento").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dsc_tipo_aditamento").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_assinatura").Value)%></span>
					</td>
					<td>
						<span class="vlr style5">
							<%
								If rs_lista.Fields.Item("vlr_aditamento").Value Then
									Response.Write Replace(rs_lista.Fields.Item("vlr_aditamento").Value, ",", ".")
								End If
							%>
						</span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("prz_meses").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_vigencia").Value)%></span>
					</td>
					<td>
						<form id="form-upload" method="post" enctype="multipart/form-data"
							action="novo_upload.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>&folder=ADITAMENTO&retUrl=<%=(Request.ServerVariables("URL"))%>?<%=(Request.QueryString)%>">
							<input type="file" name="blob">
							<br/>
							<input type="submit" id="btnSubmit" value="Upload">
						</form>

						<%
							cod_convenio_aditamento = rs_lista.Fields.Item("id").Value
							strF = "SELECT * FROM tb_convenio_aditamento_arquivo WHERE cod_referencia = " & cod_convenio_aditamento

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
									<a href="download.asp?path=/ARQUIVOS/ADITAMENTO&filename=<%=(rs_lista.Fields.Item("id").Value)%>_<%=(rs_files.Fields.Item("nme_arquivo").Value)%>">
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