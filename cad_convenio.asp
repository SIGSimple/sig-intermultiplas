<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
	Response.CharSet = "UTF-8"
	
	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	If Not IsEmpty(Request.Form) Then
		strQ = "SELECT * FROM tb_convenio Where 1 <> 1"

		Set rs_update = Server.CreateObject("ADODB.Recordset")
			rs_update.CursorLocation = 3
			rs_update.CursorType = 0
			rs_update.LockType = 3
			rs_update.Open strQ, objCon, , , &H0001
			rs_update.Addnew()
			
			' INÍCIO CAMPOS
			rs_update("num_autos") 					= Trim(Request.Form("num_autos"))
			rs_update("cod_projetista_convenio") 	= Trim(Request.Form("cod_projetista_convenio"))
			rs_update("cod_enquadramento") 			= Trim(Request.Form("cod_enquadramento"))
			rs_update("cod_programa") 				= Trim(Request.Form("cod_programa"))
			rs_update("num_convenio") 				= Trim(Request.Form("num_convenio"))
			rs_update("dta_assinatura") 			= Trim(Request.Form("dta_assinatura"))
			rs_update("dta_publicacao_doe")			= Trim(Request.Form("dta_publicacao_doe"))
			rs_update("vlr_original") 				= Replace(Trim(Request.Form("vlr_original")), ",", ".")
			rs_update("prz_meses") 					= Trim(Request.Form("prz_meses"))
			rs_update("dta_vigencia") 				= Trim(Request.Form("dta_vigencia"))
			rs_update("nme_fonte_recurso") 			= Trim(Request.Form("nme_fonte_recurso"))
			rs_update("cod_coordenador_daee") 		= Trim(Request.Form("cod_coordenador_daee"))
			rs_update("dsc_observacoes")		 	= Trim(Request.Form("dsc_observacoes"))
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
		<script type="text/javascript">
			$(document).ready(function() {
				$(".datepicker").datepicker($.datepicker.regional["pt-BR"]);

				$("input[name=prz_meses]").on("blur", function() {
					dta_assinatura = moment($("input[name=dta_assinatura]").val(), "DD/MM/YYYY");
					prz_meses = $("input[name=prz_meses]").val()
					$("input[name=dta_vigencia]").val(dta_assinatura.add(prz_meses, "M").format("DD/MM/YYYY"));
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
				<span class="style17">Cadastro de Convênios</span>
			</strong>
		</p>

		<form method="post" name="form1">
			<table align="center">
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Autos:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="num_autos" value="" size="32">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Projetista:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_projetista_convenio">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM tb_Construtora ORDER BY Construtora ASC "

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("Construtora").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("cod_construtora").Value)%>"><%=(rs_combo.Fields.Item("Construtora").Value)%></option>
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
						<span class="style22">Enquadramento:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_enquadramento">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM tb_convenio_enquadramento "

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("dsc_enquadramento").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("id").Value)%>"><%=(rs_combo.Fields.Item("dsc_enquadramento").Value)%></option>
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
						<span class="style22">Programa:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_programa">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM tb_depto "

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("desc_depto").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("cod_depto").Value)%>"><%=(rs_combo.Fields.Item("sigla").Value)%> - <%=(rs_combo.Fields.Item("desc_depto").Value)%></option>
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
						<span class="style22">Núm. Convênio:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="num_convenio" value="" size="32">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data Assinatura:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_assinatura" value="" size="32"/>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data Publicação DOE:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_publicacao_doe" value="" size="32"/>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Valor Original:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="vlr_original" value="" size="32">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Prazo Original (Meses):</span>
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
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Fonte de Recurso:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="nme_fonte_recurso" value="" size="32">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Coord. DAEE:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_coordenador_daee">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM tb_responsavel "

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("Responsável").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("cod_fiscal").Value)%>"><%=(rs_combo.Fields.Item("Responsável").Value)%></option>
							<%
										End If
										rs_combo.MoveNext
									Wend
								End If
							%>
						</select>
					</td>
				</tr>
				<tr valign="middle">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Observações:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<textarea name="dsc_observacoes" cols="25" style="width: 98%;"></textarea>
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
						<span class="style7">Autos</span>
					</td>
					<td>
						<span class="style7">Projetista</span>
					</td>
					<td>
						<span class="style7">Enquadramento</span>
					</td>
					<td>
						<span class="style7">Programa</span>
					</td>
					<td>
						<span class="style7">Núm. Convênio</span>
					</td>
					<td>
						<span class="style7">Data Assinatura</span>
					</td>
					<td>
						<span class="style7">Data de Publicação DOE</span>
					</td>

					<td>
						<span class="style7">Valor Original</span>
					</td>
					<td>
						<span class="style7">Valor Total</span>
					</td>

					<td>
						<span class="style7">Prazo Original (Meses)</span>
					</td>
					<td>
						<span class="style7">Prazo Total (Meses)</span>
					</td>
					<td>
						<span class="style7">Vigência Até</span>
					</td>

					<td>
						<span class="style7">Fonte de Recurso</span>
					</td>
					<td>
						<span class="style7">Coord. DAEE</span>
					</td>
					<td>
						<span class="style7">Observações</span>
					</td>
					<td>&nbsp;</td>
					<td>
						<span class="style7">Upload de Arquivos (Máx. 4MB)</span>
					</td>
				</tr>
				<%
					strQ = "SELECT * FROM c_lista_convenios ORDER BY id DESC"

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
						<a href="altera_convenio.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>">
							<img src="const/imagens/edit.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<a href="delete_convenio.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>">
							<img src="const/imagens/delete.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("num_autos").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("projetista").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("enquadramento").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("programa").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("num_convenio").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_assinatura").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_publicacao_doe").Value)%></span>
					</td>
					
					<td>
						<span class="vlr style5"><%=(rs_lista.Fields.Item("vlr_original").Value)%></span>
					</td>
					<td>
						<span class="vlr style5"><%=(rs_lista.Fields.Item("vlr_total_aditamento").Value)%></span>
					</td>

					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("prz_meses").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("prz_total_aditamento").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_vigencia_total").Value)%></span>
					</td>

					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("nme_fonte_recurso").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("coordenador_daee").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dsc_observacoes").Value)%></span>
					</td>
					<td>
						<a href="cad_convenio_aditamento.asp?cod_convenio=<%=(rs_lista.Fields.Item("id").Value)%>&num_convenio=<%=(rs_lista.Fields.Item("num_convenio").Value)%>">
							<span class="style5">
								Aditamentos
							</span>
						</a>
					</td>
					<td>
						<form id="form-upload" method="post" enctype="multipart/form-data"
							action="novo_upload.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>&folder=CONVENIO&retUrl=<%=(Request.ServerVariables("URL"))%>?<%=(Request.QueryString)%>">
							<input type="file" name="blob">
							<br/>
							<input type="submit" id="btnSubmit" value="Upload">
						</form>

						<%
							cod_convenio = rs_lista.Fields.Item("id").Value
							strF = "SELECT * FROM tb_convenio_arquivo WHERE cod_referencia = " & cod_convenio

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
									<a href="download.asp?path=/ARQUIVOS/CONVENIO&filename=<%=(rs_lista.Fields.Item("id").Value)%>_<%=(rs_files.Fields.Item("nme_arquivo").Value)%>">
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