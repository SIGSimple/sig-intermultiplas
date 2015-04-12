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
			rs_update("cod_empreendimento") 					= Trim(Request.Form("cod_empreendimento"))
			rs_update("cod_bacia_hidrografica") 				= Trim(Request.Form("cod_bacia_hidrografica"))
			rs_update("cod_manancial_lancamento") 				= Trim(Request.Form("cod_manancial_lancamento"))
			rs_update("qtd_metragem_coletor_tronco") 			= Trim(Request.Form("qtd_metragem_coletor_tronco"))
			rs_update("qtd_metragem_interceptor") 				= Trim(Request.Form("qtd_metragem_interceptor"))
			rs_update("qtd_metragem_emissario_fluente_bruto") 	= Trim(Request.Form("qtd_metragem_emissario_fluente_bruto"))
			rs_update("qtd_eee") 								= Trim(Request.Form("qtd_eee"))
			rs_update("qtd_metragem_linha_recalque") 			= Trim(Request.Form("qtd_metragem_linha_recalque"))
			rs_update("cod_tipo_ete") 							= Trim(Request.Form("cod_tipo_ete"))
			rs_update("dsc_estacao_tratamento") 				= Trim(Request.Form("dsc_estacao_tratamento"))
			rs_update("qtd_metragem_emissario_efluente_tratado")= Trim(Request.Form("qtd_metragem_emissario_efluente_tratado"))
			rs_update("flg_estudo_elaborado_daee") 				= Trim(Request.Form("flg_estudo_elaborado_daee"))
			rs_update("dsc_observacoes") 						= Trim(Request.Form("dsc_observacoes"))
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
		<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.9.0/moment.min.js"></script>
		<script type="text/javascript">
			$(document).ready(function() {
				$(".datepicker").datepicker($.datepicker.regional["pt-BR"]);
			});
		</script>
	</head>

	<body>
		<p align="center">
			<strong>
				<span class="style17">Registro de Dados de Empreendimento</span>
			</strong>
		</p>

		<form method="post" name="form1">
			<input type="hidden" name="cod_empreendimento" value="<%=(Request.QueryString("cod_empreendimento"))%>">
			<table align="center">
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Bacia Hidrográfica:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_bacia_hidrografica">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM tb_bacia_hidrografica "

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("nme_bacia_hidrografica").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("id").Value)%>"><%=(rs_combo.Fields.Item("nme_bacia_hidrografica").Value)%></option>
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
						<span class="style22">Manancial de Lançamento:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_manancial_lancamento">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM tb_manancial_lancamento"

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("nme_manancial").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("id").Value)%>"><%=(rs_combo.Fields.Item("nme_manancial").Value)%></option>
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
						<span class="style22">Coletor Tronco (m):</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="qtd_metragem_coletor_tronco" value="" size="32">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Interceptor (m):</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="qtd_metragem_interceptor" value="" size="32"/>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Emissário de Efluente Bruto (m):</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="qtd_metragem_emissario_fluente_bruto" value="" size="32"/>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Estação Elevatória de Esgoto (und):</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="qtd_eee" value="" size="32">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Linha de Recalque (m):</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="qtd_metragem_linha_recalque" value="" size="32">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Tipo de ETE:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_tipo_ete">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM tb_tipo_ete "

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("nme_tipo_ete").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("id").Value)%>"><%=(rs_combo.Fields.Item("nme_tipo_ete").Value)%></option>
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
						<span class="style22">Estação de Tratamento:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<textarea name="dsc_estacao_tratamento" cols="25" style="width: 98%;"></textarea>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Emissário de Efluente Tratado (m):</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="qtd_metragem_emissario_efluente_tratado" value="" size="32">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Estudo Elaborado pelo DAEE?:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input name="flg_estudo_elaborado_daee" type="checkbox" />
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
					<!-- <td>&nbsp;</td>
					<td>&nbsp;</td> -->
					<td>
						<span class="style7">Município</span>
					</td>
					<td>
						<span class="style7">Localidade</span>
					</td>
					<td>
						<span class="style7">Autos</span>
					</td>
					<td>
						<span class="style7">Bacia DAEE</span>
					</td>
					<td>
						<span class="style7">Objeto da Obra</span>
					</td>
					<td>
						<span class="style7">Bacia Hidrografica</span>
					</td>
					<td>
						<span class="style7">Manancial de Lançamento</span>
					</td>
					<td>
						<span class="style7">Coletor Tronco (m)</span>
					</td>
					<td>
						<span class="style7">Interceptor (m)</span>
					</td>
					<td>
						<span class="style7">Emissário de Efluente Bruto (m)</span>
					</td>
					<td>
						<span class="style7">Estação Elevatória de Esgoto (und)</span>
					</td>
					<td>
						<span class="style7">Linha de Recalque (m)</span>
					</td>
					<td>
						<span class="style7">Tipo de ETE</span>
					</td>
					<td>
						<span class="style7">Estação de Tratamento</span>
					</td>
					<td>
						<span class="style7">Emissário de Efluente Tratado (m)</span>
					</td>
					<td>
						<span class="style7">Estudo Elaborado pelo DAEE?</span>
					</td>
					<td>
						<span class="style7">Observações</span>
					</td>
				</tr>
				<%
					strQ = "SELECT tb_registo_dados.*, tb_bacia_hidrografica.nme_bacia_hidrografica, tb_manancial_lancamento.nme_manancial, tb_tipo_ete.nme_tipo_ete, tb_pi.nome_empreendimento, tb_predio.Município, tb_predio.[Diretoria de Ensino] AS bacia_daee, tb_pi.[Descrição da Intervenção FDE], tb_pi.PI FROM ((tb_bacia_hidrografica RIGHT JOIN (tb_manancial_lancamento RIGHT JOIN (tb_tipo_ete RIGHT JOIN tb_registo_dados ON tb_tipo_ete.id = tb_registo_dados.cod_tipo_ete) ON tb_manancial_lancamento.id = tb_registo_dados.cod_manancial_lancamento) ON tb_bacia_hidrografica.id = tb_registo_dados.cod_bacia_hidrografica) INNER JOIN tb_pi ON tb_registo_dados.cod_empreendimento = tb_pi.Código) INNER JOIN tb_predio ON tb_pi.id_predio = tb_predio.id_predio"

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
						<a href="altera_convenio.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>">
							<img src="const/imagens/edit.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<a href="delete_convenio.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>">
							<img src="const/imagens/delete.gif" width="16" height="15" border="0" />
						</a>
					</td> -->
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("nome_empreendimento").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("PI").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("bacia_daee").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("Descrição da Intervenção FDE").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("nme_manancial").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("qtd_metragem_coletor_tronco").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("qtd_metragem_interceptor").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("qtd_metragem_emissario_fluente_bruto").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("qtd_eee").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("qtd_metragem_linha_recalque").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("nme_tipo_ete").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dsc_estacao_tratamento").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("qtd_metragem_emissario_efluente_tratado").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("flg_estudo_elaborado_daee").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dsc_observacoes").Value)%></span>
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