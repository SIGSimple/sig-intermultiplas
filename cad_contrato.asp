<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
	Response.CharSet = "UTF-8"
	
	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	If Not IsEmpty(Request.Form) Then
		strQ = "SELECT * FROM tb_contrato Where 1 <> 1"

		Set rs_update = Server.CreateObject("ADODB.Recordset")
			rs_update.CursorLocation = 3
			rs_update.CursorType = 0
			rs_update.LockType = 3
			rs_update.Open strQ, objCon, , , &H0001
			rs_update.Addnew()
			
			' INÍCIO CAMPOS
			rs_update("cod_empresa_contratada") 			= Trim(Request.Form("cod_empresa_contratada"))
			rs_update("cod_engenheiro_empresa_contratada") 	= Trim(Request.Form("cod_engenheiro_empresa_contratada"))
			rs_update("num_autos") 							= Trim(Request.Form("num_autos"))
			rs_update("num_contrato") 						= Trim(Request.Form("num_contrato"))
			rs_update("dta_assinatura")						= Trim(Request.Form("dta_assinatura"))
			rs_update("dta_publicacao_doe") 				= Trim(Request.Form("dta_publicacao_doe"))
			rs_update("dta_pedido_empenho")					= Trim(Request.Form("dta_pedido_empenho"))
			rs_update("dta_os") 							= Trim(Request.Form("dta_os"))
			rs_update("prz_original_contrato_meses") 		= Trim(Request.Form("prz_original_contrato_meses"))
			rs_update("prz_aditivos_contrato_meses") 		= Trim(Request.Form("prz_aditivos_contrato_meses"))
			rs_update("prz_original_execucao_meses") 		= Trim(Request.Form("prz_original_execucao_meses"))
			rs_update("dta_vigencia")		 				= Trim(Request.Form("dta_vigencia"))
			rs_update("dta_base") 							= Trim(Request.Form("dta_base"))
			rs_update("dta_inauguracao") 					= Trim(Request.Form("dta_inauguracao"))
			rs_update("dta_termo_recebimento_provisorio") 	= Trim(Request.Form("dta_termo_recebimento_provisorio"))
			rs_update("dta_termo_recebimento_definitivo") 	= Trim(Request.Form("dta_termo_recebimento_definitivo"))
			rs_update("dta_encerramento_contrato") 			= Trim(Request.Form("dta_encerramento_contrato"))
			rs_update("dta_recisao_contratual") 			= Trim(Request.Form("dta_recisao_contratual"))
			rs_update("cod_contratante") 					= Trim(Request.Form("cod_contratante"))
			rs_update("cod_situacao") 						= Trim(Request.Form("cod_situacao"))
			rs_update("cod_licitacao") 						= Trim(Request.Form("cod_licitacao"))
			rs_update("vlr_original_contrato") 				= Replace(Trim(Request.Form("vlr_original_contrato")), ",", ".")
			rs_update("vlr_aditivos_contrato") 				= Replace(Trim(Request.Form("vlr_aditivos_contrato")), ",", ".")
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

				$("input[name=prz_original_contrato_meses]").on("blur", function() {
					dta_assinatura = moment($("input[name=dta_assinatura]").val(), "DD/MM/YYYY");
					prz_meses = $("input[name=prz_original_contrato_meses]").val()
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
				<span class="style17">Cadastro/Acompanhamento de Contratos</span>
			</strong>
		</p>

		<form method="post" name="form1">
			<table align="center">
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Empresa Contratada:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_empresa_contratada">
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
						<span class="style22">Engenheiro Responsável (Emp. Contratada):</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_engenheiro_empresa_contratada">
							<option value=""></option>
							<%
								strQ = "SELECT *  FROM c_lista_fiscal "

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("nme_interessado").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("cod_fiscal").Value)%>"><%=(rs_combo.Fields.Item("nme_interessado").Value)%></option>
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
						<span class="style22">Licitação:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_licitacao">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM tb_licitacao ORDER BY num_autos ASC"

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
							%>
							<option value="<%=(rs_combo.Fields.Item("id").Value)%>">Nº Autos: <%=(rs_combo.Fields.Item("num_autos").Value)%> - Nº Edital: <%=(rs_combo.Fields.Item("num_edital").Value)%></option>
							<%
										rs_combo.MoveNext
									Wend
								End If
							%>
						</select>
					</td>
				</tr>

				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Nº Autos:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="num_autos" value="" size="32">
					</td>
				</tr>

				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Nº Contrato:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="num_contrato" value="" size="32">
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
						<span class="style22">Data Pedido Empenho:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_pedido_empenho" value="" size="32"/>
					</td>
				</tr>

				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data Ordem de Serviço:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_os" value="" size="32"/>
					</td>
				</tr>

				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Prazo Original (Meses):</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="prz_original_contrato_meses" value="" size="32">
					</td>
				</tr>

				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Valor Original</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="vlr_original_contrato" value="" size="32">
					</td>
				</tr>

				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Prazo Aditivos (Meses):</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="prz_aditivos_contrato_meses" value="" size="32">
					</td>
				</tr>

				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Valor Total dos Aditivos:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="vlr_aditivos_contrato" value="" size="32">
					</td>
				</tr>

				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Prazo de Execução (Meses):</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="prz_original_execucao_meses" value="" size="32">
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
						<span class="style22">Data Base:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_base" value="" size="32">
					</td>
				</tr>

				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data Inauguração:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_inauguracao" value="" size="32">
					</td>
				</tr>

				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data Rec. Termo Provisório:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_termo_recebimento_provisorio" value="" size="32">
					</td>
				</tr>

				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data Rec. Termo Definitivo:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_termo_recebimento_definitivo" value="" size="32">
					</td>
				</tr>

				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data Encerramento Contrato:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_encerramento_contrato" value="" size="32">
					</td>
				</tr>

				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data Recisão Contratual:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_recisao_contratual" value="" size="32">
					</td>
				</tr>

				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Contratante:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_contratante">
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
						<span class="style22">Situação:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_situacao">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM tb_situacao_pi ORDER BY desc_situacao ASC"

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("desc_situacao").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("cod_situacao").Value)%>">Status: <%=(rs_combo.Fields.Item("desc_situacao").Value)%> - Situação: <%=(rs_combo.Fields.Item("cod_atendimento").Value)%></option>
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
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>
						<span class="style7">#</span>
					</td>
					<td style="min-width: 200px; text-align: center;">
						<span class="style7">Município</span>
					</td>
					<td style="min-width: 200px; text-align: center;">
						<span class="style7">Localidade</span>
					</td>
					<td style="min-width: 200px; text-align: center;">
						<span class="style7">Empresa Contratatada</span>
					</td>
					<td style="min-width: 200px; text-align: center;">
						<span class="style7">Engenheiro Responsável</span>
					</td>
					<td style="min-width: 100px; text-align: center;">
						<span class="style7">Núm. Autos</span>
					</td>
					<td style="min-width: 100px; text-align: center;">
						<span class="style7">Núm. Contrato</span>
					</td>
					<td style="min-width: 150px; text-align: center;">
						<span class="style7">Data Assinatura</span>
					</td>
					<td style="min-width: 150px; text-align: center;">
						<span class="style7">Data de Publicação DOE</span>
					</td>
					<td style="min-width: 150px; text-align: center;">
						<span class="style7">Data do Pedido Empenho</span>
					</td>
					<td style="min-width: 150px; text-align: center;">
						<span class="style7">Data da Ordem de Serviço</span>
					</td>
					<td style="min-width: 100px; text-align: center;">
						<span class="style7">Prazo Original (Meses)</span>
					</td>
					<td style="min-width: 100px; text-align: center;">
						<span class="style7">Prazo Execução (Meses)</span>
					</td>
					<td style="min-width: 150px; text-align: center;">
						<span class="style7">Vigência Até</span>
					</td>
					<td style="min-width: 150px; text-align: center;">
						<span class="style7">Data Base</span>
					</td>
					<td style="min-width: 150px; text-align: center;">
						<span class="style7">Data de Inauguração</span>
					</td>
					<td style="min-width: 150px; text-align: center;">
						<span class="style7">Data Rec. Termo Provisório</span>
					</td>
					<td style="min-width: 150px; text-align: center;">
						<span class="style7">Data Rec. Termo Definitivo</span>
					</td>
					<td style="min-width: 150px; text-align: center;">
						<span class="style7">Data Encerramento Contrato</span>
					</td>
					<td style="min-width: 150px; text-align: center;">
						<span class="style7">Data Recisão Contratual</span>
					</td>
					<td style="min-width: 200px; text-align: center;">
						<span class="style7">Contratante</span>
					</td>
					<td style="min-width: 200px; text-align: center;">
						<span class="style7">Situação</span>
					</td>
					<td>
						<span class="style7">Upload de Arquivos (Máx. 4MB)</span>
					</td>
				</tr>
				<%
					strQ = "SELECT tb_pi_contrato.id, c_lista_dados_obras.Código as cod_empreendimento, c_lista_dados_obras.*, c_lista_contrato.id as cod_contrato, c_lista_contrato.* FROM (tb_pi_contrato LEFT JOIN c_lista_dados_obras ON tb_pi_contrato.cod_empreendimento = c_lista_dados_obras.Código) RIGHT JOIN c_lista_contrato ON tb_pi_contrato.cod_contrato = c_lista_contrato.id ORDER BY c_lista_dados_obras.municipio ASC, c_lista_dados_obras.nome_empreendimento ASC"

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
						<a href="altera_contrato.asp?id=<%=(rs_lista.Fields.Item("cod_contrato").Value)%>">
							<img src="const/imagens/edit.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<a href="delete_contrato.asp?id=<%=(rs_lista.Fields.Item("cod_contrato").Value)%>">
							<img src="const/imagens/delete.gif" width="16" height="15" border="0" />
						</a>
					</td>
										<td>
						<a href="cad_contrato_item.asp?cod_contrato=<%=(rs_lista.Fields.Item("cod_contrato").Value)%>&num_autos=<%=(rs_lista.Fields.Item("num_autos").Value)%>&num_contrato=<%=(rs_lista.Fields.Item("num_contrato").Value)%>">
							<span class="style5">
								Itens
							</span>
						</a>
					</td>
					<td>
						<a href="cad_contrato_cronograma.asp?cod_contrato=<%=(rs_lista.Fields.Item("cod_contrato").Value)%>&num_autos=<%=(rs_lista.Fields.Item("num_autos").Value)%>&num_contrato=<%=(rs_lista.Fields.Item("num_contrato").Value)%>">
							<span class="style5">
								Cronogramas
							</span>
						</a>
					</td>
					<td>
						<a href="cad_contrato_aditamento.asp?cod_contrato=<%=(rs_lista.Fields.Item("cod_contrato").Value)%>&num_autos=<%=(rs_lista.Fields.Item("num_autos").Value)%>&num_contrato=<%=(rs_lista.Fields.Item("num_contrato").Value)%>">
							<span class="style5">
								Aditamentos
							</span>
						</a>
					</td>
					<td>
						<a href="cad_contrato_medicao.asp?cod_contrato=<%=(rs_lista.Fields.Item("cod_contrato").Value)%>&num_autos=<%=(rs_lista.Fields.Item("num_autos").Value)%>&num_contrato=<%=(rs_lista.Fields.Item("num_contrato").Value)%>">
							<span class="style5">
								Medições
							</span>
						</a>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("cod_contrato").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("municipio").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("nome_empreendimento").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("c_lista_contrato.empresa_contratada").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("engenheiro_empresa_contratada").Value)%></span>
					</td>
					<td style="text-align: center;">
						<span class="style5"><%=(rs_lista.Fields.Item("num_autos").Value)%></span>
					</td>
					<td style="text-align: center;">
						<span class="style5"><%=(rs_lista.Fields.Item("num_contrato").Value)%></span>
					</td>
					<td style="text-align: center;">
						<span class="style5"><%=(rs_lista.Fields.Item("c_lista_contrato.dta_assinatura").Value)%></span>
					</td>
					<td style="text-align: center;">
						<span class="style5"><%=(rs_lista.Fields.Item("dta_publicacao_doe").Value)%></span>
					</td>
					<td style="text-align: center;">
						<span class="style5"><%=(rs_lista.Fields.Item("dta_pedido_empenho").Value)%></span>
					</td>
					<td style="text-align: center;">
						<span class="style5"><%=(rs_lista.Fields.Item("dta_os").Value)%></span>
					</td>
					<td style="text-align: center;">
						<span class="style5"><%=(rs_lista.Fields.Item("c_lista_contrato.prz_original_contrato_meses").Value)%></span>
					</td>
					<td style="text-align: center;">
						<span class="style5"><%=(rs_lista.Fields.Item("c_lista_contrato.prz_original_execucao_meses").Value)%></span>
					</td>
					<td style="text-align: center;">
						<span class="style5"><%=(rs_lista.Fields.Item("c_lista_contrato.dta_vigencia").Value)%></span>
					</td>
					<td style="text-align: center;">
						<span class="style5"><%=(rs_lista.Fields.Item("dta_base").Value)%></span>
					</td>
					<td style="text-align: center;">
						<span class="style5"><%=(rs_lista.Fields.Item("c_lista_contrato.dta_inauguracao").Value)%></span>
					</td>
					<td style="text-align: center;">
						<span class="style5"><%=(rs_lista.Fields.Item("dta_termo_recebimento_provisorio").Value)%></span>
					</td>
					<td style="text-align: center;">
						<span class="style5"><%=(rs_lista.Fields.Item("dta_termo_recebimento_definitivo").Value)%></span>
					</td>
					<td style="text-align: center;">
						<span class="style5"><%=(rs_lista.Fields.Item("dta_encerramento_contrato").Value)%></span>
					</td>
					<td style="text-align: center;">
						<span class="style5"><%=(rs_lista.Fields.Item("dta_recisao_contratual").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("contratante").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("situacao").Value)%></span>
					</td>
					<td>
						<form id="form-upload" method="post" enctype="multipart/form-data"
							action="novo_upload.asp?id=<%=(rs_lista.Fields.Item("cod_contrato").Value)%>&folder=CONTRATO&retUrl=<%=(Request.ServerVariables("URL"))%>?<%=(Request.QueryString)%>">
							<input type="file" name="blob">
							<br/>
							<input type="submit" id="btnSubmit" value="Upload">
						</form>

						<%
							cod_convenio = rs_lista.Fields.Item("cod_contrato").Value
							strF = "SELECT * FROM tb_contrato_arquivo WHERE cod_referencia = " & cod_convenio

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
									<a href="download.asp?path=/ARQUIVOS/CONTRATO&filename=<%=(rs_lista.Fields.Item("cod_contrato").Value)%>_<%=(rs_files.Fields.Item("nme_arquivo").Value)%>">
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