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
			rs_update("num_autos") = Trim(Request.Form("num_autos"))
			rs_update("num_contrato") = Trim(Request.Form("num_contrato"))
			rs_update("cod_projetista") = Trim(Request.Form("cod_projetista"))
			rs_update("cod_empresa_contratada") = Trim(Request.Form("cod_empresa_contratada"))
			rs_update("cod_engenheiro_empresa_contratada") = Trim(Request.Form("cod_engenheiro_empresa_contratada"))
			rs_update("dta_assinatura") = Trim(Request.Form("dta_assinatura"))
			rs_update("dta_publicacao_doe") = Trim(Request.Form("dta_publicacao_doe"))
			rs_update("dta_pedido_empenho") = Trim(Request.Form("dta_pedido_empenho"))
			rs_update("dta_os") = Trim(Request.Form("dta_os"))
			rs_update("prz_original_contrato_meses") = Trim(Request.Form("prz_original_contrato_meses"))
			rs_update("prz_original_execucao_meses") = Trim(Request.Form("prz_original_execucao_meses"))
			rs_update("dta_vigencia") = Trim(Request.Form("dta_vigencia"))
			rs_update("dta_base") = Trim(Request.Form("dta_base"))
			rs_update("dta_inauguracao") = Trim(Request.Form("dta_inauguracao"))
			rs_update("dta_termo_recebimento_provisorio") = Trim(Request.Form("dta_termo_recebimento_provisorio"))
			rs_update("dta_termo_recebimento_definitivo") = Trim(Request.Form("dta_termo_recebimento_definitivo"))
			rs_update("dta_encerramento_contrato") = Trim(Request.Form("dta_encerramento_contrato"))
			rs_update("dta_recisao_contratual") = Trim(Request.Form("dta_recisao_contratual"))
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
				<span class="style17">Cadastro de Contratos</span>
			</strong>
		</p>

		<form method="post" name="form1">
			<table align="center">
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Núm. Autos:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="num_autos" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Núm. Contrato:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="num_contrato" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Projetista:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_projetista">
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
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Empresa Contratada:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_empresa_contratada">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM tb_Construtora "

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
						<span class="style22">Eng. Empresa Contratada:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_engenheiro_empresa_contratada">
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
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data de Assinatura:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_assinatura" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data de Publicação no DOE:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_publicacao_doe" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data de Pedido do Empenho:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_pedido_empenho" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data da Ordem de Serviço:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_os" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Prazo Original do Contrato (meses):</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="prz_original_contrato_meses" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Prazo Original de Execução (meses):</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="prz_original_execucao_meses" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Vigente Até:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_vigencia" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data Base:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_base" value="" size="10">
					</td>
				</tr>

				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data de Inauguração:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_inauguracao" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data de Recebimento Termo Provisório:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_termo_recebimento_provisorio" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data de Recebimento Termo Definitivo:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_termo_recebimento_definitivo" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data de Encerramento do Contrato:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_encerramento_contrato" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data de Recisão Contratual:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_recisao_contratual" value="" size="10">
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
					<td><span class="style7">Núm. Autos</span></td>
					<td><span class="style7">Núm. Contrato</span></td>
					<td><span class="style7">Projetista</span></td>
					<td><span class="style7">Empresa Contratada</span></td>
					<td><span class="style7">Eng. Empresa Contratada</span></td>
					<td><span class="style7">Data de Assinatura</span></td>
					<td><span class="style7">Data de Publicação no DOE</span></td>
					<td><span class="style7">Data de Pedido do Empenho</span></td>
					<td><span class="style7">Data da Ordem de Serviço</span></td>
					<td><span class="style7">Prazo Original do Contrato (meses)</span></td>
					<td><span class="style7">Prazo Original de Execução (meses)</span></td>
					<td><span class="style7">Vigente Até</span></td>
					<td><span class="style7">Data Base</span></td>
					<td><span class="style7">Data de Inauguração</span></td>
					<td><span class="style7">Data de Recebimento Termo Provisório</span></td>
					<td><span class="style7">Data de Recebimento Termo Definitivo</span></td>
					<td><span class="style7">Data de Encerramento do Contrato</span></td>
					<td><span class="style7">Data de Recisão Contratual</span></td>
				</tr>
				<%
					strQ = "SELECT * from c_lista_contrato"

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
						<span class="style5"><%=(rs_lista.Fields.Item("num_autos").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("num_contrato").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("projetista").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("empresa_contratada").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("engenheiro_empresa_contratada").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_assinatura").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_publicacao_doe").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_pedido_empenho").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_os").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("prz_original_contrato_meses").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("prz_original_execucao_meses").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_vigencia").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_base").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_inauguracao").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_termo_recebimento_provisorio").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_termo_recebimento_definitivo").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_encerramento_contrato").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_recisao_contratual").Value)%></span>
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