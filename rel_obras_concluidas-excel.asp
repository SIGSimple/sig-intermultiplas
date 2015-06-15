<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="logout.asp" -->
<%
	Response.CharSet = "UTF-8"

	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
		objCon.Open MM_cpf_STRING

	Dim sql

	sql = ""
	sql = sql & "SELECT *, "
	sql = sql & 	"tb_info_emp_concluidos.id,"
	sql = sql & 	"tb_info_emp_concluidos.num_autos,"
	sql = sql & 	"tb_info_emp_concluidos.nme_municipio,"
	sql = sql & 	"tb_info_emp_concluidos.nme_localidade,"

	sql = sql & 	"tb_info_emp_concluidos.dsc_necessita_reparos,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_problemas_bombas,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_falta_limpeza,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_despejo_irregular_residuos,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_falta_funcionario_operacao,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_danos_cercamento,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_danos_tratamento_preliminar,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_danos_talude,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_danos_lagoa,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_problemas_diversos,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_danos_emissarios,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_danos_caixa_passagem_interligacoes,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_danos_drenagem,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_partes_inoperantes,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_situacao_operacao,"

	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_necessita_reparos]),0,1) 					AS flg_necessita_reparos,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_problemas_bombas]),0,1) 					AS flg_problemas_bombas,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_falta_limpeza]),0,1) 						AS flg_falta_limpeza,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_despejo_irregular_residuos]),0,1) 			AS flg_despejo_irregular_residuos,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_falta_funcionario_operacao]),0,1) 			AS flg_falta_funcionario_operacao,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_danos_cercamento]),0,1) 					AS flg_danos_cercamento,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_danos_tratamento_preliminar]),0,1) 			AS flg_danos_tratamento_preliminar,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_danos_talude]),0,1) 						AS flg_danos_talude,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_danos_lagoa]),0,1) 							AS flg_danos_lagoa,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_problemas_diversos]),0,1) 					AS flg_problemas_diversos,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_danos_emissarios]),0,1) 					AS flg_danos_emissarios,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_danos_caixa_passagem_interligacoes]),0,1) 	AS flg_danos_caixa_passagem_interligacoes,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_danos_drenagem]),0,1) 						AS flg_danos_drenagem,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_partes_inoperantes]),0,1) 					AS flg_partes_inoperantes "
	sql = sql & "FROM tb_info_emp_concluidos "

	Set rs_lista_matriz = Server.CreateObject("ADODB.Recordset")
		rs_lista_matriz.CursorLocation = 3
		rs_lista_matriz.CursorType = 3
		rs_lista_matriz.LockType = 1
		rs_lista_matriz.Open sql, objCon, , , &H0001

	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename="& now() &" | Banco de Dados Relatório de Vistorias.xls"
%>
<!DOCTYPE html>
<html>
<head>
	<title>:: DAEE ::</title>
	<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
	<style type="text/css">
		html, body {
			font-family: monospace;
		}

		table {
			text-align: left;
		}

		table thead {
			font-size: 14px;
		}

		table thead th {
			background-color: #2C3E50;
			color: #FFF;
			vertical-align: middle !important;
		}

		td {
			vertical-align: middle !important;
		}

		.text-center {
			text-align: center;
		}
	</style>
</head>
<body>
	<table border="1">
		<thead>
			<th width="200" style="min-width: 200px; width: 200px; max-width: 200px;">Município</th>
			<th width="200" style="min-width: 200px; width: 200px; max-width: 200px;">Localidade</th>
			<th width="200" style="min-width: 200px; width: 200px; max-width: 200px;" class="text-center">Situação</th>

			<th width="300" style="min-width: 300px; width: 300px; max-width: 300px;" class="text-center">Diretoria de Bacia</th>
			<th width="300" style="min-width: 300px; width: 300px; max-width: 300px;" class="text-center">UGRHI</th>
			<th width="300" style="min-width: 300px; width: 300px; max-width: 300px;" class="text-center">Prefeito</th>
			<th width="300" style="min-width: 300px; width: 300px; max-width: 300px;" class="text-center">Técnico Vistoriador</th>
			<th width="300" style="min-width: 300px; width: 300px; max-width: 300px;" class="text-center">Endereço</th>
			<th width="300" style="min-width: 300px; width: 300px; max-width: 300px;" class="text-center">E-Mail</th>
			<th width="300" style="min-width: 300px; width: 300px; max-width: 300px;" class="text-center">População Atual (2010)</th>
			<th width="300" style="min-width: 300px; width: 300px; max-width: 300px;" class="text-center">População Futura (2030)</th>
			<th width="300" style="min-width: 300px; width: 300px; max-width: 300px;" class="text-center">Dados Sobre o Esgotamento Sanitário</th>
			<th width="300" style="min-width: 300px; width: 300px; max-width: 300px;" class="text-center">Coordenadas UTM - Chegada do Esgoto</th>
			<th width="300" style="min-width: 300px; width: 300px; max-width: 300px;" class="text-center">Coordenadas UTM - Lançamento do Esgoto no Corpo Receptor</th>
			<th width="300" style="min-width: 300px; width: 300px; max-width: 300px;" class="text-center">ETE - Aspectos Administrativos e de Logística</th>
			<th width="300" style="min-width: 300px; width: 300px; max-width: 300px;" class="text-center">ETE - Croqui s/ Escala - Indicação dos Dispositivos (Composição do Tratamento)</th>
			<th width="300" style="min-width: 300px; width: 300px; max-width: 300px;" class="text-center">ETE - Dispositivos - Conservação e Manutenção</th>
			<th width="300" style="min-width: 300px; width: 300px; max-width: 300px;" class="text-center">ETE - Entorno - Descrição e Manutenção</th>
			<th width="300" style="min-width: 300px; width: 300px; max-width: 300px;" class="text-center">Comentários em Geral</th>

			<th width="200" style="min-width: 200px; width: 200px; max-width: 200px;" class="text-center">
				<i class="fa fa-wrench" data-toggle="tooltip" data-placement="top" title="Necessita de Reparos"></i> Necessita de Reparos
			</th>
			<th width="200" style="min-width: 200px; width: 200px; max-width: 200px;" class="text-center">
				<i class="fa fa-bomb" data-toggle="tooltip" data-placement="top" title="Problema nas Bombas"></i> Problema nas Bombas
			</th>
			<th width="200" style="min-width: 200px; width: 200px; max-width: 200px;" class="text-center">
				<i class="fa fa-eraser" data-toggle="tooltip" data-placement="top" title="Falta Limpeza"></i> Falta Limpeza
			</th>
			<th width="200" style="min-width: 200px; width: 200px; max-width: 200px;" class="text-center">
				<i class="fa fa-trash-o" data-toggle="tooltip" data-placement="top" title="Despejo Irregular de Resíduos"></i> Despejo Irregular de Resíduos
			</th>
			<th width="200" style="min-width: 200px; width: 200px; max-width: 200px;" class="text-center">
				<i class="fa fa-user" data-toggle="tooltip" data-placement="top" title="Falta Funcionário p/ Operação Diária"></i> Falta Funcionário p/ Operação Diária
			</th>
			<th width="200" style="min-width: 200px; width: 200px; max-width: 200px;" class="text-center">
				<i class="fa fa-table" data-toggle="tooltip" data-placement="top" title="Danos no Cercamento"></i> Danos no Cercamento
			</th>
			<th width="200" style="min-width: 200px; width: 200px; max-width: 200px;" class="text-center">
				<i class="fa fa-filter" data-toggle="tooltip" data-placement="top" title="Danos no Tratamento Preliminar"></i> Danos no Tratamento Preliminar
			</th>
			<th width="200" style="min-width: 200px; width: 200px; max-width: 200px;" class="text-center">
				<i class="fa fa-area-chart" data-toggle="tooltip" data-placement="top" title="Danos no Talude"></i> Danos no Talude
			</th>
			<th width="200" style="min-width: 200px; width: 200px; max-width: 200px;" class="text-center">
				<i class="fa fa-tint" data-toggle="tooltip" data-placement="top" title="Danos nas Lagoas"></i> Danos nas Lagoas
			</th>
			<th width="200" style="min-width: 200px; width: 200px; max-width: 200px;" class="text-center">
				<i class="fa fa-warning" data-toggle="tooltip" data-placement="top" title="Problemas Diversos"></i> Problemas Diversos
			</th>
			<th width="200" style="min-width: 200px; width: 200px; max-width: 200px;" class="text-center">
				<i class="fa fa-download" data-toggle="tooltip" data-placement="top" title="Danos nos Emissários"></i> Danos nos Emissários
			</th>
			<th width="200" style="min-width: 200px; width: 200px; max-width: 200px;" class="text-center">
				<i class="fa fa-cube" data-toggle="tooltip" data-placement="top" title="Danos nas Caixas de Passagem/Interligações"></i> Danos nas Caixas de Passagem/Interligações
			</th>
			<th width="200" style="min-width: 200px; width: 200px; max-width: 200px;" class="text-center">
				<i class="fa fa-code-fork" data-toggle="tooltip" data-placement="top" title="Danos de Drenagem"></i> Danos de Drenagem
			</th>
		</thead>
		<tbody>
			<%
				While (Not rs_lista_matriz.EOF)
			%>
			<tr>
				<td width="200"><%=(rs_lista_matriz.Fields.Item("nme_municipio").Value)%></td>
				<td width="200"><%=(rs_lista_matriz.Fields.Item("nme_localidade").Value)%></td>
				<td width="200" class="text-center"><%=(rs_lista_matriz.Fields.Item("dsc_situacao_operacao").Value)%></td>

				<td width="300"><%=(rs_lista_matriz.Fields.Item("nme_bacia_daee").Value)%></td>
				<td width="300"><%=(rs_lista_matriz.Fields.Item("nme_bacia_secretaria").Value)%></td>
				<td width="300"><%=(rs_lista_matriz.Fields.Item("nme_prefeito").Value)%></td>
				<td width="300"><%=(rs_lista_matriz.Fields.Item("nme_vistoriador").Value)%></td>
				<td width="300"><%=(rs_lista_matriz.Fields.Item("dsc_endereco").Value)%></td>
				<td width="300"><%=(rs_lista_matriz.Fields.Item("end_email").Value)%></td>
				<td width="300"><%=(rs_lista_matriz.Fields.Item("qtd_populacao_2010").Value)%></td>
				<td width="300"><%=(rs_lista_matriz.Fields.Item("qtd_populacao_2030").Value)%></td>
				<td width="300"><%=(rs_lista_matriz.Fields.Item("dsc_dados_basicos_esgotamento_sanitario").Value)%></td>
				<td width="300"><%=(rs_lista_matriz.Fields.Item("num_latitude_chegada_esgoto").Value)%> - <%=(rs_lista_matriz.Fields.Item("num_longitude_chegada_esgoto").Value)%></td>
				<td width="300"><%=(rs_lista_matriz.Fields.Item("num_latitude_lancamento_esgoto").Value)%> - <%=(rs_lista_matriz.Fields.Item("num_longitude_lancamento_esgoto").Value)%></td>
				<td width="300"><%=(rs_lista_matriz.Fields.Item("dsc_aspectos_administrativos_logistica").Value)%></td>
				<td width="300"><%=(rs_lista_matriz.Fields.Item("dsc_ete_dispositivos_composicao_tratamento").Value)%></td>
				<td width="300"><%=(rs_lista_matriz.Fields.Item("dsc_ete_dispositivos_conservacao_manutencao").Value)%></td>
				<td width="300"><%=(rs_lista_matriz.Fields.Item("dsc_ete_entorno_descricao_manuntencao").Value)%></td>
				<td width="300"><%=(rs_lista_matriz.Fields.Item("dsc_comentarios_gerais").Value)%></td>

				<td width="200" class="text-center">
					<%
						If rs_lista_matriz.Fields.Item("flg_necessita_reparos").Value = 1 Then
					%>
					<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
						data-placement="bottom" data-trigger="hover" title="Necessita de Reparos"
						data-content="<%=(rs_lista_matriz.Fields.Item("dsc_necessita_reparos").Value)%>">
						<i class="fa fa-wrench"></i> <%=(rs_lista_matriz.Fields.Item("dsc_necessita_reparos").Value)%>
					</a>
					<%
						End If
					%>
				</td>
				<td width="200" class="text-center">
					<%
						If rs_lista_matriz.Fields.Item("flg_problemas_bombas").Value = 1 Then
					%>
					<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
						data-placement="bottom" data-trigger="hover" title="Problema nas Bombas"
						data-content="<%=(rs_lista_matriz.Fields.Item("dsc_problemas_bombas").Value)%>">
						<i class="fa fa-bomb"></i> <%=(rs_lista_matriz.Fields.Item("dsc_problemas_bombas").Value)%>
					</a>
					<%
						End If
					%>
				</td>
				<td width="200" class="text-center">
					<%
						If rs_lista_matriz.Fields.Item("flg_falta_limpeza").Value = 1 Then
					%>
					<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
						data-placement="bottom" data-trigger="hover" title="Falta Limpeza"
						data-content="<%=(rs_lista_matriz.Fields.Item("dsc_falta_limpeza").Value)%>">
						<i class="fa fa-eraser"></i> <%=(rs_lista_matriz.Fields.Item("dsc_falta_limpeza").Value)%>
					</a>
					<%
						End If
					%>
				</td>
				<td width="200" class="text-center">
					<%
						If rs_lista_matriz.Fields.Item("flg_despejo_irregular_residuos").Value = 1 Then
					%>
					<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
						data-placement="bottom" data-trigger="hover" title="Despejo Irregular de Resíduos"
						data-content="<%=(rs_lista_matriz.Fields.Item("dsc_despejo_irregular_residuos").Value)%>">
						<i class="fa fa-trash-o"></i> <%=(rs_lista_matriz.Fields.Item("dsc_despejo_irregular_residuos").Value)%>
					</a>
					<%
						End If
					%>
				</td>
				<td width="200" class="text-center">
					<%
						If rs_lista_matriz.Fields.Item("flg_falta_funcionario_operacao").Value = 1 Then
					%>
					<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
						data-placement="bottom" data-trigger="hover" title="Falta Funcionário p/ Operação Diária"
						data-content="<%=(rs_lista_matriz.Fields.Item("dsc_falta_funcionario_operacao").Value)%>">
						<i class="fa fa-user"></i> <%=(rs_lista_matriz.Fields.Item("dsc_falta_funcionario_operacao").Value)%>
					</a>
					<%
						End If
					%>
				</td>
				<td width="200" class="text-center">
					<%
						If rs_lista_matriz.Fields.Item("flg_danos_cercamento").Value = 1 Then
					%>
					<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
						data-placement="bottom" data-trigger="hover" title="Danos no Cercamento"
						data-content="<%=(rs_lista_matriz.Fields.Item("dsc_danos_cercamento").Value)%>">
						<i class="fa fa-table"></i> <%=(rs_lista_matriz.Fields.Item("dsc_danos_cercamento").Value)%>
					</a>
					<%
						End If
					%>
				</td>
				<td width="200" class="text-center">
					<%
						If rs_lista_matriz.Fields.Item("flg_danos_tratamento_preliminar").Value = 1 Then
					%>
					<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
						data-placement="bottom" data-trigger="hover" title="Danos no Tratamento Preliminar"
						data-content="<%=(rs_lista_matriz.Fields.Item("dsc_danos_tratamento_preliminar").Value)%>">
						<i class="fa fa-filter"></i> <%=(rs_lista_matriz.Fields.Item("dsc_danos_tratamento_preliminar").Value)%>
					</a>
					<%
						End If
					%>
				</td>
				<td width="200" class="text-center">
					<%
						If rs_lista_matriz.Fields.Item("flg_danos_talude").Value = 1 Then
					%>
					<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
						data-placement="bottom" data-trigger="hover" title="Danos no Talude"
						data-content="<%=(rs_lista_matriz.Fields.Item("dsc_danos_talude").Value)%>">
						<i class="fa fa-area-chart"></i> <%=(rs_lista_matriz.Fields.Item("dsc_danos_talude").Value)%>
					</a>
					<%
						End If
					%>
				</td>
				<td width="200" class="text-center">
					<%
						If rs_lista_matriz.Fields.Item("flg_danos_lagoa").Value = 1 Then
					%>
					<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
						data-placement="bottom" data-trigger="hover" title="Danos nas Lagoas"
						data-content="<%=(rs_lista_matriz.Fields.Item("dsc_danos_lagoa").Value)%>">
						<i class="fa fa-tint"></i> <%=(rs_lista_matriz.Fields.Item("dsc_danos_lagoa").Value)%>
					</a>
					<%
						End If
					%>
				</td>
				<td width="200" class="text-center">
					<%
						If rs_lista_matriz.Fields.Item("flg_problemas_diversos").Value = 1 Then
					%>
					<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
						data-placement="bottom" data-trigger="hover" title="Problemas Diversos"
						data-content="<%=(rs_lista_matriz.Fields.Item("dsc_problemas_diversos").Value)%>">
						<i class="fa fa-warning"></i> <%=(rs_lista_matriz.Fields.Item("dsc_problemas_diversos").Value)%>
					</a>
					<%
						End If
					%>
				</td>
				<td width="200" class="text-center">
					<%
						If rs_lista_matriz.Fields.Item("flg_danos_emissarios").Value = 1 Then
					%>
					<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
						data-placement="bottom" data-trigger="hover" title="Danos nos Emissários"
						data-content="<%=(rs_lista_matriz.Fields.Item("dsc_danos_emissarios").Value)%>">
						<i class="fa fa-download"></i> <%=(rs_lista_matriz.Fields.Item("dsc_danos_emissarios").Value)%>
					</a>
					<%
						End If
					%>
				</td>
				<td width="200" class="text-center">
					<%
						If rs_lista_matriz.Fields.Item("flg_danos_caixa_passagem_interligacoes").Value = 1 Then
					%>
					<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
						data-placement="bottom" data-trigger="hover" title="Danos nas Caixas de Passagem/Interligações"
						data-content="<%=(rs_lista_matriz.Fields.Item("dsc_danos_caixa_passagem_interligacoes").Value)%>">
						<i class="fa fa-cube"></i> <%=(rs_lista_matriz.Fields.Item("dsc_danos_caixa_passagem_interligacoes").Value)%>
					</a>
					<%
						End If
					%>
				</td>
				<td width="200" class="text-center">
					<%
						If rs_lista_matriz.Fields.Item("flg_danos_drenagem").Value = 1 Then
					%>
					<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
						data-placement="bottom" data-trigger="hover" title="Danos de Drenagem"
						data-content="<%=(rs_lista_matriz.Fields.Item("dsc_danos_drenagem").Value)%>">
						<i class="fa fa-code-fork"></i> <%=(rs_lista_matriz.Fields.Item("dsc_danos_drenagem").Value)%>
					</a>
					<%
						End If
					%>
				</td>
			</tr>
			<%
					rs_lista_matriz.MoveNext()
				Wend
			%>
		</tbody>
	</table>
</body>
</html>