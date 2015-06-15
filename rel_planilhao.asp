<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="logout.asp" -->
<%
	Response.CharSet = "UTF-8"

	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	sql = "SELECT * FROM c_lista_rel_planilhao"

	Dim rs_lista

	Set rs_lista = Server.CreateObject("ADODB.Recordset")
		rs_lista.ActiveConnection = MM_cpf_STRING
		rs_lista.Source = sql
		rs_lista.CursorType = 0
		rs_lista.CursorLocation = 2
		rs_lista.LockType = 1
		rs_lista.Open()
%>
<!DOCTYPE html>
<html>
<head>
	<title>:: DAEE ::</title>
	<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
	<link rel="stylesheet" type="text/css" href="css/bootstrap-flaty.min.css">
	<link rel="stylesheet" type="text/css" href="css/daee.css">
	<link rel="stylesheet" href="//maxcdn.bootstrapcdn.com/font-awesome/4.3.0/css/font-awesome.min.css">
	<script type="text/javascript" src="//code.jquery.com/jquery-1.11.2.min.js"></script>
	<script type="text/javascript" src="js/jquery.number.min.js"></script>
	<script type="text/javascript" src="js/jquery.table2excel.js"></script>
	<script type="text/javascript" src="js/moment.min.js"></script>
	<script type="text/javascript" src="js/fullscreen.js"></script>
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
	<script type="text/javascript" src="js/fancybox/jquery.fancybox.pack.js?v=2.1.5"></script>
	<script type="text/javascript" src="js/jquery.floatThead.min.js"></script>
	<script type="text/javascript">
		$(function(){
			$("li a.print").on("click", function(){
				window.print();
			});

			$("li a.excel").on("click", function(){
				$("table#data").table2excel();
			});

			$("table#data").floatThead({
				scrollingTop: 60
			});

			var vlr_lines = $(".vlr");
			$.each(vlr_lines, function(i, item){
				$(item).val("R$ " + $.number($(item).val(), 2, ",", "."));
				$(item).text("R$ " + $.number($(item).text(), 2, ",", "."));
			});
		});
	</script>
</head>
<body>
	<nav class="navbar navbar-default navbar-fixed-top">
		<div class="container-fluid">
			<div class="navbar-header">
				<button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#bs-example-navbar-collapse-1">
					<span class="sr-only">Toggle navigation</span>
					<span class="icon-bar"></span>
					<span class="icon-bar"></span>
					<span class="icon-bar"></span>
				</button>
				<a class="navbar-brand" href="#">SIG - Planilhão de Informações</a>
			</div>

			<div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
				<ul class="nav navbar-nav navbar-right">n
					<li><a href="javascript:window.close();"><i class="fa fa-times-circle"></i> Fechar Janela</a></li>
					<li><a href="#" class="print"><i class="fa fa-print"></i> Imprimir</a></li>
					<li><a href="#" class="excel"><i class="fa fa-file-excel-o"></i> Exportar p/ Excel</a></li>
					<li><a href="#" class="expand"><i class="fa fa-expand"></i>&nbsp;&nbsp;Tela Cheia</a></li>
					<li><a href="<%= MM_Logout %>" class="sign-out"><i class="fa fa-sign-out"></i> Sair do Sistema</a></li>
				</ul>
			</div>
		</div>
	</nav>

	<div>
		<div class="panel panel-default">
			<div class="panel-body">
				<div class="row">
					<div class="col-xs-2 text-center">
						<img src="LogoProjetoAguaLimpa.jpg" class="logo-intermultiplas report"/>
					</div>
					<div class="col-xs-8 text-center">
						<h3><strong>Planilhão de Informações</strong><br/><small>Listagem do Banco de Dados</small></h3>
					</div>
					<div class="col-xs-2"></div>
				</div>

				<hr/>

				<div class="row">
					<div class="col-xs-12">
						<table id="data" class="table table-bordered table-condensed table-hover table-striped">
							<thead>
								<tr class="active">
									<th class="text-center text-middle" rowspan="3">Município</th>
									<th class="text-center text-middle" rowspan="3">Localidade</th>
									<th class="text-center text-middle" rowspan="3">Situação Atual</th>
									<th class="text-center text-middle" colspan="13">Resumo de Situação por Município e Localidade</th>
									<th class="text-center text-middle" colspan="14">Convênios</th>
									<th class="text-center text-middle" colspan="8">Licitações</th>
									<th class="text-center text-middle" colspan="25">Contratos</th>
									<th class="text-center text-middle" colspan="6">Gestores</th>
									<th class="text-center text-middle" colspan="17">Informações Complementares</th>
								</tr>
								<tr class="active">
									<!-- RESUMO DE SITUACAO POR MUNICIPIO E LOCALIDADE -->
									<th class="text-middle text-center" rowspan="2" style="min-width: 180px;">Bacia</th>
									<th class="text-middle text-center" rowspan="2" style="min-width: 100px;">Status</th>
									<th class="text-middle text-center" rowspan="2" style="min-width: 100px;">IBGE 2010</th>
									<th class="text-middle text-center" rowspan="2" style="min-width: 100px;">POP 2030</th>
									<th class="text-middle text-center" rowspan="2" style="min-width: 150px;">Situação Atual</th>
									<th class="text-middle text-center" rowspan="2" style="min-width: 150px;">Investimento Governo SP</th>
									<th class="text-middle text-center" colspan="3">Em Atendimento</th>
									<th class="text-middle text-center" rowspan="2" style="min-width: 120px;">Concluída<br/>Inaugurada</th>
									<th class="text-middle text-center" rowspan="2" style="min-width: 120px;">Previsão para<br/>Inauguração</th>
									<th class="text-middle text-center" rowspan="2" style="min-width: 170px;">Carga Orgânica Retirada<br/>(toneladas/mês)</th>
									<th class="text-middle text-center" rowspan="2" style="min-width: 250px;">Notas / Observações</th>

									<!-- CONVENIO -->
									<th class="text-center text-middle" rowspan="2">Nº Autos Convênio</th>
									<th class="text-center text-middle" rowspan="2">Nº Convênio</th>
									<th class="text-center text-middle" rowspan="2">Dt. Assinatura</th>
									<th class="text-center text-middle" rowspan="2">Dt. Publ. D.O.E.</th>
									<th class="text-center text-middle" rowspan="2">Vigência Até</th>
									<th class="text-center text-middle" rowspan="2">Enquadramento</th>
									<th class="text-center text-middle" rowspan="2">Programa</th>
									<th class="text-center text-middle" rowspan="2">Coord. DAEE</th>
									<th class="text-center text-middle" rowspan="2">Prazo Original</th>
									<th class="text-center text-middle" rowspan="2">Aditivos (Prazo)</th>
									<th class="text-center text-middle" rowspan="2">Prazo Total</th>
									<th class="text-center text-middle" rowspan="2" style="min-width: 150px;">Valor Original</th>
									<th class="text-center text-middle" rowspan="2" style="min-width: 150px;">Aditivos (Valor)</th>
									<th class="text-center text-middle" rowspan="2" style="min-width: 150px;">Valor Total</th>

									<!-- LICITACAO -->
									<th class="text-center text-middle" rowspan="2">Núm. Autos Licitação</td>
									<th class="text-center text-middle" rowspan="2">Tipo de Contratação</td>
									<th class="text-center text-middle" rowspan="2">Financiador</td>
									<th class="text-center text-middle" rowspan="2">Modalidade de Contratação</td>
									<th class="text-center text-middle" rowspan="2">Núm. Edital</td>
									<th class="text-center text-middle" rowspan="2">Data de Publicação D.O.E</td>
									<th class="text-center text-middle" rowspan="2">Data da Licitação</td>
									<th class="text-center text-middle" rowspan="2">Status/Situação</td>

									<!-- CONTRATO -->
									<th class="text-center text-middle" rowspan="2">Empresa Contratada</th>
									<th class="text-center text-middle" rowspan="2">Eng. Empresa Contr.</th>
									<th class="text-center text-middle" rowspan="2">Nº Autos Contrato</th>
									<th class="text-center text-middle" rowspan="2">Nº Contrato</th>
									<th class="text-center text-middle" rowspan="2">Dt. Assinatura</th>
									<th class="text-center text-middle" rowspan="2">Dt. Publ. D.O.E.</th>
									<th class="text-center text-middle" rowspan="2">Dt. Pedido Empenho</th>
									<th class="text-center text-middle" rowspan="2">Dt. Base</th>
									<th class="text-center text-middle" rowspan="2">Dt. Inauguração</th>
									<th class="text-center text-middle" rowspan="2">Dt. Termo Rec. Provisório</th>
									<th class="text-center text-middle" rowspan="2">Dt. Termo Rec. Definitivo</th>
									<th class="text-center text-middle" rowspan="2">Dt. Enc. Contrato</th>
									<th class="text-center text-middle" rowspan="2">Dt. Rec. Contratual</th>
									<th class="text-center text-middle" rowspan="2">Dt. O.S.</th>
									<th class="text-center text-middle" rowspan="2">Vigência Até (Digitado)</th>
									<th class="text-center text-middle" rowspan="2">Vigência Até (Calculado)</th>
									<th class="text-center text-middle" rowspan="2">Prazo Original Execução Serviço</th>
									<th class="text-center text-middle" rowspan="2">Aditivos (Prazo)</th>
									<th class="text-center text-middle" rowspan="2">Prazo Total Serviço</th>
									<th class="text-center text-middle" rowspan="2">Prazo Original Contrato</th>
									<th class="text-center text-middle" rowspan="2">Aditivos (Prazo)</th>
									<th class="text-center text-middle" rowspan="2">Prazo Total Contrato</th>								
									<th class="text-center text-middle" rowspan="2" style="min-width: 150px;">Valor Original</th>
									<th class="text-center text-middle" rowspan="2" style="min-width: 150px;">Aditivos (Valor)</th>
									<th class="text-center text-middle" rowspan="2" style="min-width: 150px;">Valor Total</th>

									<!-- GESTORES -->
									<th class="text-center text-middle" rowspan="2">Diretor de Bacia</th>
									<th class="text-center text-middle" rowspan="2">Eng. DAEE</th>
									<th class="text-center text-middle" rowspan="2">Eng. Obras Consórcio</th>
									<th class="text-center text-middle" rowspan="2">Eng. Plan. Consórcio</th>
									<th class="text-center text-middle" rowspan="2">Fiscal Consórcio</th>
									<th class="text-center text-middle" rowspan="2">Eng. Resp. Medição</th>

									<!-- INFORMACOES COMPLEMENTARES -->
									<th class="text-center text-middle" rowspan="2">Nº Autos Empreendimento</th>
									<th class="text-center text-middle" rowspan="2">Objeto da Obra</th>
									<th class="text-center text-middle" rowspan="2">Bacia Hidrográfica</th>
									<th class="text-center text-middle" rowspan="2">Manancial de Lançamento</th>
									<th class="text-center text-middle" rowspan="2">Latitude/Longitude</th>
									<th class="text-center text-middle" rowspan="2">Coletor Tronco (metros)</th>
									<th class="text-center text-middle" rowspan="2">Interceptor (metros)</th>
									<th class="text-center text-middle" rowspan="2">Emissário fluente Bruto (metros)</th>
									<th class="text-center text-middle" rowspan="2">EEE (qtd)</th>
									<th class="text-center text-middle" rowspan="2">Linha de Recalque (metros)</th>
									<th class="text-center text-middle" rowspan="2">Tipo ETE</th>
									<th class="text-center text-middle" rowspan="2">Estação de Tratamento (desc.)</th>
									<th class="text-center text-middle" rowspan="2">Emissário Efluente Tratado (metros)</th>
									<th class="text-center text-middle" rowspan="2">Estudo Elab. DAEE</th>
									<th class="text-center text-middle" rowspan="2">Observações</th>
									<th class="text-center text-middle" rowspan="2">Benefício Geral da Obra</th>
									<th class="text-center text-middle" rowspan="2">Parceria/Realização</th>
								</tr>
								<tr>
									<th class="text-middle text-center" style="min-width: 120px;">Início das Obras</th>
									<th class="text-middle text-center" style="min-width: 120px;">Executado</th>
									<th class="text-middle text-center" style="min-width: 120px;">Previsão de Término</th>
								</tr>
							</thead>
							<tbody>
								<%
									While (NOT rs_lista.EOF)
										prz_total_execucao 		= rs_lista.Fields.Item("prz_original_execucao_meses_contrato").Value + rs_lista.Fields.Item("prz_aditivo_contrato").Value
										prz_total_contrato 		= rs_lista.Fields.Item("prz_original_contrato_meses_contrato").Value + rs_lista.Fields.Item("prz_aditivo_contrato").Value

										If rs_lista.Fields.Item("vlr_total_reajuste").Value <> "" Then
											vlr_total_contrato 	= rs_lista.Fields.Item("vlr_total_reajuste").Value
										Else
											vlr_total_contrato 	= rs_lista.Fields.Item("vlr_original_contrato").Value + rs_lista.Fields.Item("vlr_aditivo_contrato").Value
										End If

										dta_os_contrato			= rs_lista.Fields.Item("dta_os_contrato").Value
								%>
								<tr>
									<td><%=(rs_lista.Fields.Item("nme_municipio").Value)%></td>
									<td><%=(rs_lista.Fields.Item("nome_empreendimento").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dsc_situacao_externa").Value)%></td>

									<!-- RESUMO DE SITUACAO POR MUNICIPIO E LOCALIDADE -->
									<td class="text-middle text-center" style="max-width: 180px;"><%=( Mid(rs_lista.Fields.Item("bacia_daee").Value, 1, 3) )%></td>					
									<td class="text-middle text-center" style="max-width: 100px;"><%=(rs_lista.Fields.Item("cod_status_situacao_interna").Value)%></td>
									<td class="text-middle text-center num" style="max-width: 100px;"><%=(rs_lista.Fields.Item("qtd_populacao_urbana_2010").Value)%></td>
									<td class="text-middle text-center num" style="max-width: 100px;">
										<%
											If Not IsNull(rs_lista.Fields.Item("qtd_populacao_urbana_2010").Value) Then
												data = rs_lista.Fields.Item("qtd_populacao_urbana_2010").Value
												data = data * 1.25
												a = Round(data/100, 0)
												b = a * 100
												Response.Write b
											End If
										%>
									</td>
									<td class="text-middle text-center" style="max-width: 150px;"><%=(rs_lista.Fields.Item("dsc_situacao_externa").Value)%></td>
									<td class="text-middle text-center vlr" style="max-width: 150px;">
										<%
											If rs_lista.Fields.Item("cod_situacao_externa").Value = 41 Or rs_lista.Fields.Item("cod_situacao_externa").Value = 44 Then
												Response.Write rs_lista.Fields.Item("vlr_total_aditamento_convenio").Value
											Else
												If rs_lista.Fields.Item("cod_tipo_contratacao").Value <> "" Then
													If CInt(rs_lista.Fields.Item("cod_tipo_contratacao").Value) = 1 Or CInt(rs_lista.Fields.Item("cod_tipo_contratacao").Value) = 2 Then
														Response.Write rs_lista.Fields.Item("vlr_total_aditamento_convenio").Value
													Else
														If rs_lista.Fields.Item("vlr_total_reajuste").Value <> "" Then
															Response.Write rs_lista.Fields.Item("vlr_total_reajuste").Value
														Else
															If rs_lista.Fields.Item("vlr_original_contrato").Value > 0 Then
																Response.Write rs_lista.Fields.Item("vlr_original_contrato").Value + rs_lista.Fields.Item("vlr_aditivo_contrato").Value
															Else
																Response.Write rs_lista.Fields.Item("vlr_total_aditamento_convenio").Value
															End If
														End If
													End If
												Else
													If rs_lista.Fields.Item("vlr_total_reajuste").Value <> "" Then
														Response.Write rs_lista.Fields.Item("vlr_total_reajuste").Value
													Else
														If rs_lista.Fields.Item("vlr_original_contrato").Value > 0 Then
															Response.Write rs_lista.Fields.Item("vlr_original_contrato").Value + rs_lista.Fields.Item("vlr_aditivo_contrato").Value
														Else
															Response.Write rs_lista.Fields.Item("vlr_total_aditamento_convenio").Value
														End If
													End If
												End If
											End If
										%>
									</td>
									<td class="text-middle text-center" style="max-width: 120px;">
										<%
											If Not IsNull(rs_lista.Fields.Item("mes_inicio_obras").Value) And Not IsEmpty(rs_lista.Fields.Item("mes_inicio_obras").Value) And rs_lista.Fields.Item("mes_inicio_obras").Value <> "" Then
												Response.Write UCase(MonthName(rs_lista.Fields.Item("mes_inicio_obras").Value,True)) & "/" & rs_lista.Fields.Item("ano_inicio_obras").Value
											End If
										%>
									</td>
									<td class="text-middle text-center prc" style="max-width: 120px;">
										<%
											If Not IsNull(rs_lista.Fields.Item("num_percentual_executado").Value) And Not IsEmpty(rs_lista.Fields.Item("num_percentual_executado").Value) Then
												num_percentual_executado = Replace(rs_lista.Fields.Item("num_percentual_executado").Value, ",", ".")

										%>
										<%=(num_percentual_executado)%>
										<%
											End If
										%>
									</td>
									<td class="text-middle text-center" style="max-width: 120px;">
										<%
											If Not IsNull(rs_lista.Fields.Item("mes_previsao_termino").Value) And Not IsEmpty(rs_lista.Fields.Item("mes_previsao_termino").Value) And rs_lista.Fields.Item("mes_previsao_termino").Value <> "" Then
												Response.Write UCase(MonthName(rs_lista.Fields.Item("mes_previsao_termino").Value,True)) & "/" & rs_lista.Fields.Item("ano_previsao_termino").Value
											End If
										%>
									</td>
									<td class="text-middle text-center" style="max-width: 120px;"><%=(rs_lista.Fields.Item("dta_inauguracao_contrato").Value)%></td>
									<td class="text-middle text-center" style="max-width: 120px;"><%=(rs_lista.Fields.Item("dta_previsao_inauguracao").Value)%></td>
									<td class="text-middle text-center" style="max-width: 170px;">
										<%
											If Not IsNull(rs_lista.Fields.Item("qtd_populacao_urbana_2010").Value) Then
												If Not IsNull(b) And (rs_lista.Fields.Item("cod_status_situacao_interna").Value) = 1 Then
													' Base de cálculo = qtd_populacao_urbana_2030 * 0,06 * 30 / 1000
													Response.Write b * 0.0018
												End If
											End If
										%>
									</td>
									<td class="text-middle" style="max-width: 180px;"><%=(rs_lista.Fields.Item("dsc_observacoes_relatorio_mensal").Value)%></td>

									<!-- CONVENIO -->
									<td><%=(rs_lista.Fields.Item("num_autos_convenio").Value)%></td>
									<td><%=(rs_lista.Fields.Item("num_convenio").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_assinatura_convenio").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_publicacao_doe_convenio").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_vigencia_total_convenio").Value)%></td>
									<td><%=(rs_lista.Fields.Item("enquadramento").Value)%></td>
									<td><%=(rs_lista.Fields.Item("programa_convenio").Value)%></td>
									<td><%=(rs_lista.Fields.Item("coordenador_daee").Value)%></td>
									<td><%=(rs_lista.Fields.Item("prz_meses_convenio").Value)%></td>
									<td><%=(rs_lista.Fields.Item("prz_aditivo_convenio").Value)%></td>
									<td><%=(rs_lista.Fields.Item("prz_total_aditamento_convenio").Value)%></td>
									<td class="vlr"><%=(rs_lista.Fields.Item("vlr_original_convenio").Value)%></td>
									<td class="vlr"><%=(rs_lista.Fields.Item("vlr_aditivo_convenio").Value)%></td>
									<td class="vlr"><%=(rs_lista.Fields.Item("vlr_total_aditamento_convenio").Value)%></td>

									<!-- LICITACAO -->
									<td><%=(rs_lista.Fields.Item("num_autos_licitacao"))%></td>
									<td><%=(rs_lista.Fields.Item("dsc_tipo_contratacao").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dsc_financiador").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dsc_modalidade").Value)%></td>
									<td><%=(rs_lista.Fields.Item("num_edital").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_publicacao_doe_licitacao").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_licitacao").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dsc_situacao_licitacao").Value)%></td>

									<!-- CONTRATO -->
									<td><%=(rs_lista.Fields.Item("empresa_contratada").Value)%></td>
									<td><%=(rs_lista.Fields.Item("engenheiro_empresa_contratada").Value)%></td>
									<td><%=(rs_lista.Fields.Item("num_autos_contrato").Value)%></td>
									<td><%=(rs_lista.Fields.Item("num_contrato").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_assinatura_contrato").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_publicacao_doe_contrato").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_pedido_empenho_contrato").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_base_contrato").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_inauguracao_contrato").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_termo_recebimento_provisorio_contrato").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_termo_recebimento_definitivo_contrato").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_encerramento_contrato").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_recisao_contratual").Value)%></td>
									<td><%=(dta_os_contrato)%></td>
									<td></td>
									<td>
										<%
											If IsNull(rs_lista.Fields.Item("dta_vigencia_contrato").Value) Or IsEmpty(rs_lista.Fields.Item("dta_vigencia_contrato").Value) Then
												If dta_os_contrato <> "" Then
													If prz_total_contrato > 0 Then
														dta_vigencia_contrato = DateAdd("m", prz_total_contrato, dta_os_contrato)
														Response.Write dta_vigencia_contrato
													End If
												End If
											Else
												Response.Write rs_lista.Fields.Item("dta_vigencia_contrato").Value
											End If
										%>
									</td>
									
									<td><%=(rs_lista.Fields.Item("prz_original_execucao_meses_contrato").Value)%></td>
									<td><%=(rs_lista.Fields.Item("prz_aditivo_contrato").Value)%></td>
									<td><%=(prz_total_execucao)%></td>
									
									<td><%=(rs_lista.Fields.Item("prz_original_contrato_meses_contrato").Value)%></td>
									<td><%=(rs_lista.Fields.Item("prz_aditivo_contrato").Value)%></td>
									<td><%=(prz_total_contrato)%></td>

									<td class="text-center vlr"><%=(rs_lista.Fields.Item("vlr_original_contrato").Value)%></td>
									<td class="text-center vlr"><%=(rs_lista.Fields.Item("vlr_aditivo_contrato").Value)%></td>
									<td class="text-center vlr"><%=(vlr_total_contrato)%></td>

									<!-- GESTORES -->
									<td><%=(rs_lista.Fields.Item("nme_diretor_bacia_daee").Value)%></td>
									<td><%=(rs_lista.Fields.Item("eng_daee").Value)%></td>
									<td><%=(rs_lista.Fields.Item("eng_obras_consorcio").Value)%></td>
									<td><%=(rs_lista.Fields.Item("eng_plan_consorcio").Value)%></td>
									<td><%=(rs_lista.Fields.Item("fiscal_consorcio").Value)%></td>
									<td><%=(rs_lista.Fields.Item("eng_medicao").Value)%></td>

									<!-- INFORMACOES COMPLEMENTARES -->
									<td><%=(rs_lista.Fields.Item("PI").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dsc_objeto_obra").Value)%></td>
									<td><%=(rs_lista.Fields.Item("nme_bacia_hidrografica").Value)%></td>
									<td><%=(rs_lista.Fields.Item("nme_manancial").Value)%></td>
									<td><%=(rs_lista.Fields.Item("latitude_longitude").Value)%></td>
									<td><%=(rs_lista.Fields.Item("qtd_metragem_coletor_tronco").Value)%></td>
									<td><%=(rs_lista.Fields.Item("qtd_metragem_interceptor").Value)%></td>
									<td><%=(rs_lista.Fields.Item("qtd_metragem_emissario_fluente_bruto").Value)%></td>
									<td><%=(rs_lista.Fields.Item("qtd_eee").Value)%></td>
									<td><%=(rs_lista.Fields.Item("qtd_metragem_linha_recalque").Value)%></td>
									<td><%=(rs_lista.Fields.Item("nme_tipo_ete").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dsc_estacao_tratamento").Value)%></td>
									<td><%=(rs_lista.Fields.Item("qtd_metragem_emissario_efluente_tratado").Value)%></td>
									<td>
										<%=(rs_lista.Fields.Item("flg_estudo_elaborado_daee").Value)%>
									</td>
									<td><%=(rs_lista.Fields.Item("dsc_observacoes_obra").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dsc_resultado_obtido").Value)%></td>
									<td><%=(rs_lista.Fields.Item("nme_parceria_realizacao").Value)%></td>
								</tr>
								<%
										rs_lista.MoveNext()
									Wend
								%>
							</tbody>
						</table>
					</div>
				</div>
			</div>
		</div>
	</div>

</body>
</html>