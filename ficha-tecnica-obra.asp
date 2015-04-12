<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="logout.asp" -->
<!--#include file="daee_restrict_access.asp" -->
<!--#include file="functions.asp" -->
<%
	Response.CharSet = "UTF-8"

	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
		objCon.Open MM_cpf_STRING

	Dim cod_municipio
	Dim cod_empreendimento

	Set cod_municipio = Request.QueryString("cod_municipio")
	Set cod_empreendimento = Request.QueryString("cod_empreendimento")

	strQ = "SELECT * FROM c_lista_pi WHERE PI = '"& cod_empreendimento &"'"

	Set rs_dados_obra = Server.CreateObject("ADODB.Recordset")
		rs_dados_obra.CursorLocation = 3
		rs_dados_obra.CursorType = 3
		rs_dados_obra.LockType = 1
		rs_dados_obra.Open strQ, objCon, , , &H0001
	
	Dim qtd_populacao_urbana_2030
	Dim qtd_carga_organica_removida

	If Not IsNull(rs_dados_obra.Fields.Item("qtd_populacao_urbana_2010").Value) Then
		data = rs_dados_obra.Fields.Item("qtd_populacao_urbana_2010").Value
		data = data * 1.25
		qtd_populacao_urbana_2010 = Round(data/100, 0)
		qtd_populacao_urbana_2030 = qtd_populacao_urbana_2010 * 100

		If Not IsNull(qtd_populacao_urbana_2030) Then
			qtd_carga_organica_removida = qtd_populacao_urbana_2030 * 0.0018
		End If
	End If

	strQ = "SELECT c_lista_contrato.* FROM c_lista_pi INNER JOIN (tb_pi_contrato INNER JOIN c_lista_contrato ON tb_pi_contrato.cod_contrato = c_lista_contrato.id) ON c_lista_pi.Código = tb_pi_contrato.cod_empreendimento WHERE c_lista_pi.PI = '"& cod_empreendimento &"'"

	Set rs_dados_contrato = Server.CreateObject("ADODB.Recordset")
		rs_dados_contrato.CursorLocation = 3
		rs_dados_contrato.CursorType = 3
		rs_dados_contrato.LockType = 1
		rs_dados_contrato.Open strQ, objCon, , , &H0001

	Dim dta_assinatura
	Dim dta_vigencia
	Dim prz_original_execucao_meses
	Dim cod_contrato

	If Not rs_dados_contrato.EOF Then 
		dta_assinatura 				= rs_dados_contrato.Fields.Item("dta_assinatura").Value
		dta_vigencia 				= rs_dados_contrato.Fields.Item("dta_vigencia").Value
		prz_original_execucao_meses = rs_dados_contrato.Fields.Item("prz_original_execucao_meses").Value
		cod_contrato 				= rs_dados_contrato.Fields.Item("id").Value
	End If
%>
<!DOCTYPE html>
<html>
<head>
	<title>:: DAEE ::</title>
	<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
	<link rel="stylesheet" type="text/css" href="css/bootstrap-flaty.min.css">
	<link rel="stylesheet" type="text/css" href="css/daee.css">
	<link rel="stylesheet" href="//maxcdn.bootstrapcdn.com/font-awesome/4.3.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="js/fancybox/jquery.fancybox.css?v=2.1.5" type="text/css" media="screen" />
	<script type="text/javascript" src="//code.jquery.com/jquery-1.11.2.min.js"></script>
	<script type="text/javascript" src="js/jquery.number.min.js"></script>
	<script type="text/javascript" src="js/underscore-min.js"></script>
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
	<script type="text/javascript" src="http://code.highcharts.com/highcharts.js"></script>
	<script type="text/javascript" src="http://code.highcharts.com/modules/exporting.js"></script>
	<script type="text/javascript" src="js/fancybox/jquery.fancybox.pack.js?v=2.1.5"></script>
	<script type="text/javascript" src="js/moment.min.js"></script>
	<script type="text/javascript" src="js/fullscreen.js"></script>
	<script type="text/javascript">
		function capitalizeFirstLetter(string) {
			return string.charAt(0).toUpperCase() + string.slice(1);
		}

		$(function () {
			$(".fancybox").fancybox();

			var colors = ['#f45b5b', '#7cb5ec', '#90ed7d', '#f7a35c', '#8085e9', '#f15c80', '#e4d354', '#2b908f', '#434348', '#91e8e1'];

			Highcharts.setOptions({
				lang: {
					decimalPoint: ',',
					thousandsSep: '.'
				}
			});

			var vlr_lines = $(".vlr");
			$.each(vlr_lines, function(i, item){
				$(item).val("R$ " + $.number($(item).val(), 2, ",", "."));
				$(item).text("R$ " + $.number($(item).text(), 2, ",", "."));
			});

			var prc_lines = $(".prc");
			$.each(prc_lines, function(i, item){
				$(item).val($.number($(item).val(), 0, ",", ".") + "%");
				$(item).text($.number($(item).text(), 0, ",", ".") + "%");
			});

			var num_lines = $(".num");
			$.each(num_lines, function(i, item){
				$(item).val($.number($(item).val(), 0, ",", "."));
				$(item).text($.number($(item).text(), 0, ",", "."));
			});

			$("li a.print").on("click", function(){
				window.print();
			});

			/****************************************************************************************
			 * INICIO								DATASETS										*
			 ****************************************************************************************/

				var months = [
					<%	
						If Not IsNull(cod_contrato) And Not IsEmpty(cod_contrato) Then
							strQ = "SELECT * FROM c_planejamento_aditado_contrato WHERE cod_contrato = " & cod_contrato & " ORDER BY dta_planejamento ASC"

							Set rs_plan = Server.CreateObject("ADODB.Recordset")
								rs_plan.CursorLocation = 3
								rs_plan.CursorType = 3
								rs_plan.LockType = 1
								rs_plan.Open strQ, objCon, , , &H0001

							If Not rs_plan.EOF Then
								While Not rs_plan.EOF
									Response.Write "'"& CaptalizeText(MonthName(rs_plan.Fields.Item("mes").Value,True)) & "/" & rs_plan.Fields.Item("ano").Value &"',"
									rs_plan.MoveNext
								Wend
							End If
						End If
					%>
				];

				var items = [
					<%
						If Not IsNull(cod_contrato) And Not IsEmpty(cod_contrato) Then
							strQ = "SELECT * FROM c_soma_itens_ultima_medicao WHERE cod_contrato = " & cod_contrato

							Set rs_plan = Server.CreateObject("ADODB.Recordset")
								rs_plan.CursorLocation = 3
								rs_plan.CursorType = 3
								rs_plan.LockType = 1
								rs_plan.Open strQ, objCon, , , &H0001

							If Not rs_plan.EOF Then
								While Not rs_plan.EOF
									Response.Write "'"&  rs_plan.Fields.Item("dsc_item").Value &"',"
									rs_plan.MoveNext
								Wend
							End If
						End If
					%>
				]

				var arr_itens_ultima_medicao = [
					<%
						If Not IsNull(cod_contrato) And Not IsEmpty(cod_contrato) Then
							strQ = "SELECT * FROM c_soma_itens_ultima_medicao WHERE cod_contrato = " & cod_contrato

							Set rs_plan = Server.CreateObject("ADODB.Recordset")
								rs_plan.CursorLocation = 3
								rs_plan.CursorType = 3
								rs_plan.LockType = 1
								rs_plan.Open strQ, objCon, , , &H0001

							If Not rs_plan.EOF Then
								While Not rs_plan.EOF
					%>
						{
							cod_item:<%=(Replace(rs_plan.Fields.Item("cod_item").Value, ",","."))%>,
							dta_item:"<%=(Replace(rs_plan.Fields.Item("dta_medicao").Value, ",","."))%>",
							vlr_item:<%=(Replace(rs_plan.Fields.Item("vlr_medido").Value, ",","."))%>,
						},
					<%
									rs_plan.MoveNext
								Wend
							End If
						End If
					%>
				]

				var itemsData = [];

				var arr_itens_planejamento_original = [
					<%
						If Not IsNull(cod_contrato) And Not IsEmpty(cod_contrato) Then
							strQ = "SELECT * FROM c_soma_itens_planejamento_original WHERE cod_contrato = " & cod_contrato & " ORDER BY dta_planejamento ASC, cod_item_planejamento ASC"

							Set rs_plan = Server.CreateObject("ADODB.Recordset")
								rs_plan.CursorLocation = 3
								rs_plan.CursorType = 3
								rs_plan.LockType = 1
								rs_plan.Open strQ, objCon, , , &H0001

							vlr_planejamento = 0

							If Not rs_plan.EOF Then
								While Not rs_plan.EOF
					%>
						{
							cod_item:<%=(Replace(rs_plan.Fields.Item("cod_item_planejamento").Value, ",","."))%>,
							dta_item:"<%=(Replace(rs_plan.Fields.Item("dta_planejamento").Value, ",","."))%>",
							vlr_item:<%=(Replace(rs_plan.Fields.Item("vlr_planejamento").Value, ",","."))%>,
						},
					<%
									rs_plan.MoveNext
								Wend
							End If
						End If
					%>
				];

				arr_itens_planejamento_original = _.sortBy(arr_itens_planejamento_original, 'dta_item');

				var arr_aux = {};
				var dta_corte = moment(arr_itens_ultima_medicao[0].dta_item, "DD/MM/YYYY").format("YYYY/MM/DD");

				$.each(arr_itens_planejamento_original , function(i,itemPlanejamento){
					dta_item = moment(itemPlanejamento.dta_item, "DD/MM/YYYY").format("YYYY/MM/DD");
					if(dta_item <= dta_corte) {
						if(arr_aux[itemPlanejamento.cod_item] == undefined){
							arr_aux[itemPlanejamento.cod_item] = 0
						}
						
						arr_aux[itemPlanejamento.cod_item] += itemPlanejamento.vlr_item;
					}
				});

				$.each(arr_itens_ultima_medicao, function(i,item){
					var vlr_planejado = arr_aux[item.cod_item];
					itemsData.push(parseFloat(((item.vlr_item / vlr_planejado)*100).toFixed(2)));
				});

				var arr_vlr_planejado = [
					<%
						If Not IsNull(cod_contrato) And Not IsEmpty(cod_contrato) Then
							strQ = "SELECT * FROM c_planejamento_aditado_contrato WHERE cod_contrato = " & cod_contrato & " ORDER BY dta_planejamento ASC"

							Set rs_plan = Server.CreateObject("ADODB.Recordset")
								rs_plan.CursorLocation = 3
								rs_plan.CursorType = 3
								rs_plan.LockType = 1
								rs_plan.Open strQ, objCon, , , &H0001

							vlr_planejamento = 0

							If Not rs_plan.EOF Then
								While Not rs_plan.EOF
									Response.Write Replace(rs_plan.Fields.Item("vlr_planejado").Value, ",",".") & ","
									rs_plan.MoveNext
								Wend
							End If
						End If
					%>
				];

				var arr_vlr_planejado_acumulado = $.extend([], arr_vlr_planejado);

				if(arr_vlr_planejado_acumulado) {
					$.each(arr_vlr_planejado_acumulado, function(i, item) {
						if(i > 0)
							arr_vlr_planejado_acumulado[i] += arr_vlr_planejado_acumulado[i-1]
					})
				}

				var arr_vlr_medido = [
					<%
						If Not IsNull(cod_contrato) And Not IsEmpty(cod_contrato) Then
							strQ = "SELECT * FROM c_medido_contrato WHERE cod_contrato = " & cod_contrato & " ORDER BY dta_medicao ASC"

							Set rs_plan = Server.CreateObject("ADODB.Recordset")
								rs_plan.CursorLocation = 3
								rs_plan.CursorType = 3
								rs_plan.LockType = 1
								rs_plan.Open strQ, objCon, , , &H0001

							vlr_planejamento = 0

							If Not rs_plan.EOF Then
								While Not rs_plan.EOF
									Response.Write Replace(rs_plan.Fields.Item("vlr_medido").Value, ",",".") & ","
									rs_plan.MoveNext
								Wend
							End If
						End If
					%>
				];

				var arr_vlr_medido_acumulado = $.extend([], arr_vlr_medido);

				if(arr_vlr_medido_acumulado) {
					$.each(arr_vlr_medido_acumulado, function(i, item) {
						if(i > 0)
							arr_vlr_medido_acumulado[i] += arr_vlr_medido_acumulado[i-1]
					})
				}

				var arr_vlr_pago = [
					<%
						If Not IsNull(cod_contrato) And Not IsEmpty(cod_contrato) Then
							strQ = "SELECT * FROM c_pago_medido_contrato WHERE cod_contrato = " & cod_contrato

							Set rs_plan = Server.CreateObject("ADODB.Recordset")
								rs_plan.CursorLocation = 3
								rs_plan.CursorType = 3
								rs_plan.LockType = 1
								rs_plan.Open strQ, objCon, , , &H0001

							vlr_planejamento = 0

							If Not rs_plan.EOF Then
								While Not rs_plan.EOF
									Response.Write Replace(rs_plan.Fields.Item("vlr_pagamento").Value, ",",".") & ","
									rs_plan.MoveNext
								Wend
							End If
						End If
					%>
				];

				var arr_vlr_pago_acumulado = $.extend([], arr_vlr_pago);

				if(arr_vlr_pago_acumulado) {
					$.each(arr_vlr_pago_acumulado, function(i, item) {
						if(i > 0)
							arr_vlr_pago_acumulado[i] += arr_vlr_pago_acumulado[i-1]
					})
				}

			/****************************************************************************************
			 * FIM									DATASETS										*
			 ****************************************************************************************/

			dta_os = '<%=(rs_dados_contrato.Fields.Item("dta_os").Value)%>';
			dta_os = moment(dta_os, "DD/MM/YYYY");
			prz_total_execucao = '<%=(rs_dados_contrato.Fields.Item("prz_original_execucao_meses").Value)%>';
			prz_total_execucao = parseInt(prz_total_execucao, 10);
			dta_conclusao_obra = dta_os.add(prz_total_execucao, 'months');
			$(".dta_conclusao_obra").val(dta_conclusao_obra.format("DD/MM/YYYY"));
			dta_assinatura = '<%=(rs_dados_contrato.Fields.Item("dta_assinatura").Value)%>';
			dta_assinatura = moment(dta_assinatura, "DD/MM/YYYY");
			dta_vigencia = dta_assinatura.add(prz_total_execucao, 'months');
			$(".dta_vigencia").val(dta_vigencia.format("DD/MM/YYYY"));

			$('#chart-curva').highcharts({
				colors: colors,
				title: {
					text: 'Curvas Financeiras Acumuladas'
				},
				subtitle: {
					text: 'Medido x Previsão Contratual'
				},
				legend: {
					layout: 'horizontal',
					align: 'center',
					floating: false,
					backgroundColor: (Highcharts.theme && Highcharts.theme.legendBackgroundColor) || '#FFFFFF'
				},
				xAxis: {
					categories: months
				},
				yAxis: {
					title: {
						text: 'Valores (Milhares de R$)'
					}
				},
				tooltip: {
					headerFormat: '<span style="font-size:10px">{point.key}</span><table>',
					pointFormat: '<tr><td style="color:{series.color};padding:0">{series.name}: </td>' +
						'<td style="padding:0"><b>R$ {point.y:,.2f}</b></td></tr>',
					footerFormat: '</table>',
					shared: true,
					useHTML: true
				},
				credits: {
					enabled: false
				},
				plotOptions: {
					areaspline: {
						fillOpacity: 0.5
					}
				},
				series: [{
					name: 'Previsto Acumulado',
					data: arr_vlr_planejado_acumulado
				}, {
					name: 'Medido Acumulado',
					data: arr_vlr_medido_acumulado
				}]
			});

			moment.locale("pt-br");
			dta_corte = moment(dta_corte, "YYYY/MM/DD");
			mes_referencia = capitalizeFirstLetter(dta_corte.format("MMMM YYYY"));

			$('#chart-bar-basic').highcharts({
				colors: ['#7cb5ec'],
				chart: {
					type: 'bar'
				},
				title: {
					text: 'Meta Mensal Financeira - ' + mes_referencia
				},
				xAxis: {
					categories: items,
					title: {
						text: 'Itens do Contrato'
					}
				},
				yAxis: {
					title: {
						text: '% de avanço'
					}
				},
				tooltip: {
					valueSuffix: '%'
				},
				plotOptions: {
					bar: {
						dataLabels: {
							enabled: true
						}
					}
				},
				legend: {
					layout: 'horizontal',
					align: 'center',
					floating: false,
					backgroundColor: ((Highcharts.theme && Highcharts.theme.legendBackgroundColor) || '#FFFFFF')
				},
				credits: {
					enabled: false
				},
				series: [{
					name: 'Medido Acumulado',
					data: itemsData,
					dataLabels: {
						format: '{point.y:,.2f} %'
					}
				}]
			});

			$('#chart-bar-column').highcharts({
				colors: colors,
				chart: {
					type: 'column'
				},
				title: {
					text: 'Curva de Medições x Planilha Contratual'
				},
				xAxis: {
					categories: months,
					crosshair: true
				},
				yAxis: {
					min: 0,
					title: {
						text: 'Valores (Milhares de R$)'
					}
				},
				tooltip: {
					headerFormat: '<span style="font-size:10px">{point.key}</span><table>',
					pointFormat: '<tr><td style="color:{series.color};padding:0">{series.name}: </td>' +
						'<td style="padding:0"><b>R$ {point.y:,.2f}</b></td></tr>',
					footerFormat: '</table>',
					shared: true,
					useHTML: true
				},
				plotOptions: {
					column: {
						pointPadding: 0.2,
						borderWidth: 0
					}
				},
				credits: {
					enabled: false
				},
				series: [{
					name: 'Previsto',
					data: arr_vlr_planejado

				}, {
					name: 'Medido',
					data: arr_vlr_medido

				}, {
					name: 'Pago',
					data: arr_vlr_pago
				}]
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
				<a class="navbar-brand" href="#">SIG - Ficha Técnica da Obra</a>
			</div>

			<div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
				<ul class="nav navbar-nav navbar-right">
					<li><a href="javascript:window.history.back();"><i class="fa fa-chevron-left"></i> Voltar</a></li>
					<li><a href="#" class="print"><i class="fa fa-print"></i> Imprimir</a></li>
					<li><a href="#" class="expand"><i class="fa fa-expand"></i>&nbsp;&nbsp;Tela Cheia</a></li>
					<li><a href="<%= MM_Logout %>" class="sign-out"><i class="fa fa-sign-out"></i> Sair do Sistema</a></li>
				</ul>
			</div>
		</div>
	</nav>

	<div class="container-fluid">
		<div class="row">
			<div class="col-xs-12">
				<div class="panel panel-info">
					<div class="panel-heading">
						<h3 class="panel-title">Informações da Obra</h3>
					</div>
					<div class="panel-body">
						<form class="form-horizontal">
							<div class="form-group">
								<label class="col-lg-2 control-label">Município:</label>
								<div class="col-lg-5">
									<input type="text" class="form-control" readonly="readonly" value="<%=(rs_dados_obra.Fields.Item("municipio").Value)%>">
								</div>

								<label class="col-lg-1 control-label">Localidade:</label>
								<div class="col-lg-4">
									<input type="text" class="form-control" readonly="readonly" value="<%=(rs_dados_obra.Fields.Item("nome_empreendimento").Value)%>">
								</div>
							</div>

							<div class="form-group">
								<label class="col-lg-2 control-label">Diretoria de Bacia:</label>
								<div class="col-lg-3">
									<input type="text" class="form-control" readonly="readonly" value="<%=(rs_dados_obra.Fields.Item("bacia_daee").Value)%>">
								</div>

								<label class="col-lg-1 control-label">UGRHI:</label>
								<div class="col-lg-2">
									<input type="text" class="form-control" readonly="readonly" value="<%=(rs_dados_obra.Fields.Item("bacia_secretaria").Value)%>">
								</div>

								<label class="col-lg-3 control-label">Carga Orgânica Removida (ton/mês):</label>
								<div class="col-lg-1">
									<input type="text" class="form-control num text-center" readonly="readonly" value="<%=qtd_carga_organica_removida%>">
								</div>
							</div>

							<div class="form-group">
								<label class="col-lg-2 control-label">População Urbana (2010):</label>
								<div class="col-lg-3">
									<input type="text" class="form-control num" readonly="readonly" value="<%=(rs_dados_obra.Fields.Item("qtd_populacao_urbana_2010").Value)%>">
								</div>

								<label class="col-lg-4 control-label">Projeção de População (2030):</label>
								<div class="col-lg-3">
									<input type="text" class="form-control num" readonly="readonly" value="<%=qtd_populacao_urbana_2030%>">
								</div>
							</div>

							<div class="form-group">
								<label class="col-lg-2 control-label">Objeto da Obra:</label>
								<div class="col-lg-10">
									<textarea readonly="readonly" class="form-control" rows="5"><%=(rs_dados_obra.Fields.Item("Descrição da Intervenção FDE").Value)%></textarea>
								</div>
							</div>
						</form>
					</div>
				</div>
			</div>
		</div>

		<%
			If Not IsNull(cod_contrato) And Not IsEmpty(cod_contrato) Then
		%>
		<div class="row">
			<div class="col-xs-12">
				<div class="panel panel-info">
					<div class="panel-heading">
						<h3 class="panel-title">Informações do Contrato</h3>
					</div>
					<div class="panel-body">
						<form class="form-horizontal">
							<div class="form-group">
								<label class="col-lg-3 control-label">Contratada:</label>
								<div class="col-lg-9">
									<input type="text" class="form-control" readonly="readonly" value="<%=(rs_dados_contrato.Fields.Item("empresa_contratada").Value)%>">
								</div>
							</div>

							<div class="form-group">
								<label class="col-lg-3 control-label">Nº do Autos:</label>
								<div class="col-lg-3">
									<input type="text" class="form-control" readonly="readonly" value="<%=(rs_dados_contrato.Fields.Item("num_autos").Value)%>">
								</div>

								<label class="col-lg-3 control-label">Nº do Contrato:</label>
								<div class="col-lg-3">
									<input type="text" class="form-control" readonly="readonly" value="<%=(rs_dados_contrato.Fields.Item("num_contrato").Value)%>">
								</div>
							</div>

							<div class="form-group">
								<label class="col-lg-3 control-label">Data de Assinatura:</label>
								<div class="col-lg-3">
									<input type="text" class="form-control" readonly="readonly" value="<%=(rs_dados_contrato.Fields.Item("dta_assinatura").Value)%>">
								</div>

								<label class="col-lg-3 control-label">Vigente Até:</label>
								<div class="col-lg-3">
									<input type="text" class="form-control dta_vigencia" readonly="readonly" value="">
								</div>
							</div>

							<div class="form-group">
								<label class="col-lg-3 control-label">Data da Ordem de Serviço:</label>
								<div class="col-lg-3">
									<input type="text" class="form-control dta_os" readonly="readonly" value="<%=(rs_dados_contrato.Fields.Item("dta_os").Value)%>">
								</div>

								<label class="col-lg-3 control-label">Prazo de Execução:</label>
								<div class="col-lg-3">
									<input type="text" class="form-control" readonly="readonly" value="<%=(rs_dados_contrato.Fields.Item("prz_original_execucao_meses").Value)%> mese(s)">
								</div>
							</div>

							<div class="form-group">
								<label class="col-lg-3 control-label">Valor Original do Contrato:</label>
								<div class="col-lg-3">
									<input type="text" class="form-control vlr" readonly="readonly" value="<%=(rs_dados_contrato.Fields.Item("vlr_original").Value)%>">
								</div>

								<label class="col-lg-3 control-label">Prazo Total de Execução:</label>
								<div class="col-lg-3">
									<input type="text" class="form-control prz_total_execucao" readonly="readonly" value="<%=(rs_dados_contrato.Fields.Item("prz_original_execucao_meses").Value)%> mese(s)">
								</div>
							</div>

							<div class="form-group">
								<label class="col-lg-3 control-label">Valor Total do Contrato:</label>
								<div class="col-lg-3">
									<input type="text" class="form-control vlr" readonly="readonly" value="<%=(rs_dados_contrato.Fields.Item("vlr_aditado").Value)%>">
								</div>

								<label class="col-lg-3 control-label">Data de Conclusão da Obra:</label>
								<div class="col-lg-3">
									<input type="text" class="form-control dta_conclusao_obra" readonly="readonly" value="">
								</div>
							</div>

							<div class="form-group">
								<label class="col-lg-3 control-label">Aditamento Acumulado %:</label>
								<div class="col-lg-9">
									<input type="text" class="form-control prc" readonly="readonly"
										value="<%=((((rs_dados_contrato.Fields.Item("vlr_aditado").Value) - (rs_dados_contrato.Fields.Item("vlr_original").Value))/(rs_dados_contrato.Fields.Item("vlr_original").Value))*100)%>">
								</div>
							</div>
						</form>
					</div>
				</div>
			</div>
		</div>
		
		<div class="row">
			<div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
				<div class="panel panel-default">
					<div class="panel-body">
						<div id="chart-curva" style="min-width: 100%; height: 100%; margin: 0 auto"></div>
					</div>
				</div>
			</div>
		</div>

		<div class="row">
			<div class="col-xs-12 col-sm-6 col-md-6 col-lg-6">
				<div class="panel panel-default">
					<div class="panel-body">
						<div id="chart-bar-basic" style="min-width: 100%; height: 100%; margin: 0 auto"></div>
					</div>
				</div>
			</div>
			
			<div class="col-xs-12 col-sm-6 col-md-6 col-lg-6">
				<div class="panel panel-default">
					<div class="panel-body">
						<div id="chart-bar-column" style="min-width: 100%; height: 100%; margin: 0 auto"></div>
					</div>
				</div>
			</div>
		</div>

		<div class="row">
			<div class="col-xs-12">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title"><i class="fa fa-picture-o"></i> Galeria de Fotos</h3>
					</div>

					<div class="panel-body">
						<div class="row">
							<%
								strQ = "SELECT * FROM c_lista_fotos_obra WHERE PI = '" & cod_empreendimento & "'"

								Set rs_fotos = Server.CreateObject("ADODB.Recordset")
									rs_fotos.CursorLocation = 3
									rs_fotos.CursorType = 3
									rs_fotos.LockType = 1
									rs_fotos.Open strQ, objCon, , , &H0001

								If Not rs_fotos.EOF Then
									While Not rs_fotos.EOF
										pth_url = rs_fotos.Fields.Item("pth_arquivo").Value
										pth_url = Replace(pth_url, "\\10.0.75.125\intermultiplas.net\public\", "")
										pth_url = Replace(pth_url, "\", "/")
										img_url = pth_url & rs_fotos.Fields.Item("id_arquivo").Value &"_"& rs_fotos.Fields.Item("nme_arquivo").Value
							%>

							<div class="col-xs-12 col-sm-3 col-md-3">
								<div class="thumbnail">
									<img src="<%=(img_url)%>" alt="">
									<div class="caption">
										<!-- <h4>Thumbnail label</h4> -->
										<p><%=(rs_fotos.Fields.Item("dsc_observacoes").Value)%></p>
										<p>
											<a href="<%=(img_url)%>" rel="group" title="<%=(rs_fotos.Fields.Item("dsc_observacoes").Value)%>" class="btn btn-default btn-block btn-sm fancybox" role="button"><i class="fa fa-expand"></i> Ampliar imagem</a>
										</p>
									</div>
								</div>
							</div>
							<%
										rs_fotos.MoveNext
									Wend
								End If
							%>
						</div>
					</div>
				</div>
			</div>
		</div>
		
		<div class="row">
			<div class="col-xs-12">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title"><i class="fa fa-clock-o"></i> Histórico da Obra</h3>
					</div>

					<div class="panel-body">
						<table class="table table-history table-bordered table-hover table-striped table-condensed">
							<thead>
								<th class="text-center" width="150">Data do Registro</th>
								<th class="text-center">Registro</th>
							</thead>
							<tbody>
								<%
									strQ = "SELECT * FROM c_lista_acompanhamento WHERE PI = '" & cod_empreendimento & "' ORDER BY [Data do Registro] DESC"

									Set rs_fotos = Server.CreateObject("ADODB.Recordset")
										rs_fotos.CursorLocation = 3
										rs_fotos.CursorType = 3
										rs_fotos.LockType = 1
										rs_fotos.Open strQ, objCon, , , &H0001

									If Not rs_fotos.EOF Then
										While Not rs_fotos.EOF
								%>
								<tr>
									<td class="text-center text-middle"><%=(rs_fotos.Fields.Item("Data do Registro").Value)%></td>
									<td><%=(rs_fotos.Fields.Item("Registro").Value)%></td>
								</tr>
								<%
											rs_fotos.MoveNext
										Wend
									End If
								%>
							</tbody>
						</table>
						</div>
					</div>
				</div>
			</div>
		</div>
		<%
			Else
		%>
		<div class="row">
			<div class="col-xs-12">
				<div class="alert alert-warning"><i class="fa fa-warning"></i> Informações de contrato não encontradas no banco de dados! Verifique com o administrador do sistema.</div>
			</div>
		</div>
		<%
			End If
		%>
	</div>

</body>
</html>