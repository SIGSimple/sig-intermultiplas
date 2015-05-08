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
	
	cod_situacao = rs_dados_obra.Fields.Item("cod_situacao").Value

	Select Case cod_situacao
		Case 39
			Response.Redirect("ficha-tecnica-obra-concluida.asp?cod_empreendimento="& cod_empreendimento)
		Case 44
		Case 41
			Response.Redirect("ficha-tecnica-obra-programada.asp?cod_empreendimento="& cod_empreendimento)
		Case 45
			Response.Redirect("ficha-tecnica-obra-potencial.asp?cod_empreendimento="& cod_empreendimento)
	End Select

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
	<script type="text/javascript" src="http://maps.googleapis.com/maps/api/js?libraries=places&sensor=false"></script>
	<script type="text/javascript" src="js/fullscreen.js"></script>
	<script type="text/javascript">
		// inicio variáveis de controle google maps ***********************************
		var geocoder;
		var map;
		var marker;
		var directionsDisplay;
		var directionsService;
		var initialLocation = new google.maps.LatLng(<%=(Replace(rs_dados_obra.Fields.Item("latitude_longitude").Value, "°",""))%>);
		// fim variáveis google maps ************************************

		function initMap(){
			// objeto de preferências e configurações do mapa do google maps
			var mapOptions = {
				streetViewControl: false,
				scrollWheel: false,
				zoom: 17,
				mapTypeId: google.maps.MapTypeId.SATELLITE
			};

			geocoder = new google.maps.Geocoder();

			// inicializa o google maps and adiciona-o ao html
			map = new google.maps.Map(document.getElementById("map-canvas"), mapOptions);
			map.setCenter(initialLocation);

			marker = new google.maps.Marker({
				map: map,
				position: initialLocation,
				title: '<%=(rs_dados_obra.Fields.Item("municipio").Value)%> - <%=(rs_dados_obra.Fields.Item("nome_empreendimento").Value)%>'
			});

			var content = '<strong>Município: </strong><%=(rs_dados_obra.Fields.Item("municipio").Value)%><br/><strong>Localidade: </strong><%=(rs_dados_obra.Fields.Item("nome_empreendimento").Value)%>';
			var infowindow = new google.maps.InfoWindow();

			google.maps.event.addListener(marker,'click', (function(marker,content,infowindow){
				return function() {
					infowindow.setContent(content);
					infowindow.open(map,marker);
				};
			})(marker,content,infowindow));
		}

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

			$("li a.map").on("click", function(){
				$('#modalMapaLocalizacao').modal("show");
				initMap();
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
				var dta_corte;

				if(arr_itens_ultima_medicao.length > 0)
					dta_corte = moment(arr_itens_ultima_medicao[0].dta_item, "DD/MM/YYYY").format("YYYY/MM/DD");

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

			 if(itemsData.length > 0){
				$("#chart-bar-basic").closest(".row").removeClass("hide");
			}
			
			if(months.length > 0){
				$('#chart-curva').closest(".row").removeClass("hide");
				$('#chart-bar-column').closest(".row").removeClass("hide");
			}

			dta_os = '<% If Not rs_dados_contrato.EOF Then Response.Write rs_dados_contrato.Fields.Item("dta_os").Value End If %>';
			dta_os = moment(dta_os, "DD/MM/YYYY");
			prz_total_original = '<% If Not rs_dados_contrato.EOF Then Response.Write rs_dados_contrato.Fields.Item("prz_original_contrato_meses").Value End If %>';
			prz_total_original = parseInt(prz_total_original, 10);

			prz_total_execucao = '<% If Not rs_dados_contrato.EOF Then Response.Write rs_dados_contrato.Fields.Item("prz_original_execucao_meses").Value End If %>';
			prz_total_execucao = parseInt(prz_total_execucao, 10);

			dta_conclusao_obra = dta_os.add(prz_total_execucao, 'months');
			$(".dta_conclusao_obra").val(dta_conclusao_obra.format("DD/MM/YYYY"));
			
			dta_assinatura = '<% If Not rs_dados_contrato.EOF Then Response.Write rs_dados_contrato.Fields.Item("dta_assinatura").Value End If%>';
			dta_assinatura = moment(dta_assinatura, "DD/MM/YYYY");
			dta_vigencia = dta_assinatura.add(prz_total_original, 'months');
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
			mes_referencia = capitalizeFirstLetter(dta_corte.format("MMMM"));
			ano_referencia = capitalizeFirstLetter(dta_corte.format("YYYY"));

			$('#chart-bar-basic').highcharts({
				colors: ['#7cb5ec'],
				chart: {
					type: 'bar'
				},
				title: {
					text: 'Previsto x Realizado Acumulado até '+ mes_referencia +' de ' + ano_referencia
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
					name: 'Prev. x Real. Acumulado',
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
					<!-- <li><a href="#" class="map"><i class="fa fa-map-marker"></i> Mapa de Localização</a></li> -->
					<li><a href="#" class="print"><i class="fa fa-print"></i> Imprimir</a></li>
					<li><a href="#" class="expand"><i class="fa fa-expand"></i>&nbsp;&nbsp;Tela Cheia</a></li>
					<li><a href="<%= MM_Logout %>" class="sign-out"><i class="fa fa-sign-out"></i> Sair do Sistema</a></li>
				</ul>
			</div>
		</div>
	</nav>

	<div class="container container-box ficha-andamento">
		<div class="panel panel-default">
			<div class="panel-body">
				<div class="row row-header">
					<div class="col-xs-3">
						<img src="img/governo_estado_500.png" class="img-responsive img-governo">
					</div>
					
					<div class="col-xs-7 text-center">
						<small><strong>Governo do Estado de São Paulo</strong></small>
						<br/>
						<small>Secretaria de Saneamento e Recursos Hídricos</small>
						<br/>
						<small>Departamento de Águas e Energia Elétrica</small>
					</div>

					<div class="col-xs-2 text-right">
						<img src="logo_daee.jpg" class="img-daee">
					</div>
				</div>

				<div class="row">
					<div class="col-xs-12">
						<table class="table table-condensed">
							<tbody>
								<tr class="warning">
									<td class="text-bold text-title">
										<%=(rs_dados_obra.Fields.Item("municipio").Value)%> - <%=(rs_dados_obra.Fields.Item("nome_empreendimento").Value)%>
									</td>
								</tr>
							</tbody>
						</table>
					</div>
				</div>
				
				<div class="row">
					<div class="col-xs-12">
						<div class="panel panel-default">
							<div class="panel-heading">
								<h3 class="panel-title"><i class="fa fa-building-o"></i> Informações da Obra</h3>
							</div>
							<div class="panel-body">
								<form class="form-horizontal">
									<div class="form-group">
										<label class="col-lg-2 control-label">Diretoria de Bacia:</label>
										<div class="col-lg-4">
											<input type="text" class="form-control" readonly="readonly" value="<%=(rs_dados_obra.Fields.Item("bacia_daee").Value)%>">
										</div>

										<label class="col-lg-2 control-label">UGRHI:</label>
										<div class="col-lg-4">
											<input type="text" class="form-control" readonly="readonly" value="<%=(rs_dados_obra.Fields.Item("bacia_secretaria").Value)%>">
										</div>
									</div>

									<div class="form-group">
										<label class="col-lg-2 control-label">População Urbana <small>(2010)</small>:</label>
										<div class="col-lg-2">
											<input type="text" class="form-control num" readonly="readonly" value="<%=(rs_dados_obra.Fields.Item("qtd_populacao_urbana_2010").Value)%>">
										</div>

										<label class="<% If rs_dados_obra.Fields.Item("cod_situacao").Value <> 39 Then %> col-lg-4 <% Else %> col-lg-2 <% End If %> control-label">Projeção de População <small>(2030)</small>:</label>
										<div class="col-lg-2">
											<input type="text" class="form-control num" readonly="readonly" value="<%=qtd_populacao_urbana_2030%>">
										</div>

										<%
											If rs_dados_obra.Fields.Item("cod_situacao").Value = 39 Then
										%>
										<label class="col-lg-2 control-label">Carga Orgânica Rem. <small>(ton/mês)</small>:</label>
										<div class="col-lg-2">
											<input type="text" class="form-control num text-center" readonly="readonly" value="<%=qtd_carga_organica_removida%>">
										</div>
										<%
											End If
										%>
									</div>

									<div class="form-group">
										<label class="col-lg-2 control-label">Objeto da Obra:</label>
										<div class="col-lg-10">
											<textarea readonly="readonly" class="form-control" rows="5"><%=(rs_dados_obra.Fields.Item("Descrição da Intervenção FDE").Value)%></textarea>
										</div>
									</div>

									<div class="form-group">
										<label class="col-lg-2 control-label">Situação:</label>
										<div class="col-lg-10">
											<%
												If Session("MM_UserAuthorization") = 8 Or Session("MM_UserAuthorization") = 9 Then
											%>
											<input type="text" class="form-control" readonly="readonly" value="<%=(rs_dados_obra.Fields.Item("dsc_situacao_externa").Value)%>">
											<%
												Else
											%>
											<input type="text" class="form-control" readonly="readonly" value="<%=(rs_dados_obra.Fields.Item("dsc_situacao_interna").Value)%>">
											<%
												End If
											%>
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
						<div class="panel panel-default">
							<div class="panel-heading">
								<h3 class="panel-title"><i class="fa fa-money"></i> Informações do Contrato</h3>
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
											<input type="text" class="form-control prz_total_execucao" readonly="readonly" value="<%=(rs_dados_contrato.Fields.Item("prz_original_contrato_meses").Value)%> mese(s)">
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
				
				<div class="row hide">
					<div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
						<div class="panel panel-default">
							<div class="panel-body">
								<div id="chart-curva" style="min-width: 100%; height: 100%; margin: 0 auto"></div>
							</div>
						</div>
					</div>
				</div>

				<div class="row hide">
					<div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
						<div class="panel panel-default">
							<div class="panel-body">
								<div id="chart-bar-column" style="min-width: 100%; height: 100%; margin: 0 auto"></div>
							</div>
						</div>
					</div>
				</div>

				<div class="row hide">
					<div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
						<div class="panel panel-default">
							<div class="panel-body">
								<div id="chart-bar-basic" style="min-width: 100%; height: 100%; margin: 0 auto"></div>
							</div>
						</div>
					</div>
				</div>
				<%
					End If
				%>

				<%
					strQueryFotos = "SELECT * FROM c_lista_fotos_obra WHERE PI = '" & cod_empreendimento & "'"

					Set rs_fotos = Server.CreateObject("ADODB.Recordset")
					rs_fotos.CursorLocation = 3
					rs_fotos.CursorType = 3
					rs_fotos.LockType = 1
					rs_fotos.Open strQueryFotos, objCon, , , &H0001

					If Not rs_fotos.EOF Then
				%>

				<div class="row">
					<div class="col-xs-12">
						<div class="panel panel-default">
							<div class="panel-heading">
								<h3 class="panel-title"><i class="fa fa-picture-o"></i> Galeria de Fotos</h3>
							</div>

							<div class="panel-body">
								<div class="row">
									<%
										While Not rs_fotos.EOF
											pth_url = LCase(rs_fotos.Fields.Item("pth_arquivo").Value)
											pth_url = Replace(pth_url, "\\10.0.75.125\intermultiplas.net\public\", "")
											pth_url = Replace(pth_url, "e:\home\programaagualimpa\web\", "")
											pth_url = Replace(pth_url, "\", "/")
											img_url = pth_url

											If Not rs_fotos.Fields.Item("flg_pmweb_file").Value Then
												img_url = img_url & rs_fotos.Fields.Item("cod_referencia").Value & "_"
											End If

											img_url = img_url & rs_fotos.Fields.Item("nme_arquivo").Value
									%>
									<div class="col-xs-12 col-sm-4 col-md-4">
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
									%>
								</div>
							</div>
						</div>
					</div>
				</div>
				
				<%
					End If

					If (Session("MM_UserAuthorization") = 8 OR Session("MM_UserAuthorization") = 9) Then
						If rs_dados_obra.Fields.Item("dsc_observacoes_relatorio_mensal").Value <> "" Then
				%>
				<div class="row">
					<div class="col-xs-12">
						<div class="panel panel-default">
							<div class="panel-heading">
								<h3 class="panel-title"><i class="fa fa-info-circle"></i> Situação da Obra</h3>
							</div>

							<div class="panel-body">
								<%=(rs_dados_obra.Fields.Item("dsc_observacoes_relatorio_mensal").Value)%>
							</div>
						</div>
					</div>
				</div>
				<%
						End If
					Else
						strQueryLicencas = "SELECT * FROM tb_licenca_ambiental INNER JOIN tb_tipo_licenca ON tb_tipo_licenca.id = tb_licenca_ambiental.cod_tipo_licenca WHERE cod_empreendimento = " & cod_empreendimento & ""

						Set rs_licencas = Server.CreateObject("ADODB.Recordset")
							rs_licencas.CursorLocation = 3
							rs_licencas.CursorType = 3
							rs_licencas.LockType = 1
							rs_licencas.Open strQueryLicencas, objCon, , , &H0001

						strQueryOutorgas = "SELECT * FROM tb_outorga WHERE cod_empreendimento = " & cod_empreendimento & ""

						Set rs_outorgas = Server.CreateObject("ADODB.Recordset")
							rs_outorgas.CursorLocation = 3
							rs_outorgas.CursorType = 3
							rs_outorgas.LockType = 1
							rs_outorgas.Open strQueryOutorgas, objCon, , , &H0001

						strQueryApps = "SELECT * FROM tb_app WHERE cod_empreendimento = " & cod_empreendimento & ""

						Set rs_apps = Server.CreateObject("ADODB.Recordset")
							rs_apps.CursorLocation = 3
							rs_apps.CursorType = 3
							rs_apps.LockType = 1
							rs_apps.Open strQueryApps, objCon, , , &H0001

						strQueryTCRAs = "SELECT * FROM tb_tcra WHERE cod_empreendimento = " & cod_empreendimento & ""

						Set rs_tcras = Server.CreateObject("ADODB.Recordset")
							rs_tcras.CursorLocation = 3
							rs_tcras.CursorType = 3
							rs_tcras.LockType = 1
							rs_tcras.Open strQueryTCRAs, objCon, , , &H0001

						If Not rs_licencas.EOF Or Not rs_outorgas.EOF Or Not rs_apps.EOF Or Not rs_tcras.EOF Then
				%>

				<div class="row">
					<div class="col-xs-12">
						<div class="panel panel-default">
							<div class="panel-heading">
								<h3 class="panel-title"><i class="fa fa-recycle"></i> Meio Ambiente</h3>
							</div>

							<div class="panel-body">
								<%
									If Not rs_licencas.EOF Then
								%>
								<div class="form-group">
									<label class="control-label">Licenças Ambientais</label>
									<table class="table table-history table-bordered table-hover table-striped table-condensed">
										<thead>
											<th>Nº Licença</th>
											<th class="text-center">Tipo de Licença</th>
											<th class="text-center">Data de Concessão</th>
											<th class="text-center">Data de Vencimento</th>
										</thead>
										<tbody>
											<%
												While Not rs_licencas.EOF
											%>
											<tr>
												<td><%=(rs_licencas.Fields.Item("num_licenca").Value)%></td>
												<td><%=(rs_licencas.Fields.Item("dsc_tipo_licenca").Value)%></td>
												<td class="text-center"><%=(rs_licencas.Fields.Item("dta_concessao").Value)%></td>
												<td class="text-center"><%=(rs_licencas.Fields.Item("dta_vencimento").Value)%></td>
											</tr>
											<%
													rs_licencas.MoveNext
												Wend
											%>
										</tbody>
									</table>
								</div>
								<%
									End If

									If Not rs_outorgas.EOF Then
								%>
								<div class="form-group">
									<label class="control-label">Outorgas</label>
									<table class="table table-history table-bordered table-hover table-striped table-condensed">
										<thead>
											<th>Nº Outorga</th>
											<th class="text-center">Data de Concessão</th>
											<th class="text-center">Data de Vencimento</th>
										</thead>
										<tbody>
											<%
												While Not rs_outorgas.EOF
											%>
											<tr>
												<td><%=(rs_outorgas.Fields.Item("num_outorga").Value)%></td>
												<td class="text-center"><%=(rs_outorgas.Fields.Item("dta_concessao").Value)%></td>
												<td class="text-center"><%=(rs_outorgas.Fields.Item("dta_vencimento").Value)%></td>
											</tr>
											<%
													rs_outorgas.MoveNext
												Wend
											%>
										</tbody>
									</table>
								</div>
								<%
									End If

									If Not rs_apps.EOF Then
								%>
								<div class="form-group">
									<label class="control-label">Autorizações p/ Intervenção em APPs</label>
									<table class="table table-history table-bordered table-hover table-striped table-condensed">
										<thead>
											<th>Nº App</th>
											<th class="text-center">Data de Concessão</th>
											<th class="text-center">Data de Vencimento</th>
										</thead>
										<tbody>
											<%
												While Not rs_apps.EOF
											%>
											<tr>
												<td><%=(rs_apps.Fields.Item("num_app").Value)%></td>
												<td class="text-center"><%=(rs_apps.Fields.Item("dta_concessao").Value)%></td>
												<td class="text-center"><%=(rs_apps.Fields.Item("dta_vencimento").Value)%></td>
											</tr>
											<%
													rs_apps.MoveNext
												Wend
											%>
										</tbody>
									</table>
								</div>
								<%
									End If

									If Not rs_tcras.EOF Then
								%>
								<div class="form-group">
									<label class="control-label">TCRAs</label>
									<table class="table table-history table-bordered table-hover table-striped table-condensed">
										<thead>
											<th>Cod. TCRA</th>
											<th class="text-center">Data de Concessão</th>
											<th class="text-center">Data de Vencimento</th>
										</thead>
										<tbody>
											<%
												While Not rs_tcras.EOF
											%>
											<tr>
												<td><%=(rs_tcras.Fields.Item("cod_tcra").Value)%></td>
												<td class="text-center"><%=(rs_tcras.Fields.Item("dta_concessao").Value)%></td>
												<td class="text-center"><%=(rs_tcras.Fields.Item("dta_vencimento").Value)%></td>
											</tr>
											<%
													rs_tcras.MoveNext
												Wend
											%>
										</tbody>
									</table>
								</div>
								<%
									End If
								%>
							</div>
						</div>
					</div>
				</div>

				<%
						End If

						strQueryHistoricoObra = "SELECT * FROM c_lista_acompanhamento WHERE PI = '" & cod_empreendimento & "' ORDER BY [Data do Registro] DESC"

						Set rs_historicoObra = Server.CreateObject("ADODB.Recordset")
							rs_historicoObra.CursorLocation = 3
							rs_historicoObra.CursorType = 3
							rs_historicoObra.LockType = 1
							rs_historicoObra.Open strQueryHistoricoObra, objCon, , , &H0001

						If Not rs_historicoObra.EOF Then
				%>

				<div class="row">
					<div class="col-xs-12">
						<div class="panel panel-default">
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
											While Not rs_historicoObra.EOF
										%>
										<tr>
											<td class="text-center text-middle"><%=(rs_historicoObra.Fields.Item("Data do Registro").Value)%></td>
											<td><%=(rs_historicoObra.Fields.Item("Registro").Value)%></td>
										</tr>
										<%
												rs_historicoObra.MoveNext
											Wend
										%>
									</tbody>
								</table>
								</div>
							</div>
						</div>
					</div>
				</div>
				<%
						End If
					End If
				%>
			</div>
		</div>
	</div>

	<div class="modal fade" id="modalMapaLocalizacao" tabindex="-1" role="dialog" aria-labelledby="modalMapaLocalizacaoLabel" aria-hidden="true">
		<div class="modal-dialog modal-sm">
			<div class="modal-content">
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
					<h4 class="modal-title" id="modalMapaLocalizacaoLabel"><i class="fa fa-map-marker"></i> Mapa de Localização</h4>
				</div>
				<div class="modal-body">
					<div id="map-canvas"></div>
				</div>
			</div>
		</div>
	</div>
</body>
</html>