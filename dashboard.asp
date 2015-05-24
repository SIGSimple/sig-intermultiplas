<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="logout.asp" -->
<!DOCTYPE html>
<html>
<head>
	<title>:: DAEE ::</title>
	<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
	<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
	<link rel="stylesheet" type="text/css" href="css/bootstrap-flaty.min.css">
	<link rel="stylesheet" type="text/css" href="css/daee.css">
	<link rel="stylesheet" href="js/fancybox/jquery.fancybox.css?v=2.1.5" type="text/css" media="screen" />
	<link rel="stylesheet" href="//maxcdn.bootstrapcdn.com/font-awesome/4.3.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.11.3/themes/smoothness/jquery-ui.css">
	<script type="text/javascript" src="//code.jquery.com/jquery-1.11.2.min.js"></script>
	<script type="text/javascript" src="http://code.highcharts.com/highcharts.js"></script>
	<script type="text/javascript" src="http://code.highcharts.com/modules/exporting.js"></script>
	<script type="text/javascript" src="js/jquery.number.min.js"></script>
	<script type="text/javascript" src="js/jquery.table2excel.js"></script>
	<script type="text/javascript" src="js/jquery.lazyload.min.js"></script>
	<script type="text/javascript" src="js/moment.min.js"></script>
	<script type="text/javascript" src="js/fullscreen.js"></script>
	<script type="text/javascript" src="//code.jquery.com/ui/1.11.4/jquery-ui.min.js"></script>
	<script type="text/javascript" src="js/datepicker-pt-BR.js"></script>
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
	<script type="text/javascript" src="js/fancybox/jquery.fancybox.pack.js?v=2.1.5"></script>
	
	<script type="text/javascript">
		$(function(){
			$(".datepicker").datepicker($.datepicker.regional["pt-BR"]);

			$(".input-group-addon.calendar.calendar-de").on("click", function() {
				$("#dta_inicio.datepicker").trigger("focus");
			});

			$(".input-group-addon.calendar.calendar-ate").on("click", function() {
				$("#dta_ate.datepicker").trigger("focus");
			});

			$({ someValue: 0 }).animate({ someValue: parseFloat($('.count.vlr-investido').text()) }, {
				duration: 1000,
				easing: 'swing', // can be anything
				step: function () { // called on every step
					// Update the element's text with rounded-up value:
					$('.count.vlr-investido').text( $.number(Math.round(this.someValue), 0, ",", ".") );
				}
			});

			$({ someValue: 0 }).animate({ someValue: parseFloat($('.count.qtd-carga-organica').text()) }, {
				duration: 1000,
				easing: 'swing', // can be anything
				step: function () { // called on every step
					// Update the element's text with rounded-up value:
					$('.count.qtd-carga-organica').text( $.number(Math.round(this.someValue), 0, ",", ".") );
				}
			});

			$('#chart-municipios-atendidos-bacia').highcharts({
				chart: {
					plotBackgroundColor: null,
					plotBorderWidth: null,
					plotShadow: false
				},
				title: {
					text: 'Municípios Atendidos p/ Bacia'
				},
				tooltip: {
					pointFormat: '{series.name}: <b>{point.percentage:.1f}%</b>'
				},
				plotOptions: {
					pie: {
						allowPointSelect: true,
						cursor: 'pointer',
						dataLabels: {
							enabled: false
						},
						showInLegend: true
					}
				},
				series: [{
					type: 'pie',
					name: 'Browser share',
					data: [
						['Firefox',   45.0],
						['IE',       26.8],
						{
							name: 'Chrome',
							y: 12.8,
							sliced: true,
							selected: true
						},
						['Safari',    8.5],
						['Opera',     6.2],
						['Others',   0.7]
					]
				}],
				credits: {
					enabled: false
				},
				legend: {
					align: "right",
					layout: "vertical"
				}
			});

			$('#chart-investimento-governo-bacia').highcharts({
				chart: {
					plotBackgroundColor: null,
					plotBorderWidth: null,
					plotShadow: false
				},
				title: {
					text: 'Inv. do Governo p/ Bacia'
				},
				tooltip: {
					pointFormat: '{series.name}: <b>{point.percentage:.1f}%</b>'
				},
				plotOptions: {
					pie: {
						allowPointSelect: true,
						cursor: 'pointer',
						dataLabels: {
							enabled: false
						},
						showInLegend: true
					}
				},
				series: [{
					type: 'pie',
					name: 'Browser share',
					data: [
						['Firefox',   45.0],
						['IE',       26.8],
						{
							name: 'Chrome',
							y: 12.8,
							sliced: true,
							selected: true
						},
						['Safari',    8.5],
						['Opera',     6.2],
						['Others',   0.7]
					]
				}],
				credits: {
					enabled: false
				},
				legend: {
					align: "right",
					layout: "vertical"
				}
			});

			$('#chart-investimento-governo-ano').highcharts({
				colors: ['#f45b5b', '#7cb5ec'],
				chart: {
					type: 'areaspline'
				},
				title: {
					text: 'Inv. Governo p/ Ano'
				},
				legend: {
					layout: 'horizontal',
					align: 'center',
					x: 150,
					y: 100,
					floating: false,
					borderWidth: 1,
					backgroundColor: (Highcharts.theme && Highcharts.theme.legendBackgroundColor) || '#FFFFFF'
				},
				xAxis: {
					categories: [
						'2009',
						'2010',
						'2011',
						'2012',
						'2013',
						'2014',
						'2015'
					],
					plotBands: [{ // visualize the weekend
						from: 4.5,
						to: 6.5,
						color: 'rgba(68, 170, 213, .2)'
					}]
				},
				yAxis: {
					title: {
						text: 'Fruit units'
					}
				},
				tooltip: {
					shared: true,
					valueSuffix: ' units'
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
					name: 'John',
					data: [3, 4, 3, 5, 4, 10, 12]
				}, {
					name: 'Jane',
					data: [1, 3, 4, 3, 3, 5, 4]
				}]
			});

			$('#chart-populacao-atendida-status').highcharts({
				chart: {
					plotBackgroundColor: null,
					plotBorderWidth: null,
					plotShadow: false
				},
				title: {
					text: 'População Atendida p/ Status'
				},
				tooltip: {
					pointFormat: '{series.name}: <b>{point.percentage:.1f}%</b>'
				},
				plotOptions: {
					pie: {
						allowPointSelect: true,
						cursor: 'pointer',
						dataLabels: {
							enabled: false
						},
						showInLegend: true
					}
				},
				series: [{
					type: 'pie',
					name: 'Browser share',
					data: [
						['Firefox',   45.0],
						['IE',       26.8],
						{
							name: 'Chrome',
							y: 12.8,
							sliced: true,
							selected: true
						},
						['Safari',    8.5],
						['Opera',     6.2],
						['Others',   0.7]
					]
				}],
				credits: {
					enabled: false
				},
				legend: {
					align: "right",
					layout: "vertical"
				}
			});
		});
	</script>
	
	<style type="text/css">
		.table {
			margin-bottom: 0;
		}

		i.fa.fa-master-item {
			margin-top: 8px;
		}

		fieldset legend {
			margin-bottom: 0;
		}

		fieldset legend hr {
			margin-top: 0;
		}
	</style>

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
				<a class="navbar-brand" href="#">SIG - Painel de Indicadores</a>
			</div>

			<div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
				<ul class="nav navbar-nav navbar-right">
					<li><a href="javascript:window.history.back();"><i class="fa fa-chevron-left"></i> Voltar</a></li>
					<li><a href="#" class="expand"><i class="fa fa-expand"></i>&nbsp;&nbsp;Tela Cheia</a></li>
					<li><a href="<%= MM_Logout %>" class="sign-out"><i class="fa fa-sign-out"></i> Sair do Sistema</a></li>
				</ul>
			</div>
		</div>
	</nav>

	<div class="container">
		<fieldset>
			<legend>Totalizações do Universo de Atendimento <hr/></legend>

			<div class="row">
				<div class="col-xs-12 col-lg-3">
					<div class="alert alert-success text-center">
						<i class="fa fa-dollar fa-master-item fa-4x"></i>
						<div class="clearfix"></div>
						<hr/>
						<h2><strong class="count vlr-investido">5497</strong></h2>
						<div class="clearfix"></div>
						<p>
							<small>(R$ Milhões)</small>
							<br/>
							<strong>Total Investido pelo Governo</strong>
						</p>
					</div>
				</div>

				<div class="col-xs-12 col-lg-3">
					<div class="alert alert-warning text-center">
						<i class="fa fa-leaf fa-master-item fa-4x"></i>
						<div class="clearfix"></div>
						<hr/>
						<h2><strong class="count qtd-carga-organica">3450</strong></h2>
						<div class="clearfix"></div>
						<p>
							<small>(Ton/mês)</small>
							<br/>
							<strong>Carga Orgânica Retirada</strong>
						</p>
					</div>
				</div>

				<div class="col-xs-12 col-lg-3">
					<div class="alert alert-danger text-center">
						<i class="fa fa-check-square fa-master-item fa-4x"></i>
						<div class="clearfix"></div>
						<hr/>
						<h2><strong class="count qtd-municipios-atendidos">225</strong></h2>
						<div class="clearfix"></div>
						<p>
							<small>(Qtd. Total)</small>
							<br/>
							<strong>Municípios Atendidos</strong>
						</p>
					</div>
				</div>

				<div class="col-xs-12 col-lg-3">
					<div class="alert alert-info text-center">
						<i class="fa fa-flag fa-master-item fa-4x"></i>
						<div class="clearfix"></div>
						<hr/>
						<h2><strong class="count qtd-obras-potenciais">75</strong></h2>
						<div class="clearfix"></div>
						<p>
							<small>(Qtd. Total)</small>
							<br/>
							<strong>Obras Potencial de Atendimento</strong>
						</p>
					</div>
				</div>
			</div>

			<div class="row">
				<div class="col-xs-12 col-lg-4">
					<div id="chart-municipios-atendidos-bacia" style="min-width: 310px; height: 300px; max-width: 600px; margin: 0 auto"></div>
				</div>

				<div class="col-xs-12 col-lg-4">
					<div id="chart-investimento-governo-bacia" style="min-width: 310px; height: 300px; max-width: 600px; margin: 0 auto"></div>
				</div>

				<div class="col-xs-12 col-lg-4">
					<div id="chart-investimento-governo-ano" style="min-width: 310px; height: 300px; max-width: 600px; margin: 0 auto"></div>
				</div>
			</div>

			<div class="row">
				<div class="col-xs-12 col-lg-6">
					<div class="panel panel-default">
						<div class="panel-heading">
							<h3 class="panel-title">
								Totais por Status
							</h3>
						</div>

						<div class="panel-body">
							<table class="table table-bordered table-condensed table-striped table-hover">
								<thead>
									<th class="text-center">#</th>
									<th>Status</th>
									<th class="text-center">Municípios</th>
									<th class="text-center">Localidades</th>
								</thead>
								<tbody>
									<tr>
										<td class="text-center">1</td>
										<td>ATENDIMENTO CONCLUÍDO</td>
										<td class="text-center">94</td>
										<td class="text-center">96</td>
									</tr>

									<tr>
										<td class="text-center">2</td>
										<td>OBRAS EM ANDAMENTO</td>
										<td class="text-center">40</td>
										<td class="text-center">43</td>
									</tr>

									<tr>
										<td class="text-center">3</td>
										<td>PROGRAMADAS PARA ATENDIMENTO</td>
										<td class="text-center">17</td>
										<td class="text-center">18</td>
									</tr>

									<tr>
										<td class="text-center">4</td>
										<td>ATENDIMENTO FUTURO</td>
										<td class="text-center">61</td>
										<td class="text-center">68</td>
									</tr>
								</tbody>
							</table>
						</div>
					</div>
				</div>

				<div class="col-xs-12 col-lg-6">
					<div id="chart-populacao-atendida-status" style="min-width: 310px; height: 250px; max-width: 600px; margin: 0 auto"></div>
				</div>
			</div>
		</fieldset>

		<fieldset>
			<legend>
				Totalizações por Período
				<hr/>
			</legend>

			<div class="row">
				<form class="form-horizontal">
					<div class="form-group">
						<label class="col-lg-1 control-label">De:</label>
						<div class="col-lg-2">
							<div class="input-group">
								<input type="text" id="dta_inicio" class="form-control datepicker">
								<span class="input-group-addon calendar calendar-de cursor"><i class="fa fa-calendar"></i></span>
							</div>
						</div>

						<label class="col-lg-1 control-label">Até:</label>
						<div class="col-lg-2">
							<div class="input-group">
								<input type="text" id="dta_ate" class="form-control datepicker">
								<span class="input-group-addon calendar calendar-ate cursor"><i class="fa fa-calendar"></i></span>
							</div>
						</div>

						<div class="col-lg-2">
							<button type="button" class="btn btn-primary"><i class="fa fa-filter"></i> Filtrar</button>
						</div>
					</div>
				</form>
			</div>

			<div class="row">
				<div class="col-xs-12 col-lg-3">
					<div class="alert alert-default text-center">
						<i class="fa fa-play-circle fa-master-item fa-4x"></i>
						<div class="clearfix"></div>
						<hr/>
						<h2><strong class="count qtd-obras-iniciaram-periodo">100</strong></h2>
						<div class="clearfix"></div>
						<p>
							<small>(Qtd. Total)</small>
							<br/>
							<strong>Obras Iniciaram</strong>
							<br/>
							<small>(no período)</small>
						</p>
					</div>
				</div>

				<div class="col-xs-12 col-lg-3">
					<div class="alert alert-default text-center">
						<i class="fa fa-tasks fa-master-item fa-4x"></i>
						<div class="clearfix"></div>
						<hr/>
						<h2><strong class="count qtd-obras-concluidas-periodo">69</strong></h2>
						<div class="clearfix"></div>
						<p>
							<small>(Qtd. Total)</small>
							<br/>
							<strong>Obras em Andamento</strong>
							<br/>
							<small>(no período)</small>
						</p>
					</div>
				</div>

				<div class="col-xs-12 col-lg-3">
					<div class="alert alert-default text-center">
						<i class="fa fa-check-square-o fa-master-item fa-4x"></i>
						<div class="clearfix"></div>
						<hr/>
						<h2><strong class="count qtd-obras-concluidas-periodo">150</strong></h2>
						<div class="clearfix"></div>
						<p>
							<small>(Qtd. Total)</small>
							<br/>
							<strong>Obras Concluídas</strong>
							<br/>
							<small>(no período)</small>
						</p>
					</div>
				</div>

				<div class="col-xs-12 col-lg-3">
					<div class="alert alert-default text-center">
						<i class="fa fa-leaf fa-master-item fa-4x"></i>
						<div class="clearfix"></div>
						<hr/>
						<h2><strong class="count qtd-carga-organica-periodo">5629</strong></h2>
						<div class="clearfix"></div>
						<p>
							<small>(Ton/mês)</small>
							<br/>
							<strong>Carga Orgânica Retirada</strong>
							<br/>
							<small>(no período)</small>
						</p>
					</div>
				</div>
			</div>
		</fieldset>
	</div>
</body>
</html>