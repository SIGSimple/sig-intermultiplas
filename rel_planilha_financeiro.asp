<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="logout.asp" -->
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
	<script type="text/javascript" src="js/underscore-min.js"></script>
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
	<script type="text/javascript" src="js/fancybox/jquery.fancybox.pack.js?v=2.1.5"></script>
	<script type="text/javascript" src="js/jquery.floatThead.min.js"></script>
	<script type="text/javascript">
		$(function(){
			moment.locale("pt-br");

			$("li a.print").on("click", function(){
				window.print();
			});

			$("li a.excel").on("click", function(){
				$("table#data").table2excel({
					name: "Planilha Financeiro Contratos"
				});
			});

			$.ajax({
				url: "query-to-json-util.asp?sql=SELECT * FROM c_planilha_financeiro_contrato ORDER BY dta_planejamento ASC",
				beforeSend: function() {
					$("#modalLoading").modal("show");
				},
				success: function(data, textStatus, jqXHR) {
					data = JSON.parse(data);

					var arr_datas = [];

					$.each(data, function(i, item) {
						if(!_.contains(arr_datas, item.dta_planejamento))
							arr_datas.push(item.dta_planejamento);
					});

					$.each(arr_datas, function(i, item) {
						var htmlMonth = "<th class='text-middle text-center' style='min-width: 150px;'>"+ moment(item, "DD/MM/YYYY").format("MMM/YYYY") +"</th>";
						$("thead tr").append(htmlMonth);
					});

					$("thead tr").append("<th class='text-middle text-center' style='min-width: 150px;'>Total</th>");
					$("thead tr").append("<th class='text-middle text-center' style='min-width: 150px;'>Última Medição</th>");
					$("thead tr").append("<th class='text-middle text-center' style='min-width: 150px;'>Previsão até Última Medição</th>");
					$("thead tr").append("<th class='text-middle text-center' style='min-width: 150px;'>Estágio Aproximado</th>");
					$("thead tr").append("<th class='text-middle text-center' style='min-width: 150px;'>Término Previsto</th>");
					$("thead tr").append("<th class='text-middle text-center' style='min-width: 150px;'>Término Projetado</th>");

					var arr_data_grouped = _.groupBy(data, "cod_contrato");

					$.each(arr_data_grouped, function(i, item) {
						var htmlLines = ""+
							"<tr>"+
								"<td rowspan='3' class='text-middle'>"+ item[0].municipio +"</td>"+
								"<td rowspan='3' class='text-middle'>"+ item[0].nome_empreendimento +"</td>";
							htmlLines += "<td>Previsto</td>";

							var dta_termino_previsto = "";
							$.each(item, function(x, xitem) {
								var hasNextItem = (x < (_.size(item)-1));

								if(!xitem.vlr_planejado){
									dta_termino_previsto = item[x-1].dta_planejamento;
									return false;
								}

								if(!hasNextItem) {
									dta_termino_previsto = xitem.dta_planejamento;
									return false;
								}
							});
							
							var dta_ultima_medicao = "";
							$.each(item, function(x, xitem) {
								if(!xitem.vlr_medido){
									dta_ultima_medicao = item[x-1].dta_planejamento;
									return false;
								}
							});

							var vlr_total_planejado 			= 0;
							var vlr_total_planejado_ult_medicao = 0;
							var vlr_total_medido 				= 0;
							var vlr_total_pago 					= 0;

							$.each(arr_datas, function(x, xitem) {
								var itemData = _.findWhere(data, {cod_contrato: parseInt(i, 10), dta_planejamento: xitem});
								var vlr = "0";

								if(itemData){
									if(itemData.vlr_planejado) {
										vlr = String(itemData.vlr_planejado).replace(".",",");
										vlr_total_planejado += itemData.vlr_planejado;

										if(moment(xitem, "DD/MM/YYYY") <= moment(dta_ultima_medicao, "DD/MM/YYYY"))
											vlr_total_planejado_ult_medicao += itemData.vlr_planejado;
									}
									
									if(itemData.vlr_medido)
										vlr_total_medido += itemData.vlr_medido;

									if(itemData.vlr_pagamento)
										vlr_total_pago += itemData.vlr_pagamento;
								}

								htmlLines += "<td class='vlr text-right'>R$ "+ $.number(vlr, 2, ",", ".") +"</td>";
							});

							var vlr_soma 		= 0;
							var vlr_soma_ant 	= 0;
							var dta_atual 		= "";
							var dta_anterior 	= "";
							var obj_min_prox 	= {vlr_soma: 0, dta_planejamento: ""};
							var obj_max_prox 	= {vlr_soma: 0, dta_planejamento: ""};
							var vlr_aprox 		= 0;
							var dta_aprox 		= "";

							$.each(arr_datas, function(x, xitem) {
								if(vlr_soma == vlr_total_medido) {
									vlr_aprox = vlr_soma;
									return false;
								} else if(vlr_soma > vlr_total_medido) {
									obj_min_prox.vlr_soma 			= vlr_soma_ant;
									obj_min_prox.dta_planejamento 	= dta_anterior;

									obj_max_prox.vlr_soma 			= vlr_soma;
									obj_max_prox.dta_planejamento 	= dta_atual;
									return false;
								}

								var itemData = _.findWhere(data, {cod_contrato: parseInt(i, 10), dta_planejamento: xitem});

								if(itemData){
									if(itemData.vlr_planejado){
										vlr_soma_ant = vlr_soma;
										dta_anterior = dta_atual;
										vlr_soma += itemData.vlr_planejado;
										dta_atual = itemData.dta_planejamento;
									}
								}
							});

							if(vlr_aprox == 0) {
								var vlr_resto_min = ((obj_min_prox.vlr_soma - vlr_total_medido) * (-1));
								var vlr_resto_max = (obj_max_prox.vlr_soma - vlr_total_medido);

								if(obj_min_prox.vlr_soma > obj_max_prox.vlr_soma){
									vlr_aprox = vlr_resto_min;
									dta_aprox = obj_min_prox.dta_planejamento;
								}
								else {
									vlr_aprox = vlr_resto_max;
									dta_aprox = obj_max_prox.dta_planejamento;
								}
							}

							var dta_termino_projetado;
							var dta_aprox_ = moment(dta_aprox, "DD/MM/YYYY");
							var dta_termino_previsto_ = moment(dta_termino_previsto, "DD/MM/YYYY");
							var diff = dta_termino_previsto_.diff(dta_aprox_, "months");
							var hoje = moment();
							dta_termino_projetado = hoje.add(diff, "months");

							vlr_total_planejado 			= String(vlr_total_planejado).replace(".",",");
							vlr_total_planejado_ult_medicao = String(vlr_total_planejado_ult_medicao).replace(".",",");
							vlr_total_medido 				= String(vlr_total_medido).replace(".",",");
							vlr_total_pago 					= String(vlr_total_pago).replace(".",",");

							htmlLines += "<td class='text-middle text-right'>R$ "+ $.number(vlr_total_planejado, 2, ",", ".") +"</td>";
							htmlLines += "<td rowspan='3' class='text-middle text-center'>"+ moment(dta_ultima_medicao, "DD/MM/YYYY").format("MMM/YYYY") +"</td>";
							htmlLines += "<td rowspan='3' class='text-middle text-right'>R$ "+ $.number(vlr_total_planejado_ult_medicao, 2, ",", ".") +"</td>";
							htmlLines += "<td rowspan='3' class='text-middle text-center'>"+ moment(dta_aprox, "DD/MM/YYYY").format("MMM/YYYY") +"</td>";
							htmlLines += "<td rowspan='3' class='text-middle text-center'>"+ moment(dta_termino_previsto, "DD/MM/YYYY").format("MMM/YYYY") +"</td>";
							htmlLines += "<td rowspan='3' class='text-middle text-center'>"+ dta_termino_projetado.format("DD/MM/YYYY") +"</td>";
							
							htmlLines += "</tr>";
							htmlLines += "<tr>";
							htmlLines += 	"<td>Medido</td>";

							$.each(arr_datas, function(x, xitem) {
								var itemData = _.findWhere(data, {cod_contrato: parseInt(i, 10), dta_planejamento: xitem});
								var vlr = "0";

								if(itemData){
									if(itemData.vlr_medido)
										vlr = String(itemData.vlr_medido).replace(".",",");
								}

								htmlLines += "<td class='vlr text-right'>R$ "+ $.number(vlr, 2, ",", ".") +"</td>";
							});
							
							htmlLines += "<td class='vlr text-right'>R$ "+ $.number(vlr_total_medido, 2, ",", ".") +"</td>";
							htmlLines += "</tr>";
							htmlLines += "<tr>";
							htmlLines += 	"<td>Pago</td>";

							$.each(arr_datas, function(x, xitem) {
								var itemData = _.findWhere(data, {cod_contrato: parseInt(i, 10), dta_planejamento: xitem});
								var vlr = "0";

								if(itemData){
									if(itemData.vlr_pagamento)
										vlr = String(itemData.vlr_pagamento).replace(".",",");
								}

								htmlLines += "<td class='vlr text-right'>R$ "+ $.number(vlr, 2, ",", ".") +"</td>";
							});


							htmlLines += "<td class='vlr text-right'>R$ "+ $.number(vlr_total_pago, 2, ",", ".") +"</td>";
							htmlLines += "</tr>";

						$("tbody").append(htmlLines);

						$("#modalLoading").modal("hide");

						$("table.table-data").floatThead({
							scrollingTop: 60
						});
					});
				},
				error: function(jqXHR, textStatus, errorThrown) {
					alert(errorThrown);
				}
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
				<a class="navbar-brand" href="#">SIG - Planilha Financeiro Contratos</a>
			</div>

			<div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
				<ul class="nav navbar-nav navbar-right">n
					<li><a href="javascript:window.history.back();"><i class="fa fa-chevron-left"></i> Voltar</a></li>
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
						<h3><strong>Planilha Financeiro Contratos</strong><br/><small>Listagem do Banco de Dados</small></h3>
					</div>
					<div class="col-xs-2"></div>
				</div>

				<hr/>

				<div class="row">
					<div class="col-xs-12">
						<table id="data" class="table table-data table-bordered table-condensed table-hover table-striped">
							<thead>
								<tr class="active">
									<th class="text-middle text-center" style="min-width: 250px;">Município</th>
									<th class="text-middle text-center" style="min-width: 250px;">Localidade</th>
									<th></th>
								</tr>
							</thead>
							<tbody>
								
							</tbody>
						</table>
					</div>
				</div>
			</div>
		</div>
	</div>

	<div class="modal fade" id="modalLoading" tabindex="-1" role="dialog" aria-labelledby="modalLoadingLabel" aria-hidden="true">
		<div class="modal-dialog modal-sm">
			<div class="modal-content">
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
					<h4 class="modal-title" id="modalLoadingLabel">Aguarde!</h4>
				</div>
				<div class="modal-body">
					<i class="fa fa-spinner fa-spin"></i> Buscando informações...
				</div>
			</div>
		</div>
	</div>
</body>
</html>