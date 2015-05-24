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
	<style type="text/css">
		body { padding-top: 70px !important; }
	</style>
	<script type="text/javascript">
		function adjustNumLayout() {
			$.each($(".num"), function(i, item){
				$(item).val($.number($(item).val(), 0, ",", "."));
				$(item).text($.number($(item).text(), 0, ",", "."));
			});
		}

		function listEmpreendimentosTable(localidades) {
			$.each(localidades, function(i, item) {
				var link = "";

				switch(item.cod_situacao_interna){
					case 39: // Concluída
						link = "ficha-tecnica-obra-concluida.asp?cod_empreendimento=" + item.cod_empreendimento;
						break;
					case 40: 
					case 42: // Em atendimento
					case 43:
						link = "ficha-tecnica-obra.asp?cod_empreendimento=" + item.cod_empreendimento;
						break;
					case 44: // Programado para atendimento
					case 41: // Não atendido
						link = "ficha-tecnica-obra-programada.asp?cod_empreendimento=" + item.cod_empreendimento;
						break;
					case 45: // Atendimento potencial
						link = "ficha-tecnica-obra-potencial.asp?cod_empreendimento=" + item.cod_empreendimento;
						break;
				}

				var itemLayout = '<tr>'+
									'<td class="text-middle">'+ item.nme_empreendimento + '</td>'+
									'<td class="text-middle">'+ item.dsc_situacao_interna + '</td>'+
									'<td class="text-middle"><a href="'+ link +'" class="btn btn-sm btn-primary">Ficha Técnica</a></td>'+
								'</tr>';
				$("table.table-localidades tbody").append(itemLayout);
			});
		}

		$(function() {
			$.ajax({
				url: "query-to-json-util.asp",
				method: "POST",
				data: {
					sql: "SELECT * FROM c_lista_predios WHERE id_predio = <%=(cod_municipio)%>"
				},
				beforeSend: function() {
					$("#modalLoading").modal("show");
				},
				success: function(data, textStatus, jqXHR){
					data = JSON.parse(data);

					if(data.length > 0) {
						var dadosMunicipio = data[0];

						var pop2030;
							pop2030 = dadosMunicipio.qtd_populacao_urbana_2010 * 1.25;
							pop2030 = parseFloat((pop2030/100).toFixed(0));
							pop2030 = parseFloat((pop2030 * 100).toFixed(0));

						$("#txt-dsc-municipio").val(dadosMunicipio['nme_prefeitura']);
						$("#txt-nme-bacia-daee").val(dadosMunicipio['bacia_daee']);
						$("#txt-nme-bacia-secretaria").val(dadosMunicipio['bacia_secretaria']);
						$("#txt-pop-2010").val(dadosMunicipio['qtd_populacao_urbana_2010']);
						$("#txt-pop-2030").val(pop2030);
						$("#txt-nme-prefeito").val(dadosMunicipio['nme_prefeito']);
						$("#txt-atendido-sabesp").val(dadosMunicipio['dsc_concessao']);

						adjustNumLayout();

						$("#btn-ficha-completa").attr("href", "informacao-municipio.asp?cod_municipio=<%=(cod_municipio)%>");

						$("#modalLoading").modal("hide");
						$("#modalInfoMunicipio").modal("show");
					}
					else
						alert("Nenhuma informação encontrada!");
				},
				error: function(jqXHR, textStatus, errorThrown){
					console.log(jqXHR, textStatus, errorThrown);
				}
			});

			$.ajax({
				url: "query-to-json-util.asp",
				method: "POST",
				data: {
					sql: "SELECT * FROM c_lista_pi WHERE id_predio = <%=(cod_municipio)%> AND cod_situacao_externa is not null"
				},
				beforeSend: function() {
					
				},
				success: function(data, textStatus, jqXHR){
					data = JSON.parse(data);

					if(data.length > 0) {
						var items = [];

						$.each(data, function(i, item) {
							var itemData = {
								cod_empreendimento: item.PI,
								nme_empreendimento: item.nome_empreendimento,
								cod_situacao_interna: item.cod_situacao,
								dsc_situacao_interna: item.dsc_situacao_interna
							};

							items.push(itemData);
						});

						listEmpreendimentosTable(items);
					}
				},
				error: function(jqXHR, textStatus, errorThrown){
					console.log(jqXHR, textStatus, errorThrown);
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
				<a class="navbar-brand" href="#">SIG - Informações do Município (Resumida)</a>
			</div>

			<div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
				<ul class="nav navbar-nav navbar-right">
					<li><a href="javascript:window.history.back();"><i class="fa fa-chevron-left"></i> Voltar</a></li>
					<li><a href="informacao-municipio.asp?cod_municipio=<%=(cod_municipio)%>"><i class="fa fa-list-alt"></i> Detalhes do Município</a></li>
					<li><a href="#" class="print"><i class="fa fa-print"></i> Imprimir</a></li>
					<li><a href="#" class="expand"><i class="fa fa-expand"></i>&nbsp;&nbsp;Tela Cheia</a></li>
					<li><a href="<%= MM_Logout %>" class="sign-out"><i class="fa fa-sign-out"></i> Sair do Sistema</a></li>
				</ul>
			</div>
		</div>
	</nav>

	<div class="container container-box">
		<div class="panel panel-default">
			<div class="panel-body">
				<div class="row">
					<div class="col-xs-12">
						<div class="form-group">
							<label class="control-label">Município</label>
							<input class="form-control input-sm" readonly="readonly" id="txt-dsc-municipio">
						</div>
					</div>
				</div>

				<div class="row">
					<div class="col-xs-6">
						<label class="control-label">Dir. de Bacia - DAEE</label>
						<input class="form-control input-sm" readonly="readonly" id="txt-nme-bacia-daee">
					</div>

					<div class="col-xs-6">
						<label class="control-label">UGRHI</label>
						<input class="form-control input-sm" readonly="readonly" id="txt-nme-bacia-secretaria">
					</div>
				</div>

				<div class="row">
					<div class="col-xs-6">
						<label class="control-label">População Urbana 2010 (IBGE)</label>
						<input class="form-control input-sm num" readonly="readonly" id="txt-pop-2010">
					</div>

					<div class="col-xs-6">
						<label class="control-label">Projeção de População (2030)</label>
						<input class="form-control input-sm num" readonly="readonly" id="txt-pop-2030">
					</div>
				</div>

				<div class="row">
					<div class="col-xs-12">
						<label class="control-label">Nome do Prefeito</label>
						<input class="form-control input-sm" readonly="readonly" id="txt-nme-prefeito">
					</div>
				</div>

				<div class="row">
					<div class="col-xs-5">
						<label class="control-label">Atendido por</label>
						<input class="form-control input-sm" readonly="readonly" id="txt-atendido-sabesp">
					</div>
				</div>

				<div class="row">
					<div class="col-xs-12">
						<div class="table-responsive">
							<table class="table table-localidades table-bordered table-condensed table-striped table-hover">
								<thead>
									<tr class="info">
										<th>Localidade</th>
										<th>Situação</th>
										<th width="50"></th>
									</tr>
								</thead>
								<tbody></tbody>
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
	</div>

</body>
</html>