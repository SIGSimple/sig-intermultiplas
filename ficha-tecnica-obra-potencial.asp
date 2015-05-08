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
	<script type="text/javascript" src="js/underscore-min.js"></script>
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
	<script type="text/javascript" src="http://code.highcharts.com/highcharts.js"></script>
	<script type="text/javascript" src="http://code.highcharts.com/modules/exporting.js"></script>
	<script type="text/javascript" src="js/moment.min.js"></script>
	<script type="text/javascript" src="js/fullscreen.js"></script>
	<script type="text/javascript">
		function adjustNumLayout() {
			$.each($(".num"), function(i, item){
				//$(item).val($.number($(item).val(), 0, ",", "."));
				$(item).text($.number($(item).text(), 0, ",", "."));
			});
		}

		function adjustVlrLayout() {
			$.each($(".vlr"), function(i, item){
				// $(item).val($.number($(item).val(), 0, ",", "."));
				if($(item).text() != "")
					$(item).text("R$ " + $.number($(item).text(), 2, ",", "."));
			});
		}

		function adjustPrcLayout() {
			$.each($(".prc"), function(i,item){
				$(item).text($.number(( parseFloat($(item).text()) * 100 ), 2, ",", ".") + "%");
			});
		}

		$(function(){
			$("li a.print").on("click", function(){
				window.print();
			});
			
			var cod_empreendimento = '<%=(Request.QueryString("cod_empreendimento"))%>'

			$.ajax({
				url: "query-to-json-util.asp",
				method: "POST",
				data: {
					sql: "SELECT * FROM c_lista_dados_obras WHERE PI = '" + cod_empreendimento + "'"
				},
				beforeSend: function() {
					$("#modalLoading").modal("show");
				},
				success: function(data, textStatus, jqXHR){
					data = JSON.parse(data);

					if(data.length > 0) {
						var dadosObra = data[0];

						var pop2030;
						pop2030 = dadosObra.qtd_populacao_urbana_2010 * 1.25;
						pop2030 = parseFloat((pop2030/100).toFixed(0));
						pop2030 = parseFloat((pop2030 * 100).toFixed(0));

						var dtaAssinatura = moment(dadosObra.dta_assinatura, "DD/MM/YYYY");
						var dtaVigencia = dtaAssinatura.add(dadosObra.prz_original_execucao_meses, 'months').format("MM/YYYY")
						dtaAssinatura = moment(dadosObra.dta_assinatura, "DD/MM/YYYY").format("MM/YYYY");

						var cargaOrganizaRetirada = (pop2030 * 0.0018).toFixed(2);

						$("#txt-municipio-localidade").text(dadosObra['Município'] +" - "+ dadosObra['nome_empreendimento']);
						$("#txt-nome-prefeitura").text(dadosObra['prefeitura']);
						$("#txt-nome-prefeito").text(dadosObra['prefeito']);
						$("#txt-nome-bacia-daee").text(dadosObra['bacia_daee']);
						$("#txt-situacao").text(dadosObra['desc_situacao_externa']);
						$("#txt-pop-2010").text(dadosObra['qtd_populacao_urbana_2010']);
						$("#txt-pop-2030").text(pop2030);

						adjustNumLayout();
						adjustVlrLayout();
					}
					else
						alert("Nenhuma informação encontrada!");

					$("#modalLoading").modal("hide");
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

	<div class="container container-box">
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
									<td class="text-bold text-title" id="txt-municipio-localidade"></td>
								</tr>
							</tbody>
						</table>
					</div>
				</div>

				<div class="row">
					<div class="col-xs-12">
						<table class="table table-bordered table-condensed">
							<tbody>
								<tr>
									<td class="text-center">Prefeitura</td>
									<td class="text-center">Prefeito</td>
								</tr>
								<tr class="active">
									<td class="text-center text-bold"><small id="txt-nome-prefeitura"></small></td>
									<td class="text-center text-bold"><small id="txt-nome-prefeito"></small></td>
								</tr>
							</tbody>
						</table>

						<table class="table table-bordered table-condensed">
							<tbody>
								<tr>
									<td class="text-middle text-bold" width="200">Diretoria de Bacia - DAEE</td>
									<td class="text-middle" id="txt-nome-bacia-daee"></td>
								</tr>
							</tbody>
						</table>

						<table class="table table-bordered table-condensed">
							<tbody>
								<tr>
									<td class="text-middle text-bold">
										Situação
									</td>
									<td class="text-right" id="txt-situacao"></td>
								</tr>
							</tbody>
						</table>

						<table class="table table-bordered table-condensed">
							<tbody>
								<tr>
									<td class="text-middle text-bold">
										População Beneficiada em 2010
									</td>
									<td class="text-middle text-right num" id="txt-pop-2010"></td>
								</tr>
								<tr>
									<td class="text-middle text-bold">
										População Beneficiada em Demanda Futura - 2030
									</td>
									<td class="text-middle text-right num" id="txt-pop-2030"></td>
								</tr>
							</tbody>
						</table>
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