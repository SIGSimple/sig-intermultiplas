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
			$(".fancybox").fancybox();

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
						$("#txt-objeto-obra").text((dadosObra['Descrição da Intervenção FDE']) ? dadosObra['Descrição da Intervenção FDE'] : "");
						$("#txt-empresa-contratada").text(dadosObra['empresa_contratada']);
						$("#txt-prazo-execucao").text((dadosObra.dta_assinatura) ? dtaAssinatura +" à "+ dtaVigencia : "");
						$("#txt-investimento-governo").text(dadosObra['Valor do Contrato']);
						$("#txt-pop-2010").text(dadosObra['qtd_populacao_urbana_2010']);
						$("#txt-pop-2030").text(pop2030);
						$("#txt-beneficio-obra").text((dadosObra['dsc_resultado_obtido']) ? dadosObra['dsc_resultado_obtido'] : "");
						$("#txt-beneficio-ambiental").text("Carga Orgânica Retirada: "+ cargaOrganizaRetirada +" (toneladas/mês)");
						$("#txt-parceria-realizacao").text((dadosObra['dsc_parceria_realizacao']) ? dadosObra['dsc_parceria_realizacao'] : "");

						adjustNumLayout();
						adjustVlrLayout();

						$("#modalFichaTecnica").modal("show");
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
				<a class="navbar-brand" href="#">SIG - Ficha Técnica de Obra Concluída</a>
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
									<td class="text-middle text-center text-bold" rowspan="2" width="100">
										Obra
									</td>
									<td id="txt-objeto-obra"></td>
								</tr>
								<tr>
									<td>
										<span class="text-bold">Empresa Executora: </span> <span id="txt-empresa-contratada"></span>
										<br/>
										<span class="text-bold">Prazo de Execução: </span> <span id="txt-prazo-execucao"></span>
									</td>
								</tr>
							</tbody>
						</table>

						<table class="table table-bordered table-condensed">
							<tbody>
								<tr>
									<td class="text-middle text-bold">
										Recursos do Governo do Estado de São Paulo
									</td>
									<td class="text-middle text-right vlr" id="txt-investimento-governo"></td>
								</tr>
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

						<table class="table table-bordered table-condensed">
							<tbody>
								<tr>
									<td class="text-middle text-bold" width="150">
										Benefício da Obra
									</td>
									<td id="txt-beneficio-obra"></td>
								</tr>
								<tr>
									<td class="text-middle text-bold" width="150">
										Benefício Ambiental
									</td>
									<td class="text-middle" id="txt-beneficio-ambiental"></td>
								</tr>
								<tr>
									<td class="text-middle text-bold" width="150">
										Parceria/Realização
									</td>
									<td id="txt-parceria-realizacao"></td>
								</tr>
							</tbody>
						</table>
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

										dta_registro = ""
										dsc_registro = ""

										If Not rs_fotos.EOF Then
											While Not rs_fotos.EOF
												dta_registro = rs_fotos.Fields.Item("Data do Registro").Value
												dsc_registro = rs_fotos.Fields.Item("Registro").Value
												pth_url = rs_fotos.Fields.Item("pth_arquivo").Value
												pth_url = Replace(pth_url, "\\10.0.75.125\intermultiplas.net\public\", "")
												pth_url = Replace(pth_url, "\", "/")
												img_url = pth_url & rs_fotos.Fields.Item("id_arquivo").Value &"_"& rs_fotos.Fields.Item("nme_arquivo").Value
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