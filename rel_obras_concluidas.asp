<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="logout.asp" -->
<!--#include file="functions.asp" -->
<%
	Response.CharSet = "UTF-8"

	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
		objCon.Open MM_cpf_STRING

	Set id = Request.QueryString("id")

	strQueryDadosGerais = "SELECT tb_info_emp_concluidos.*, tb_pi.latitude_longitude FROM tb_pi RIGHT JOIN tb_info_emp_concluidos ON tb_pi.PI = tb_info_emp_concluidos.num_autos WHERE tb_info_emp_concluidos.id = " & id

	Set rsDadosGerais = Server.CreateObject("ADODB.Recordset")
		rsDadosGerais.CursorLocation = 3
		rsDadosGerais.CursorType = 3
		rsDadosGerais.LockType = 1
		rsDadosGerais.Open strQueryDadosGerais, objCon, , , &H0001

	strQueryFotos = "SELECT tb_info_emp_concluidos.num_autos, tb_info_emp_concluidos_arquivo.nme_arquivo, tb_info_emp_concluidos_arquivo.pth_arquivo, tb_info_emp_concluidos_arquivo.dsc_observacoes FROM tb_info_emp_concluidos_arquivo INNER JOIN tb_info_emp_concluidos ON tb_info_emp_concluidos_arquivo.cod_referencia = tb_info_emp_concluidos.id WHERE tb_info_emp_concluidos.id = " & id

	Set rsFotos = Server.CreateObject("ADODB.Recordset")
	rsFotos.CursorLocation = 3
	rsFotos.CursorType = 3
	rsFotos.LockType = 1
	rsFotos.Open strQueryFotos, objCon, , , &H0001

	num_autos = rsDadosGerais.Fields.Item("num_autos").Value
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
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
	<script type="text/javascript" src="http://maps.googleapis.com/maps/api/js?libraries=places&sensor=false"></script>
	<script type="text/javascript" src="http://code.highcharts.com/highcharts.js"></script>
	<script type="text/javascript" src="http://code.highcharts.com/modules/exporting.js"></script>
	<script type="text/javascript" src="js/underscore-min.js"></script>
	<script type="text/javascript" src="js/jquery.number.min.js"></script>
	<script type="text/javascript" src="js/fullscreen.js"></script>
	<script type="text/javascript" src="js/moment.min.js"></script>
	<script type="text/javascript" src="js/fancybox/jquery.fancybox.pack.js?v=2.1.5"></script>
	<script type="text/javascript" src="js/loadMapaObra.js"></script>
	<script type="text/javascript" src="js/common.js"></script>
	<script type="text/javascript">
		function getLatitudeLongitude() {
			return "<%=(rsDadosGerais.Fields.Item("latitude_longitude").Value)%>";
		}

		$(function(){
			$(".fancybox").fancybox();

			$('#modalMapa').on('show.bs.modal', function() {
				resizeMap();
			});

			$("li a.map").on("click", function(){
				$("#modalMapa").modal("show");
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
				<a class="navbar-brand" href="#">SIG - Situação das Obras Concluídas</a>
			</div>

			<div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
				<ul class="nav navbar-nav navbar-right">
					<%
						If rsDadosGerais.Fields.Item("latitude_longitude").Value <> "" Then
					%>
					<li><a href="#" class="map"><i class="fa fa-map-marker"></i> Mapa</a></li>
					<%
						End If
					%>
					<li><a href="javascript:window.history.back();"><i class="fa fa-chevron-left"></i> Voltar</a></li>
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
					<div class="col-xs-3 col-sm-3 col-md-3 col-lg-3">
						<img src="img/governo_estado_500.png" class="img-responsive img-governo">
					</div>
					
					<div class="col-xs-7 col-sm-7 col-md-7 col-lg-7 text-center">
						<strong>Governo do Estado de São Paulo</strong>
						<br/>
						<small>Secretaria de Saneamento e Recursos Hídricos</small>
						<br/>
						<small>Departamento de Águas e Energia Elétrica</small>
					</div>

					<div class="col-xs-2 col-sm-2 col-md-2 col-lg-2 text-right">
						<img src="logo_daee.jpg" class="img-daee">
					</div>
				</div>

				<div class="row">
					<div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
						<table class="table table-condensed">
							<tbody>
								<tr class="info">
									<td class="text-bold text-title text-center">
										<%=(rsDadosGerais.Fields.Item("nme_municipio").Value)%> - <%=(rsDadosGerais.Fields.Item("nme_localidade").Value)%>
									</td>
								</tr>
							</tbody>
						</table>
					</div>
				</div>

				<div class="row">
					<div class="col-xs-12 clearfix">
						<form class="form form-horizontal">
							<div class="form-group">
								<label class="sr-only control-label"></label>
								<div class="col-xs-12 clearfix">
									<%
										pth_url_vistoria = LCase(rsDadosGerais.Fields.Item("pth_arquivo_vistoria").Value)
										pth_url_vistoria = Replace(pth_url_vistoria, "\\10.0.75.125\intermultiplas.net\public\", "")
										pth_url_vistoria = Replace(pth_url_vistoria, "e:\home\programaagualimpa\web\", "")
										pth_url_vistoria = Replace(pth_url_vistoria, "\", "/")
										pdf_url = pth_url_vistoria & "/" & rsDadosGerais.Fields.Item("nme_arquivo_vistoria").Value
									%>
									<a href="<%=(pdf_url)%>" target="_blank" class="btn btn-primary pull-right"><i class="fa fa-download"></i> RVOC</a>
								</div>
							</div>
						</form>
					</div>
				</div>

				<div class="row">
					<div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
						<div class="panel panel-default">
							<div class="panel-heading">
								<h3 class="panel-title"><i class="fa fa-building-o"></i> Informações da Obra</h3>
							</div>
							<div class="panel-body">
								<form class="form form-horizontal">
									<div class="form-group">
										<label class="col-xs-2 col-sm-2 col-md-2 col-lg-2 control-label">Diretoria de Bacia:</label>
										<div class="col-xs-4 col-sm-4 col-md-4 col-lg-4">
											<input type="text" class="form-control" readonly="readonly" value="<%=(rsDadosGerais.Fields.Item("nme_bacia_daee").Value)%>">
										</div>

										<label class="col-xs-1 col-sm-1 col-md-1 col-lg-1 control-label">UGRHI:</label>
										<div class="col-xs-5 col-sm-5 col-md-5 col-lg-5">
											<input type="text" class="form-control" readonly="readonly" value="<%=(rsDadosGerais.Fields.Item("nme_bacia_secretaria").Value)%>">
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
							<div class="panel-heading">
								<h3 class="panel-title"><i class="fa fa-globe"></i> Informações Gerais do Município</h3>
							</div>
							<div class="panel-body">
								<form class="form form-horizontal">
									<div class="form-group">
										<label class="col-xs-2 col-sm-2 col-md-2 col-lg-2 control-label">Prefeito:</label>
										<div class="col-xs-4 col-sm-4 col-md-4 col-lg-4">
											<input type="text" class="form-control" readonly="readonly" value="<%=(rsDadosGerais.Fields.Item("nme_prefeito").Value)%>">
										</div>

										<label class="col-xs-2 col-sm-2 col-md-2 col-lg-2 control-label">Técnico Vistoriador:</label>
										<div class="col-xs-4 col-sm-4 col-md-4 col-lg-4">
											<input type="text" class="form-control" readonly="readonly" value="<%=(rsDadosGerais.Fields.Item("nme_vistoriador").Value)%>">
										</div>
									</div>
								</form>

								<form class="form form-horizontal">
									<div class="form-group">
										<label class="col-xs-2 col-sm-2 col-md-2 col-lg-2 control-label">Endereço:</label>
										<div class="col-xs-5 col-sm-5 col-md-5 col-lg-5">
											<input type="text" class="form-control" readonly="readonly" value="<%=(rsDadosGerais.Fields.Item("dsc_endereco").Value)%>">
										</div>

										<label class="col-xs-1 col-sm-1 col-md-1 col-lg-1 control-label">E-mail:</label>
										<div class="col-xs-4 col-sm-4 col-md-4 col-lg-4">
											<input type="text" class="form-control" readonly="readonly" value="<%=(rsDadosGerais.Fields.Item("end_email").Value)%>">
										</div>
									</div>
								</form>

								<form class="form form-horizontal">
									<div class="form-group">
										<label class="col-xs-2 col-sm-2 col-md-2 col-lg-2 control-label">População Atual (2010):</label>
										<div class="col-xs-4 col-sm-4 col-md-4 col-lg-4">
											<input type="text" class="form-control" readonly="readonly" value="<%=(rsDadosGerais.Fields.Item("qtd_populacao_2010").Value)%>">
										</div>

										<label class="col-xs-2 col-sm-2 col-md-2 col-lg-2 control-label">População Futura (2030):</label>
										<div class="col-xs-4 col-sm-4 col-md-4 col-lg-4">
											<input type="text" class="form-control" readonly="readonly" value="<%=(rsDadosGerais.Fields.Item("qtd_populacao_2030").Value)%>">
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
							<div class="panel-heading">
								<h3 class="panel-title"><i class="fa fa-fire"></i> Dados Sobre o Esgotamento Sanitário</h3>
							</div>
							<div class="panel-body">
								<form class="form form-horizontal">
									<div class="form-group">
										<div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
											<textarea readonly="readonly" class="form-control" rows="6"><%=(rsDadosGerais.Fields.Item("dsc_dados_basicos_esgotamento_sanitario").Value)%></textarea>
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
							<div class="panel-heading">
								<h3 class="panel-title"><i class="fa fa-map-marker"></i> ETE - Aspectos Administrativos e de Logística</h3>
							</div>
							<div class="panel-body">
								<form class="form form-horizontal">
									<div class="form-group">
										<label class="col-xs-2 col-sm-2 col-md-2 col-lg-2 control-label" style="padding-top: 0px;">Coordenadas UTM<br/>Chegada do Esgoto:</label>
										<div class="col-xs-3 col-sm-3 col-md-3 col-lg-3">
											<input type="text" class="form-control" readonly="readonly" value="<%=(rsDadosGerais.Fields.Item("num_latitude_chegada_esgoto").Value)%> - <%=(rsDadosGerais.Fields.Item("num_longitude_chegada_esgoto").Value)%>">
										</div>

										<label class="col-xs-4 col-sm-4 col-md-4 col-lg-4 control-label" style="padding-top: 0px;">Coordenadas UTM<br/>Lançamento do Esgoto no Corpo Receptor:</label>
										<div class="col-xs-3 col-sm-3 col-md-3 col-lg-3">
											<input type="text" class="form-control" readonly="readonly" value="<%=(rsDadosGerais.Fields.Item("num_latitude_lancamento_esgoto").Value)%> - <%=(rsDadosGerais.Fields.Item("num_longitude_lancamento_esgoto").Value)%>">
										</div>
									</div>
								</form>

								<form class="form form-horizontal">
									<div class="form-group">
										<div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
											<textarea readonly="readonly" class="form-control" rows="10"><%=(rsDadosGerais.Fields.Item("dsc_aspectos_administrativos_logistica").Value)%></textarea>
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
							<div class="panel-heading">
								<h3 class="panel-title"><i class="fa fa-arrow-right"></i> ETE - Croqui sem Escala - Indicação dos Dispositivos</h3>
							</div>
							<div class="panel-body">
								<form class="form form-horizontal">
									<div class="form-group">
										<label class="col-xs-2 col-sm-2 col-md-2 col-lg-2 control-label">Composição do<br/>Tratamento:</label>
										<div class="col-xs-10 col-sm-10 col-md-10 col-lg-10">
											<textarea readonly="readonly" class="form-control" rows="10"><%=(rsDadosGerais.Fields.Item("dsc_ete_dispositivos_composicao_tratamento").Value)%></textarea>
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
							<div class="panel-heading">
								<h3 class="panel-title"><i class="fa fa-thumbs-o-up"></i> ETE - Dispositivos - Conservação e Manuntenção</h3>
							</div>
							<div class="panel-body">
								<form class="form form-horizontal">
									<div class="form-group">
										<div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
											<textarea readonly="readonly" class="form-control" rows="10"><%=(rsDadosGerais.Fields.Item("dsc_ete_dispositivos_conservacao_manutencao").Value)%></textarea>
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
							<div class="panel-heading">
								<h3 class="panel-title"><i class="fa fa-list-alt"></i> ETE - Entorno - Descrição e Manutenção</h3>
							</div>
							<div class="panel-body">
								<form class="form form-horizontal">
									<div class="form-group">
										<div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
											<textarea readonly="readonly" class="form-control" rows="10"><%=(rsDadosGerais.Fields.Item("dsc_ete_entorno_descricao_manuntencao").Value)%></textarea>
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
							<div class="panel-heading">
								<h3 class="panel-title"><i class="fa fa-comments-o"></i> Comentários em Geral</h3>
							</div>
							<div class="panel-body">
								<form class="form form-horizontal">
									<div class="form-group">
										<div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
											<textarea readonly="readonly" class="form-control" rows="10"><%=(rsDadosGerais.Fields.Item("dsc_comentarios_gerais").Value)%></textarea>
										</div>
									</div>
								</form>
							</div>
						</div>
					</div>
				</div>

				<%
					strQueryLicencas = "SELECT * FROM tb_licenca_ambiental INNER JOIN tb_tipo_licenca ON tb_tipo_licenca.id = tb_licenca_ambiental.cod_tipo_licenca WHERE cod_empreendimento = '" & num_autos & "'"

					Set rs_licencas = Server.CreateObject("ADODB.Recordset")
						rs_licencas.CursorLocation = 3
						rs_licencas.CursorType = 3
						rs_licencas.LockType = 1
						rs_licencas.Open strQueryLicencas, objCon, , , &H0001

					strQueryOutorgas = "SELECT * FROM tb_outorga WHERE cod_empreendimento = '" & num_autos & "'"

					Set rs_outorgas = Server.CreateObject("ADODB.Recordset")
						rs_outorgas.CursorLocation = 3
						rs_outorgas.CursorType = 3
						rs_outorgas.LockType = 1
						rs_outorgas.Open strQueryOutorgas, objCon, , , &H0001

					strQueryApps = "SELECT * FROM tb_app WHERE cod_empreendimento = '" & num_autos & "'"

					Set rs_apps = Server.CreateObject("ADODB.Recordset")
						rs_apps.CursorLocation = 3
						rs_apps.CursorType = 3
						rs_apps.LockType = 1
						rs_apps.Open strQueryApps, objCon, , , &H0001

					strQueryTCRAs = "SELECT * FROM tb_tcra WHERE cod_empreendimento = '" & num_autos & "'"

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
											<th class="text-center" width="150">Data de Concessão</th>
											<th class="text-center" width="150">Data de Validade</th>
											<th class="text-center" width="50"></th>
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
												<td class="text-center text-middle">
													<%
														qtd_dias_vencimento = DateDiff("d", Now(), rs_licencas.Fields.Item("dta_vencimento").Value)

														If qtd_dias_vencimento > 0 And qtd_dias_vencimento <= 120 Then
													%>
													<span class="label label-warning"><i class="fa fa-warning"></i> <%=(qtd_dias_vencimento)%> dia(s) p/ Expirar</span>		
													<%
														Else
															If qtd_dias_vencimento < 0 Then
													%>
													<span class="label label-danger"><i class="fa fa-warning"></i> Documento Expirado</span>	
													<%
															End If
														End If
													%>
												</td>
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
											<th class="text-center" width="150">Data de Concessão</th>
											<th class="text-center" width="150">Data de Validade</th>
											<th class="text-center" width="50"></th>
										</thead>
										<tbody>
											<%
												While Not rs_outorgas.EOF
											%>
											<tr>
												<td><%=(rs_outorgas.Fields.Item("num_outorga").Value)%></td>
												<td class="text-center"><%=(rs_outorgas.Fields.Item("dta_concessao").Value)%></td>
												<td class="text-center"><%=(rs_outorgas.Fields.Item("dta_vencimento").Value)%></td>
												<td class="text-center text-middle">
													<%
														qtd_dias_vencimento = DateDiff("d", Now(), rs_outorgas.Fields.Item("dta_vencimento").Value)

														If qtd_dias_vencimento > 0 And qtd_dias_vencimento <= 120 Then
													%>
													<span class="label label-warning"><i class="fa fa-warning"></i> <%=(qtd_dias_vencimento)%> dia(s) p/ Expirar</span>		
													<%
														Else
															If qtd_dias_vencimento < 0 Then
													%>
													<span class="label label-danger"><i class="fa fa-warning"></i> Documento Expirado</span>	
													<%
															End If
														End If
													%>
												</td>
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
											<th class="text-center" width="150">Data de Concessão</th>
											<th class="text-center" width="150">Data de Validade</th>
											<th class="text-center" width="50"></th>
										</thead>
										<tbody>
											<%
												While Not rs_apps.EOF
											%>
											<tr>
												<td><%=(rs_apps.Fields.Item("num_app").Value)%></td>
												<td class="text-center"><%=(rs_apps.Fields.Item("dta_concessao").Value)%></td>
												<td class="text-center"><%=(rs_apps.Fields.Item("dta_vencimento").Value)%></td>
												<td class="text-center text-middle">
													<%
														qtd_dias_vencimento = DateDiff("d", Now(), rs_apps.Fields.Item("dta_vencimento").Value)

														If qtd_dias_vencimento > 0 And qtd_dias_vencimento <= 120 Then
													%>
													<span class="label label-warning"><i class="fa fa-warning"></i> <%=(qtd_dias_vencimento)%> dia(s) p/ Expirar</span>		
													<%
														Else
															If qtd_dias_vencimento < 0 Then
													%>
													<span class="label label-danger"><i class="fa fa-warning"></i> Documento Expirado</span>	
													<%
															End If
														End If
													%>
												</td>
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
											<th class="text-center" width="150">Data de Concessão</th>
											<th class="text-center" width="150">Data de Validade</th>
											<th class="text-center" width="50"></th>
										</thead>
										<tbody>
											<%
												While Not rs_tcras.EOF
											%>
											<tr>
												<td><%=(rs_tcras.Fields.Item("cod_tcra").Value)%></td>
												<td class="text-center"><%=(rs_tcras.Fields.Item("dta_concessao").Value)%></td>
												<td class="text-center"><%=(rs_tcras.Fields.Item("dta_vencimento").Value)%></td>
												<td class="text-center text-middle">
													<%
														qtd_dias_vencimento = DateDiff("d", Now(), rs_tcras.Fields.Item("dta_vencimento").Value)

														If qtd_dias_vencimento > 0 And qtd_dias_vencimento <= 120 Then
													%>
													<span class="label label-warning"><i class="fa fa-warning"></i> <%=(qtd_dias_vencimento)%> dia(s) p/ Expirar</span>	
													<%
														Else
															If qtd_dias_vencimento < 0 Then
													%>
													<span class="label label-danger"><i class="fa fa-warning"></i> Documento Expirado</span>	
													<%
															End If
														End If
													%>
												</td>
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
				%>

				<div class="row hidden-print">
					<div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
						<div class="panel panel-default">
							<div class="panel-heading">
								<h3 class="panel-title"><i class="fa fa-picture-o"></i> Galeria de Fotos</h3>
							</div>

							<div class="panel-body">
								<div class="row">
									<%
										While Not rsFotos.EOF
											pth_url = LCase(rsFotos.Fields.Item("pth_arquivo").Value)
											pth_url = Replace(pth_url, "\\10.0.75.125\intermultiplas.net\public\", "")
											pth_url = Replace(pth_url, "e:\home\programaagualimpa\web\", "")
											pth_url = Replace(pth_url, "\", "/")
											img_url = pth_url & "/" & rsFotos.Fields.Item("nme_arquivo").Value
									%>
									<div class="col-xs-4 col-sm-4 col-md-4 col-lg-4">
										<div class="thumbnail">
											<img src="<%=(img_url)%>" alt="<%=(rsFotos.Fields.Item("dsc_observacoes").Value)%>">
											<div class="caption">
												<!-- <h4>Thumbnail label</h4> -->
												<p class="thumbnail-label">
													<%=(rsFotos.Fields.Item("dsc_observacoes").Value)%>
												</p>
												<p class="hidden-print">
													<a href="<%=(img_url)%>" rel="group" role="button"
														title="<%=(rsFotos.Fields.Item("dsc_observacoes").Value)%>" 
														class="btn btn-default btn-block btn-sm fancybox">
														<i class="fa fa-expand"></i> Ampliar imagem
													</a>
												</p>
											</div>
										</div>
									</div>
									<%
											rsFotos.MoveNext
										Wend
									%>
								</div>
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

	<div class="modal fade" id="modalMapa" tabindex="-1" role="dialog" aria-labelledby="modalMapaLabel" aria-hidden="true">
		<div class="modal-dialog">
			<div class="modal-content">
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
					<h4 class="modal-title" id="modalMapaLabel"><i class="fa fa-map-marker"></i> Mapa de Localização</h4>
				</div>
				<div class="modal-body">
					<div id="map-canvas" style="width: 100%; height: 400px;"></div>
				</div>
			</div>
		</div>
	</div>
</body>
</html>