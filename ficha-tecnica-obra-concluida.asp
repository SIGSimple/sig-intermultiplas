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

	strQ = "SELECT * FROM c_lista_dados_obras WHERE PI = '"& cod_empreendimento &"'"

	Set rs_dados_obra = Server.CreateObject("ADODB.Recordset")
		rs_dados_obra.CursorLocation = 3
		rs_dados_obra.CursorType = 3
		rs_dados_obra.LockType = 1
		rs_dados_obra.Open strQ, objCon, , , &H0001

	If cod_municipio = "" Then
		cod_municipio = rs_dados_obra.Fields.Item("cod_mun").Value
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
	<script type="text/javascript" src="http://maps.googleapis.com/maps/api/js?libraries=places&sensor=false"></script>
	<script type="text/javascript" src="http://code.highcharts.com/highcharts.js"></script>
	<script type="text/javascript" src="http://code.highcharts.com/modules/exporting.js"></script>
	<script type="text/javascript" src="js/fancybox/jquery.fancybox.pack.js?v=2.1.5"></script>
	<script type="text/javascript" src="js/moment.min.js"></script>
	<script type="text/javascript" src="js/fullscreen.js"></script>
	<script type="text/javascript" src="js/loadMapaObra.js"></script>
	<script type="text/javascript" src="js/common.js"></script>
	<script type="text/javascript">
		function getLatitudeLongitude() {
			return "<%=(rs_dados_obra.Fields.Item("latitude_longitude").Value)%>";
		}

		$(function(){
			$(".fancybox").fancybox();

			$('#modalMapa').on('show.bs.modal', function() {
				resizeMap();
			});

			$("li a.map").on("click", function(){
				$("#modalMapa").modal("show");
			});

			$("li a.print").on("click", function(){
				window.print();
			});

			var cod_empreendimento = '<%=(Request.QueryString("cod_empreendimento"))%>'

			$.ajax({
				url: "query-to-json-util.asp",
				method: "POST",
				data: {
					sql: "SELECT * FROM (tb_pi_contrato INNER JOIN c_lista_dados_obras ON tb_pi_contrato.cod_empreendimento = c_lista_dados_obras.Código) INNER JOIN c_lista_contrato ON tb_pi_contrato.cod_contrato = c_lista_contrato.id WHERE PI = '" + cod_empreendimento + "'"
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

						moment.locale("pt-br");
						var dta_os 			= (dadosObra['dta_os']) ? moment(dadosObra['dta_os'], "DD/MM/YYYY").format("MMMM/YYYY") : "";
						var dta_inauguracao = (dadosObra['dta_inauguracao']) ? moment(dadosObra['dta_inauguracao'], "DD/MM/YYYY").format("MMMM/YYYY") : "";

						$("#txt-municipio-localidade").text(dadosObra['Município'] +" - "+ dadosObra['nome_empreendimento']);
						$("#txt-nome-prefeitura").text(dadosObra['prefeitura']);
						$("#txt-nome-prefeito").text(dadosObra['prefeito']);
						$("#txt-nome-bacia-daee").text(dadosObra['bacia_daee']);
						
						if(dadosObra['nme_bacia_hidrografica'])
							$("#txt-nome-bacia-hidrografica").text((dadosObra['nme_bacia_hidrografica']) ? dadosObra['nme_bacia_hidrografica'] : "");
						else
							$(".tr-nome-bacia-hidrografica").hide();

						if(dadosObra['nme_manancial'])
							$("#txt-nome-manancial-lancamento").text((dadosObra['nme_manancial'])?dadosObra['nme_manancial']:"");
						else
							$(".tr-nome-manancial-lancamento").hide();

						if(dadosObra['Descrição da Intervenção FDE'])
							$("#txt-objeto-obra").text((dadosObra['Descrição da Intervenção FDE']) ? dadosObra['Descrição da Intervenção FDE'] : "");
						else
							$(".tr-objeto-obra").hide();

						$("#txt-investimento-governo").text(dadosObra['Valor do Contrato']);
						$("#txt-pop-2010").text(dadosObra['qtd_populacao_urbana_2010']);
						$("#txt-pop-2030").text(pop2030);
						
						if(dadosObra['dta_os'])
							$("#txt-dta-os").text(dta_os);
						else
							$(".tr-dta-os").hide();
						
						if(dadosObra['dta_inauguracao'])
							$("#txt-dta-inauguracao").text(dta_inauguracao);
						else
							$(".tr-dta-inauguracao").hide();
						
						if(dadosObra['dsc_resultado_obtido'])
							$("#txt-beneficio-obra").text((dadosObra['dsc_resultado_obtido']) ? dadosObra['dsc_resultado_obtido'] : "");
						else
							$(".tr-beneficio-obra").hide();
						
						$("#txt-beneficio-ambiental").text("Carga Orgânica Retirada: "+ cargaOrganizaRetirada +" (toneladas/mês)");

						if(dadosObra['dsc_parceria_realizacao'])
							$("#txt-parceria-realizacao").text((dadosObra['dsc_parceria_realizacao']) ? dadosObra['dsc_parceria_realizacao'] : "");
						else
							$(".tr-parceria-realizacao").hide();

						$("#txt-ultimas-informacoes").text((dadosObra['dsc_observacoes_relatorio_mensal']) ? dadosObra['dsc_observacoes_relatorio_mensal'] : "");
						$("#txt-observacoes-gerais").text((dadosObra['observacoes_gestor']) ? dadosObra['observacoes_gestor'] : "");

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
				<a class="navbar-brand" href="#">SIG - Ficha Técnica da Obra</a>
			</div>

			<div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
				<ul class="nav navbar-nav navbar-right">
					<li><a href="javascript:window.history.back();"><i class="fa fa-chevron-left"></i> Voltar</a></li>
					<li><a href="informacao-municipio-resumida.asp?cod_municipio=<%=(cod_municipio)%>"><i class="fa fa-list-alt"></i> Inf. Município</a></li>
					<%
						If rs_dados_obra.Fields.Item("latitude_longitude").Value <> "" Then
					%>
					
					<li><a href="#" class="map"><i class="fa fa-map-marker"></i> Mapa</a></li>
					
					<%
						End If
					
						If (CInt(Session("MM_UserAuthorization")) <> 8 And CInt(Session("MM_UserAuthorization")) <> 9) Then	
					%>
					
					<li><a href="#" class="print"><i class="fa fa-print"></i> Imprimir</a></li>

					<%
						End If
					%>
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
					<div class="col-xs-12">
						<table class="table table-condensed">
							<tbody>
								<tr class="info">
									<td class="text-bold text-title text-center">
										<%=(rs_dados_obra.Fields.Item("municipio").Value)%> - <%=(rs_dados_obra.Fields.Item("nome_empreendimento").Value)%>
									</td>
								</tr>
								<tr>
									<td class="text-bold text-center">
										<%
											If Session("MM_UserAuthorization") = 8 Or Session("MM_UserAuthorization") = 9 Then
												Response.Write rs_dados_obra.Fields.Item("desc_situacao_externa").Value
											Else
												Response.Write rs_dados_obra.Fields.Item("desc_situacao_interna").Value
											End If
										%>
									</td>
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
								<tr class="tr-nome-bacia-hidrografica">
									<td class="text-middle text-bold" width="200">Bacia Hidrográfica</td>
									<td class="text-middle" id="txt-nome-bacia-hidrografica"></td>
								</tr>
								<tr class="tr-nome-manancial-lancamento">
									<td class="text-middle text-bold" width="200">Manancial de Lançamento</td>
									<td class="text-middle" id="txt-nome-manancial-lancamento"></td>
								</tr>
							</tbody>
						</table>

						<table class="table table-bordered table-condensed">
							<tbody>
								<tr class="tr-objeto-obra">
									<td class="text-middle text-bold" width="200">
										Objeto da Obra
									</td>
									<td id="txt-objeto-obra"></td>
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
								<tr class="tr-dta-os">
									<td class="text-middle text-bold">
										Ordem de Serviço
									</td>
									<td class="text-middle text-right" id="txt-dta-os"></td>
								</tr>
								<tr class="tr-dta-inauguracao">
									<td class="text-middle text-bold">
										Conclusão/Inauguração
									</td>
									<td class="text-middle text-right" id="txt-dta-inauguracao"></td>
								</tr>
							</tbody>
						</table>

						<table class="table table-bordered table-condensed">
							<tbody>
								<tr class="tr-beneficio-obra">
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
								<tr class="tr-parceria-realizacao">
									<td class="text-middle text-bold" width="150">
										Parceria/Realização
									</td>
									<td id="txt-parceria-realizacao"></td>
								</tr>
							</tbody>
						</table>

						<table class="table table-bordered table-condensed">
							<tbody>
								<%
									If ((Session("MM_UserAuthorization") = 8 Or Session("MM_UserAuthorization") = 9)) Then
								%>
								<tr>
									<td class="text-middle text-bold" width="150">
										Últimas Informações
									</td>
									<td id="txt-ultimas-informacoes"></td>
								</tr>
								<%
									End If
								%>
								<tr>
									<td class="text-middle text-bold" width="150">
										Observações Gerais
									</td>
									<td id="txt-observacoes-gerais"></td>
								</tr>
							</tbody>
						</table>
					</div>
				</div>

				<%
					strQueryFotos = "SELECT * FROM c_lista_todas_fotos_obra WHERE PI = '" & cod_empreendimento & "' AND report = True"

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
												<p class="thumbnail-label"><%=(rs_fotos.Fields.Item("dsc_observacoes").Value)%></p>
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

					If ((Session("MM_UserAuthorization") <> 8 AND Session("MM_UserAuthorization") <> 9)) Then
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