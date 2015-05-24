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
	<script type="text/javascript" src="//code.jquery.com/jquery-1.11.2.min.js"></script>
	<script type="text/javascript" src="js/jquery.number.min.js"></script>
	<script type="text/javascript" src="js/underscore-min.js"></script>
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
	<script type="text/javascript" src="http://maps.googleapis.com/maps/api/js?libraries=places&sensor=false"></script>
	<script type="text/javascript" src="http://code.highcharts.com/highcharts.js"></script>
	<script type="text/javascript" src="http://code.highcharts.com/modules/exporting.js"></script>
	<script type="text/javascript" src="js/moment.min.js"></script>
	<script type="text/javascript" src="js/fullscreen.js"></script>
	<script type="text/javascript" src="js/loadMapaObra.js"></script>
	<script type="text/javascript" src="js/common.js"></script>
	<script type="text/javascript">
		function getLatitudeLongitude() {
			return "<%=(rs_dados_obra.Fields.Item("latitude_longitude").Value)%>";
		}

		$(function(){
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

						$("#txt-municipio-localidade").text(dadosObra['Município'] +" - "+ dadosObra['nome_empreendimento']);
						$("#txt-nome-prefeitura").text(dadosObra['prefeitura']);
						$("#txt-nome-prefeito").text(dadosObra['prefeito']);
						$("#txt-nome-bacia-daee").text(dadosObra['bacia_daee']);
						
						if(dadosObra['Descrição da Intervenção FDE'])
							$("#txt-objeto-obra").text((dadosObra['Descrição da Intervenção FDE']) ? dadosObra['Descrição da Intervenção FDE'] : "");
						else
							$(".tr-objeto-obra").hide();
						
						if(dadosObra['qtd_populacao_urbana_2010'])
							$("#txt-pop-2010").text(dadosObra['qtd_populacao_urbana_2010']);
						else
							$(".tr-pop-2010").hide();

						if(pop2030)
							$("#txt-pop-2030").text(pop2030);
						else
							$(".tr-pop-2030").hide();
						
						$("#txt-nota-obra").text((dadosObra['dsc_observacoes_relatorio_mensal']) ? dadosObra['dsc_observacoes_relatorio_mensal'] : "");
						$("#txt-observacoes-gerais").text((dadosObra['observacoes_gestor']) ? dadosObra['observacoes_gestor'] : "");

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
					<div class="col-xs-3">
						<img src="img/governo_estado_500.png" class="img-responsive img-governo">
					</div>
					
					<div class="col-xs-7 text-center">
						<strong>Governo do Estado de São Paulo</strong>
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
								<tr class="tr-pop-2010">
									<td class="text-middle text-bold">
										População Beneficiada em 2010
									</td>
									<td class="text-middle text-right num" id="txt-pop-2010"></td>
								</tr>
								<tr class="tr-pop-2030">
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
										Última Informação
									</td>
									<td id="txt-nota-obra"></td>
								</tr>
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
					If (Session("MM_UserAuthorization") <> 8 AND Session("MM_UserAuthorization") <> 9) Then
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
					End If

					strQueryPendLicenciamento = "SELECT * FROM c_lista_rel_pendencias WHERE flg_pendencia_valida = ""Sim"" AND flg_pendencia_concluida = ""Não"" AND cod_tipo_pendencia = 6 AND PI = '"& cod_empreendimento &"'"

					Set rs_pend_lic = Server.CreateObject("ADODB.Recordset")
						rs_pend_lic.CursorLocation = 3
						rs_pend_lic.CursorType = 3
						rs_pend_lic.LockType = 1
						rs_pend_lic.Open strQueryPendLicenciamento, objCon, , , &H0001

					strQueryPendDocFundiaria = "SELECT * FROM c_lista_rel_pendencias WHERE flg_pendencia_valida = ""Sim"" AND flg_pendencia_concluida = ""Não"" AND cod_tipo_pendencia = 7 AND PI = '"& cod_empreendimento &"'"

					Set rs_pend_doc_fund = Server.CreateObject("ADODB.Recordset")
						rs_pend_doc_fund.CursorLocation = 3
						rs_pend_doc_fund.CursorType = 3
						rs_pend_doc_fund.LockType = 1
						rs_pend_doc_fund.Open strQueryPendDocFundiaria, objCon, , , &H0001

					strQueryPendProjExecutivo = "SELECT * FROM c_lista_rel_pendencias WHERE flg_pendencia_valida = ""Sim"" AND flg_pendencia_concluida = ""Não"" AND cod_tipo_pendencia = 1 AND PI = '"& cod_empreendimento &"'"

					Set rs_pend_proj_exec = Server.CreateObject("ADODB.Recordset")
						rs_pend_proj_exec.CursorLocation = 3
						rs_pend_proj_exec.CursorType = 3
						rs_pend_proj_exec.LockType = 1
						rs_pend_proj_exec.Open strQueryPendProjExecutivo, objCon, , , &H0001

					strQueryPendProgRecFinanceiros = "SELECT * FROM c_lista_rel_pendencias WHERE flg_pendencia_valida = ""Sim"" AND flg_pendencia_concluida = ""Não"" AND cod_tipo_pendencia = 8 AND PI = '"& cod_empreendimento &"'"

					Set rs_pend_prog_rec_fin = Server.CreateObject("ADODB.Recordset")
						rs_pend_prog_rec_fin.CursorLocation = 3
						rs_pend_prog_rec_fin.CursorType = 3
						rs_pend_prog_rec_fin.LockType = 1
						rs_pend_prog_rec_fin.Open strQueryPendProgRecFinanceiros, objCon, , , &H0001
				%>

				<div class="row">
					<div class="col-xs-12">
						<div class="panel panel-default">
							<div class="panel-heading">
								<h3 class="panel-title">
									<i class="fa fa-warning"></i> Pendências
								</h3>
							</div>

							<div class="panel-body">
								<label class="control-label">Pendências de Licenciamento</label>
								<table class="table table-history table-bordered table-hover table-striped table-condensed">
									<tbody>
										<%
											If Not rs_pend_lic.EOF Then
												While Not rs_pend_lic.EOF
										%>
										<tr>
											<td><%=(rs_pend_lic.Fields.Item("dsc_pendencia").Value)%></td>
										</tr>
										<%
													rs_pend_lic.MoveNext
												Wend
											End If
										%>
									</tbody>
								</table>

								<label class="control-label">Pendências de Documentação Fundiária</label>
								<table class="table table-history table-bordered table-hover table-striped table-condensed">
									<tbody>
										<%
											If Not rs_pend_doc_fund.EOF Then
												While Not rs_pend_doc_fund.EOF
										%>
										<tr>
											<td><%=(rs_pend_doc_fund.Fields.Item("dsc_pendencia").Value)%></td>
										</tr>
										<%
													rs_pend_doc_fund.MoveNext
												Wend
											End If
										%>
									</tbody>
								</table>

								<label class="control-label">Pendências de Projetos Executivos</label>
								<table class="table table-history table-bordered table-hover table-striped table-condensed">
									<tbody>
										<%
											If Not rs_pend_proj_exec.EOF Then
												While Not rs_pend_proj_exec.EOF
										%>
										<tr>
											<td><%=(rs_pend_proj_exec.Fields.Item("dsc_pendencia").Value)%></td>
										</tr>
										<%
													rs_pend_proj_exec.MoveNext
												Wend
											End If
										%>
									</tbody>
								</table>

								<label class="control-label">Pendências de Programação de Recursos Financeiros</label>
								<table class="table table-history table-bordered table-hover table-striped table-condensed">
									<tbody>
										<%
											If Not rs_pend_prog_rec_fin.EOF Then
												While Not rs_pend_prog_rec_fin.EOF
										%>
										<tr>
											<td><%=(rs_pend_prog_rec_fin.Fields.Item("dsc_pendencia").Value)%></td>
										</tr>
										<%
													rs_pend_prog_rec_fin.MoveNext
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