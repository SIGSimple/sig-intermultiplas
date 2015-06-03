<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="logout.asp" -->
<%
	Response.CharSet = "UTF-8"

	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
		objCon.Open MM_cpf_STRING

	Dim sql

	sql = ""
	sql = sql & "SELECT "
	sql = sql & 	"tb_info_emp_concluidos.id,"
	sql = sql & 	"tb_info_emp_concluidos.num_autos,"
	sql = sql & 	"tb_info_emp_concluidos.nme_municipio,"
	sql = sql & 	"tb_info_emp_concluidos.nme_localidade,"

	sql = sql & 	"tb_info_emp_concluidos.dsc_necessita_reparos,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_problemas_bombas,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_falta_limpeza,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_despejo_irregular_residuos,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_falta_funcionario_operacao,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_danos_cercamento,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_danos_tratamento_preliminar,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_danos_talude,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_danos_lagoa,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_problemas_diversos,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_danos_emissarios,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_danos_caixa_passagem_interligacoes,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_danos_drenagem,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_partes_inoperantes,"
	sql = sql & 	"tb_info_emp_concluidos.dsc_situacao_operacao,"

	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_necessita_reparos]),0,1) 					AS flg_necessita_reparos,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_problemas_bombas]),0,1) 					AS flg_problemas_bombas,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_falta_limpeza]),0,1) 						AS flg_falta_limpeza,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_despejo_irregular_residuos]),0,1) 			AS flg_despejo_irregular_residuos,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_falta_funcionario_operacao]),0,1) 			AS flg_falta_funcionario_operacao,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_danos_cercamento]),0,1) 					AS flg_danos_cercamento,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_danos_tratamento_preliminar]),0,1) 			AS flg_danos_tratamento_preliminar,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_danos_talude]),0,1) 						AS flg_danos_talude,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_danos_lagoa]),0,1) 							AS flg_danos_lagoa,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_problemas_diversos]),0,1) 					AS flg_problemas_diversos,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_danos_emissarios]),0,1) 					AS flg_danos_emissarios,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_danos_caixa_passagem_interligacoes]),0,1) 	AS flg_danos_caixa_passagem_interligacoes,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_danos_drenagem]),0,1) 						AS flg_danos_drenagem,"
	sql = sql & 	"IIf(IsNull([tb_info_emp_concluidos.dsc_partes_inoperantes]),0,1) 					AS flg_partes_inoperantes "
	sql = sql & "FROM tb_info_emp_concluidos "

	If Request.QueryString("dsc_situacao_operacao") <> "" Then
		sql = sql & "WHERE dsc_situacao_operacao = '"& Request.QueryString("dsc_situacao_operacao") &"'"
	End If

	Set rs_lista_matriz = Server.CreateObject("ADODB.Recordset")
		rs_lista_matriz.CursorLocation = 3
		rs_lista_matriz.CursorType = 3
		rs_lista_matriz.LockType = 1
		rs_lista_matriz.PageSize = 10
		rs_lista_matriz.Open sql, objCon, , , &H0001

	pg = 0
	rec = 0

	If Not rs_lista_matriz.EOF Then
		If Request.QueryString("pg") = "" Then
			pg = 1
		Else
			If CInt(Request.QueryString("pg")) < 1 Then
				pg = 1
			Else
				pg = Request.QueryString("pg")
			End If
		End If
		
		rs_lista_matriz.AbsolutePage = pg
	End If
%>
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
	<script type="text/javascript" src="//code.jquery.com/jquery-1.11.2.min.js"></script>
	<script type="text/javascript" src="http://code.highcharts.com/highcharts.js"></script>
	<script type="text/javascript" src="http://code.highcharts.com/modules/exporting.js"></script>
	<script type="text/javascript" src="js/jquery.number.min.js"></script>
	<script type="text/javascript" src="js/jquery.table2excel.js"></script>
	<script type="text/javascript" src="js/jquery.lazyload.min.js"></script>
	<script type="text/javascript" src="js/moment.min.js"></script>
	<script type="text/javascript" src="js/fullscreen.js"></script>
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
	<script type="text/javascript" src="js/fancybox/jquery.fancybox.pack.js?v=2.1.5"></script>
	<script type="text/javascript" src="js/jquery.floatThead.min.js"></script>
	
	<script type="text/javascript">
		$(function(){
			$('[data-toggle="tooltip"]').tooltip();
			$('[data-toggle="popover"]').popover();

			var colors = ["#E74C3C", "#F39C12", "#18BC9C", "#3498DB", "#F9FF00"];

			$('#chart-situacao-operacao').highcharts({
				colors: colors,
				visible: false,
				title: {
					text: ''
				},
				chart: {
					plotBackgroundColor: null,
					plotBorderWidth: null,
					plotShadow: false
				},
				tooltip: {
					pointFormat: '{series.name}: <b>{point.y}</b>'
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
					name: 'Qtd. Obras',
					data: [
						<%
							sql = "SELECT Count(id) AS qtd_obras, dsc_situacao_operacao FROM tb_info_emp_concluidos GROUP BY dsc_situacao_operacao;"

							Set rs = Server.CreateObject("ADODB.Recordset")
								rs.CursorLocation = 3
								rs.CursorType = 3
								rs.LockType = 1
								rs.PageSize = 10
								rs.Open sql, objCon, , , &H0001

							While (Not rs.EOF)
								qtd_obras 				= rs.Fields.Item("qtd_obras").Value
								dsc_situacao_operacao 	= rs.Fields.Item("dsc_situacao_operacao").Value
						%>
						["<%=(dsc_situacao_operacao)%>", <%=(qtd_obras)%>],
						<%
								rs.MoveNext()
							Wend
						%>
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
		<div class="panel panel-default">
			<div class="panel-heading">
				<h3 class="panel-title">Situação de Operação das Obras Concluídas</h3>
			</div>
			<div class="panel-body">
				<div class="row">
					<div class="col-xs-12 col-lg-6">
						<div id="chart-situacao-operacao" style="min-width: 310px; height: 300px; max-width: 600px; margin: 0 auto"></div>
					</div>
					<div class="col-xs-12 col-lg-6">
						<table class="table table-bordered table-condensed table-striped table-hover">
							<thead>
								<tr class="active">
									<th><a class="btn btn-xs btn-default" href="<%=(Request.ServerVariables("URL"))%>"><i class="fa fa-times-circle"></i></a></th>
									<th>Situação</th>
									<th class="text-center" width="100">Qtd. Obras</th>
								</tr>
							</thead>
							<tbody>
								<%
									sql = "SELECT Count(id) AS qtd_obras, dsc_situacao_operacao FROM tb_info_emp_concluidos GROUP BY dsc_situacao_operacao;"

									Set rs = Server.CreateObject("ADODB.Recordset")
										rs.CursorLocation = 3
										rs.CursorType = 3
										rs.LockType = 1
										rs.PageSize = 10
										rs.Open sql, objCon, , , &H0001

									qtd_total_obras_concluidas = 0
									While (Not rs.EOF)
										qtd_obras 				= rs.Fields.Item("qtd_obras").Value
										dsc_situacao_operacao 	= rs.Fields.Item("dsc_situacao_operacao").Value
										qtd_total_obras_concluidas = qtd_total_obras_concluidas + qtd_obras
								%>
								<tr>
									<td width="20"><a class="btn btn-xs btn-primary" href="?dsc_situacao_operacao=<%=(dsc_situacao_operacao)%>#table"><i class="fa fa-search"></i></td>
									<td><%=(dsc_situacao_operacao)%></td>
									<td class="text-center"><%=(qtd_obras)%></td>
								</tr>
								<%
										rs.MoveNext()
									Wend
								%>
								<tr class="active">
									<th colspan="2">Total Obras Concluídas</th>
									<th class="text-center"><%=(qtd_total_obras_concluidas)%></th>
								</tr>
							</tbody>
						</table>
					</div>
				</div>
			</div>
		</div>

		<div class="panel panel-default">
			<div class="panel-heading">
				<h3 class="panel-title">Matriz de Controle de Manutenção</h3>
			</div>
			<div class="panel-body">
				<table id="table" class="table table-bordered table-condensed table-striped table-hover">
					<thead>
						<th></th>
						<th>Município</th>
						<th>Localidade</th>
						<th class="text-center">Situação</th>
						<th class="text-center">
							<i class="fa fa-wrench" data-toggle="tooltip" data-placement="top" title="Necessita de Reparos"></i>
						</th>
						<th class="text-center">
							<i class="fa fa-bomb" data-toggle="tooltip" data-placement="top" title="Problema nas Bombas"></i>
						</th>
						<th class="text-center">
							<i class="fa fa-eraser" data-toggle="tooltip" data-placement="top" title="Falta Limpeza"></i>
						</th>
						<th class="text-center">
							<i class="fa fa-trash-o" data-toggle="tooltip" data-placement="top" title="Despejo Irregular de Resíduos"></i>
						</th>
						<th class="text-center">
							<i class="fa fa-user" data-toggle="tooltip" data-placement="top" title="Falta Funcionário p/ Operação Diária"></i>
						</th>
						<th class="text-center">
							<i class="fa fa-table" data-toggle="tooltip" data-placement="top" title="Danos no Cercamento"></i>
						</th>
						<th class="text-center">
							<i class="fa fa-filter" data-toggle="tooltip" data-placement="top" title="Danos no Tratamento Preliminar"></i>
						</th>
						<th class="text-center">
							<i class="fa fa-area-chart" data-toggle="tooltip" data-placement="top" title="Danos no Talude"></i>
						</th>
						<th class="text-center">
							<i class="fa fa-tint" data-toggle="tooltip" data-placement="top" title="Danos nas Lagoas"></i>
						</th>
						<th class="text-center">
							<i class="fa fa-warning" data-toggle="tooltip" data-placement="top" title="Problemas Diversos"></i>
						</th>
						<th class="text-center">
							<i class="fa fa-download" data-toggle="tooltip" data-placement="top" title="Danos nos Emissários"></i>
						</th>
						<th class="text-center">
							<i class="fa fa-cube" data-toggle="tooltip" data-placement="top" title="Danos nas Caixas de Passagem/Interligações"></i>
						</th>
						<th class="text-center">
							<i class="fa fa-code-fork" data-toggle="tooltip" data-placement="top" title="Danos de Drenagem"></i>
						</th>
					</thead>
					<tbody>
						<%
							While ((rec < rs_lista_matriz.PageSize) And (Not rs_lista_matriz.EOF))
						%>
						<tr>
							<th class="row-header text-center">
								<a class="btn btn-xs btn-primary" data-toggle="tooltip" data-placement="right" title="Visualizar Ficha de Vistoria"
									href="rel_obras_concluidas.asp?id=<%=(rs_lista_matriz.Fields.Item("id").Value)%>">
									<i class="fa fa-file-text-o"></i>
								</a>
							</th>
							<th class="row-header">
								<%=(rs_lista_matriz.Fields.Item("nme_municipio").Value)%>
							</th>
							<th class="row-header">
								<%=(rs_lista_matriz.Fields.Item("nme_localidade").Value)%>
							</th>
							<th class="row-header text-center">
								<%
									labelColor = ""
									Select Case rs_lista_matriz.Fields.Item("dsc_situacao_operacao").Value
										Case "Operando"
											labelColor = "success"
										Case "Operando em teste"
											labelColor = "info"
										Case "Operando Parcialmente"
											labelColor = "yellow"
										Case "Estado de Abandono"
											labelColor = "danger"
										Case "Inoperante"
											labelColor = "warning"
									End Select
								%>

								<span class="label label-<%=(labelColor)%>"><%=(rs_lista_matriz.Fields.Item("dsc_situacao_operacao").Value)%></span>
							</th>
							<td class="text-center" width="50">
								<%
									If rs_lista_matriz.Fields.Item("flg_necessita_reparos").Value = 1 Then
								%>
								<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
									data-placement="bottom" data-trigger="hover" title="Necessita de Reparos"
									data-content="<%=(rs_lista_matriz.Fields.Item("dsc_necessita_reparos").Value)%>">
									<i class="fa fa-wrench"></i>
								</a>
								<%
									End If
								%>
							</td>
							<td class="text-center" width="50">
								<%
									If rs_lista_matriz.Fields.Item("flg_problemas_bombas").Value = 1 Then
								%>
								<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
									data-placement="bottom" data-trigger="hover" title="Problema nas Bombas"
									data-content="<%=(rs_lista_matriz.Fields.Item("dsc_problemas_bombas").Value)%>">
									<i class="fa fa-bomb"></i>
								</a>
								<%
									End If
								%>
							</td>
							<td class="text-center" width="50">
								<%
									If rs_lista_matriz.Fields.Item("flg_falta_limpeza").Value = 1 Then
								%>
								<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
									data-placement="bottom" data-trigger="hover" title="Falta Limpeza"
									data-content="<%=(rs_lista_matriz.Fields.Item("dsc_falta_limpeza").Value)%>">
									<i class="fa fa-eraser"></i>
								</a>
								<%
									End If
								%>
							</td>
							<td class="text-center" width="50">
								<%
									If rs_lista_matriz.Fields.Item("flg_despejo_irregular_residuos").Value = 1 Then
								%>
								<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
									data-placement="bottom" data-trigger="hover" title="Despejo Irregular de Resíduos"
									data-content="<%=(rs_lista_matriz.Fields.Item("dsc_despejo_irregular_residuos").Value)%>">
									<i class="fa fa-trash-o"></i>
								</a>
								<%
									End If
								%>
							</td>
							<td class="text-center" width="50">
								<%
									If rs_lista_matriz.Fields.Item("flg_falta_funcionario_operacao").Value = 1 Then
								%>
								<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
									data-placement="bottom" data-trigger="hover" title="Falta Funcionário p/ Operação Diária"
									data-content="<%=(rs_lista_matriz.Fields.Item("dsc_falta_funcionario_operacao").Value)%>">
									<i class="fa fa-user"></i>
								</a>
								<%
									End If
								%>
							</td>
							<td class="text-center" width="50">
								<%
									If rs_lista_matriz.Fields.Item("flg_danos_cercamento").Value = 1 Then
								%>
								<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
									data-placement="bottom" data-trigger="hover" title="Danos no Cercamento"
									data-content="<%=(rs_lista_matriz.Fields.Item("dsc_danos_cercamento").Value)%>">
									<i class="fa fa-table"></i>
								</a>
								<%
									End If
								%>
							</td>
							<td class="text-center" width="50">
								<%
									If rs_lista_matriz.Fields.Item("flg_danos_tratamento_preliminar").Value = 1 Then
								%>
								<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
									data-placement="bottom" data-trigger="hover" title="Danos no Tratamento Preliminar"
									data-content="<%=(rs_lista_matriz.Fields.Item("dsc_danos_tratamento_preliminar").Value)%>">
									<i class="fa fa-filter"></i>
								</a>
								<%
									End If
								%>
							</td>
							<td class="text-center" width="50">
								<%
									If rs_lista_matriz.Fields.Item("flg_danos_talude").Value = 1 Then
								%>
								<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
									data-placement="bottom" data-trigger="hover" title="Danos no Talude"
									data-content="<%=(rs_lista_matriz.Fields.Item("dsc_danos_talude").Value)%>">
									<i class="fa fa-area-chart"></i>
								</a>
								<%
									End If
								%>
							</td>
							<td class="text-center" width="50">
								<%
									If rs_lista_matriz.Fields.Item("flg_danos_lagoa").Value = 1 Then
								%>
								<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
									data-placement="bottom" data-trigger="hover" title="Danos nas Lagoas"
									data-content="<%=(rs_lista_matriz.Fields.Item("dsc_danos_lagoa").Value)%>">
									<i class="fa fa-tint"></i>
								</a>
								<%
									End If
								%>
							</td>
							<td class="text-center" width="50">
								<%
									If rs_lista_matriz.Fields.Item("flg_problemas_diversos").Value = 1 Then
								%>
								<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
									data-placement="bottom" data-trigger="hover" title="Problemas Diversos"
									data-content="<%=(rs_lista_matriz.Fields.Item("dsc_problemas_diversos").Value)%>">
									<i class="fa fa-warning"></i>
								</a>
								<%
									End If
								%>
							</td>
							<td class="text-center" width="50">
								<%
									If rs_lista_matriz.Fields.Item("flg_danos_emissarios").Value = 1 Then
								%>
								<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
									data-placement="bottom" data-trigger="hover" title="Danos nos Emissários"
									data-content="<%=(rs_lista_matriz.Fields.Item("dsc_danos_emissarios").Value)%>">
									<i class="fa fa-download"></i>
								</a>
								<%
									End If
								%>
							</td>
							<td class="text-center" width="50">
								<%
									If rs_lista_matriz.Fields.Item("flg_danos_caixa_passagem_interligacoes").Value = 1 Then
								%>
								<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
									data-placement="bottom" data-trigger="hover" title="Danos nas Caixas de Passagem/Interligações"
									data-content="<%=(rs_lista_matriz.Fields.Item("dsc_danos_caixa_passagem_interligacoes").Value)%>">
									<i class="fa fa-cube"></i>
								</a>
								<%
									End If
								%>
							</td>
							<td class="text-center" width="50">
								<%
									If rs_lista_matriz.Fields.Item("flg_danos_drenagem").Value = 1 Then
								%>
								<a tabindex="0" role="button" class="label label-primary" data-toggle="popover" 
									data-placement="bottom" data-trigger="hover" title="Danos de Drenagem"
									data-content="<%=(rs_lista_matriz.Fields.Item("dsc_danos_drenagem").Value)%>">
									<i class="fa fa-code-fork"></i>
								</a>
								<%
									End If
								%>
							</td>
						</tr>
						<%
								rs_lista_matriz.MoveNext()
								rec = (rec+1)
							Wend
						%>
					</tbody>
				</table>
			</div>
			<div class="panel-footer clearfix">
				<div class="pull-right">
					<ul class="pagination pagination-sm">
						<li class="<%IF CInt(pg) = 1 Then Response.Write "disabled" End If%>"><a href="?pg=1#table"><<</a></li>
						<li class="<%IF CInt(pg) = 1 Then Response.Write "disabled" End If%>"><a href="?pg=<%=(CInt(pg)-1)%>#table"><</a></li>
						<%
							i = 1
							While i <= rs_lista_matriz.PageCount
								If CInt(pg) = i Then
						%>
						<li class="active"><a href="?pg=<%=(i)%>#table"><%=(i)%></a></li>
						<%
								Else
						%>
						<li><a href="?pg=<%=(i)%>#table"><%=(i)%></a></li>
						<%
								End If
								i = (i+1)
							Wend
						%>
						<li class="<%If CInt(pg) = rs_lista_matriz.PageCount Then Response.Write "disabled" End If%>"><a href="?pg=<%=(CInt(pg)+1)%>#table">></a></li>
						<li class="<%If CInt(pg) = rs_lista_matriz.PageCount Then Response.Write "disabled" End If%>"><a href="?pg=<%=(rs_lista_matriz.PageCount)%>#table">>></a></li>
					</ul>
				</div>
			</div>
		</div>
	</div>
</body>
</html>