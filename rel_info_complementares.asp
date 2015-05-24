<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="logout.asp" -->
<%
	Response.CharSet = "UTF-8"

	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	sql = "SELECT * FROM c_lista_rel_info_complementares ORDER BY nme_municipio ASC, nome_empreendimento ASC"

	Dim rs_lista

	Set rs_lista = Server.CreateObject("ADODB.Recordset")
	rs_lista.ActiveConnection = MM_cpf_STRING
	rs_lista.Source = sql
	rs_lista.CursorType = 0
	rs_lista.CursorLocation = 2
	rs_lista.LockType = 1
	rs_lista.Open()
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
	<script type="text/javascript" src="js/jquery.table2excel.js"></script>
	<script type="text/javascript" src="js/moment.min.js"></script>
	<script type="text/javascript" src="js/fullscreen.js"></script>
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
	<script type="text/javascript" src="js/fancybox/jquery.fancybox.pack.js?v=2.1.5"></script>
	<script type="text/javascript">
		$(function(){
			$("li a.print").on("click", function(){
				window.print();
			});

			$("li a.excel").on("click", function(){
				$("table").table2excel({
					name: "Informações Complementares das Obras"
				});
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
				<a class="navbar-brand" href="#">SIG - Informações Complementares das Obras</a>
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
						<h3><strong>Informações Complementares das Obras</strong><br/><small>Listagem do Banco de Dados</small></h3>
					</div>
					<div class="col-xs-2"></div>
				</div>

				<hr/>

				<div class="row">
					<div class="col-xs-12">
						<table class="table table-bordered table-condensed table-hover table-striped">
							<thead>
								<th class="text-center text-middle">Município</th>
								<th class="text-center text-middle">Localidade</th>
								<th class="text-center text-middle">Situação</th>
								<th class="text-center text-middle">Nº Autos</th>
								<th class="text-center text-middle">Objeto da Obra</th>
								<th class="text-center text-middle">Bacia Hidrográfica</th>
								<th class="text-center text-middle">Manancial de Lançamento</th>
								<th class="text-center text-middle">Latitude/Longitude</th>
								<th class="text-center text-middle">Coletor Tronco (metros)</th>
								<th class="text-center text-middle">Interceptor (metros)</th>
								<th class="text-center text-middle">Emissário fluente Bruto (metros)</th>
								<th class="text-center text-middle">EEE (qtd)</th>
								<th class="text-center text-middle">Linha de Recalque (metros)</th>
								<th class="text-center text-middle">Tipo ETE</th>
								<th class="text-center text-middle">Estação de Tratamento (desc.)</th>
								<th class="text-center text-middle">Emissário Efluente Tratado (metros)</th>
								<th class="text-center text-middle">Estudo Elab. DAEE</th>
								<th class="text-center text-middle">Observações</th>
								<th class="text-center text-middle">Benefício Geral da Obra</th>
								<th class="text-center text-middle">Parceria/Realização</th>
							</thead>
							<tbody>
								<%
									While (NOT rs_lista.EOF)
								%>
								<tr>
									<td><%=(rs_lista.Fields.Item("nme_municipio").Value)%></td>
									<td><%=(rs_lista.Fields.Item("nome_empreendimento").Value)%></td>
									<td><%=(rs_lista.Fields.Item("desc_situacao_externa").Value)%></td>
									<td><%=(rs_lista.Fields.Item("PI").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dsc_objeto_obra").Value)%></td>
									<td><%=(rs_lista.Fields.Item("nme_bacia_hidrografica").Value)%></td>
									<td><%=(rs_lista.Fields.Item("nme_manancial").Value)%></td>
									<td><%=(rs_lista.Fields.Item("latitude_longitude").Value)%></td>
									<td><%=(rs_lista.Fields.Item("qtd_metragem_coletor_tronco").Value)%></td>
									<td><%=(rs_lista.Fields.Item("qtd_metragem_interceptor").Value)%></td>
									<td><%=(rs_lista.Fields.Item("qtd_metragem_emissario_fluente_bruto").Value)%></td>
									<td><%=(rs_lista.Fields.Item("qtd_eee").Value)%></td>
									<td><%=(rs_lista.Fields.Item("qtd_metragem_linha_recalque").Value)%></td>
									<td><%=(rs_lista.Fields.Item("nme_tipo_ete").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dsc_estacao_tratamento").Value)%></td>
									<td><%=(rs_lista.Fields.Item("qtd_metragem_emissario_efluente_tratado").Value)%></td>
									<td>
										<%
											If rs_lista.Fields.Item("flg_estudo_elaborado_daee").Value Then
												Response.Write "Sim"
											Else
												Response.Write "Não"
											End If
										%>
									</td>
									<td><%=(rs_lista.Fields.Item("dsc_observacoes_obra").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dsc_resultado_obtido").Value)%></td>
									<td><%=(rs_lista.Fields.Item("nme_parceria_realizacao").Value)%></td>
								</tr>
								<%
										rs_lista.MoveNext()
									Wend
								%>
							</tbody>
						</table>
					</div>
				</div>
			</div>
		</div>
	</div>

</body>
</html>