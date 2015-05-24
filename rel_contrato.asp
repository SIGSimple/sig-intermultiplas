<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="logout.asp" -->
<%
	Response.CharSet = "UTF-8"

	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	sql = "SELECT * FROM c_lista_rel_contrato"

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
					name: "Listagem de Contratos"
				});
			});

			var vlr_lines = $(".vlr");
			$.each(vlr_lines, function(i, item){
				$(item).val("R$ " + $.number($(item).val(), 2, ",", "."));
				$(item).text("R$ " + $.number($(item).text(), 2, ",", "."));
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
				<a class="navbar-brand" href="#">SIG - Listagem de Contratos</a>
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
						<h3><strong>Listagem de Contratos</strong><br/><small>Listagem do Banco de Dados</small></h3>
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
								<th class="text-center text-middle">Situação Atual</th>
								<th class="text-center text-middle">Nº Autos Licitação</th>
								<th class="text-center text-middle">Nº Edital</th>
								<th class="text-center text-middle">Nº Autos Convênio</th>
								<th class="text-center text-middle">Nº Convênio</th>
								<th class="text-center text-middle">Empresa Contratada</th>
								<th class="text-center text-middle">Eng. Empresa Contr.</th>
								<th class="text-center text-middle">Nº Autos Contrato</th>
								<th class="text-center text-middle">Nº Contrato</th>
								<th class="text-center text-middle">Dt. Assinatura</th>
								<th class="text-center text-middle">Dt. Publ. D.O.E.</th>
								<th class="text-center text-middle">Dt. Pedido Empenho</th>
								<th class="text-center text-middle">Dt. Base</th>
								<th class="text-center text-middle">Dt. Inauguração</th>
								<th class="text-center text-middle">Dt. Termo Rec. Provisório</th>
								<th class="text-center text-middle">Dt. Termo Rec. Definitivo</th>
								<th class="text-center text-middle">Dt. Enc. Contrato</th>
								<th class="text-center text-middle">Dt. Rec. Contratual</th>
								<th class="text-center text-middle">Dt. O.S.</th>
								<th class="text-center text-middle">Vigência Até (Digitado)</th>
								<th class="text-center text-middle">Vigência Até (Calculado)</th>
								
								<th class="text-center text-middle">Prazo Original Execução Serviço</th>
								<th class="text-center text-middle">Aditivos (Prazo)</th>
								<th class="text-center text-middle">Prazo Total Serviço</th>

								<th class="text-center text-middle">Prazo Original Contrato</th>
								<th class="text-center text-middle">Aditivos (Prazo)</th>
								<th class="text-center text-middle">Prazo Total Contrato</th>
								
								<th class="text-center text-middle" style="min-width: 150px;">Valor Original</th>
								<th class="text-center text-middle" style="min-width: 150px;">Aditivos (Valor)</th>
								<th class="text-center text-middle" style="min-width: 150px;">Valor Total</th>
							</thead>
							<tbody>
								<%
									While (NOT rs_lista.EOF)
										prz_total_execucao 	= rs_lista.Fields.Item("prz_original_execucao_meses").Value + rs_lista.Fields.Item("prz_aditivo").Value
										prz_total_contrato 	= rs_lista.Fields.Item("prz_original_contrato_meses").Value + rs_lista.Fields.Item("prz_aditivo").Value
										vlr_total 			= rs_lista.Fields.Item("vlr_original").Value + rs_lista.Fields.Item("vlr_aditivo").Value
										dta_os 				= rs_lista.Fields.Item("dta_os").Value
								%>
								<tr>
									<td><%=(rs_lista.Fields.Item("nme_municipio").Value)%></td>
									<td><%=(rs_lista.Fields.Item("nome_empreendimento").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dsc_situacao_externa").Value)%></td>
									<td><%=(rs_lista.Fields.Item("num_autos_licitacao").Value)%></td>
									<td><%=(rs_lista.Fields.Item("num_edital").Value)%></td>
									<td><%=(rs_lista.Fields.Item("num_autos_convenio").Value)%></td>
									<td><%=(rs_lista.Fields.Item("num_convenio").Value)%></td>
									<td><%=(rs_lista.Fields.Item("empresa_contratada").Value)%></td>
									<td><%=(rs_lista.Fields.Item("engenheiro_empresa_contratada").Value)%></td>
									<td><%=(rs_lista.Fields.Item("num_autos").Value)%></td>
									<td><%=(rs_lista.Fields.Item("num_contrato").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_assinatura").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_publicacao_doe").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_pedido_empenho").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_base").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_inauguracao").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_termo_recebimento_provisorio").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_termo_recebimento_definitivo").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_encerramento_contrato").Value)%></td>
									<td><%=(rs_lista.Fields.Item("dta_recisao_contratual").Value)%></td>
									<td><%=(dta_os)%></td>
									<td></td>
									<td>
										<%
											If IsNull(rs_lista.Fields.Item("dta_vigencia").Value) Or IsEmpty(rs_lista.Fields.Item("dta_vigencia").Value) Then
												If dta_os <> "" Then
													If prz_total_contrato > 0 Then
														dta_vigencia_contrato = DateAdd("m", prz_total_contrato, dta_os)
														Response.Write dta_vigencia_contrato
													End If
												End If
											Else
												Response.Write rs_lista.Fields.Item("dta_vigencia").Value
											End If
										%>
									</td>
									
									<td><%=(rs_lista.Fields.Item("prz_original_execucao_meses").Value)%></td>
									<td><%=(rs_lista.Fields.Item("prz_aditivo").Value)%></td>
									<td><%=(prz_total_execucao)%></td>

									<td><%=(rs_lista.Fields.Item("prz_original_contrato_meses").Value)%></td>
									<td><%=(rs_lista.Fields.Item("prz_aditivo").Value)%></td>
									<td><%=(prz_total_contrato)%></td>
									<td class="text-center vlr"><%=(rs_lista.Fields.Item("vlr_original").Value)%></td>
									<td class="text-center vlr"><%=(rs_lista.Fields.Item("vlr_aditivo").Value)%></td>
									<td class="text-center vlr"><%=(vlr_total)%></td>
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