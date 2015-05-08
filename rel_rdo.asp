<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="logout.asp" -->
<%
	Response.CharSet = "UTF-8"

	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	Dim cod_empreendimento
	Dim dta_rdo

	cod_empreendimento = Request.QueryString("cod_empreendimento")
	dta_rdo = Request.QueryString("data")

	Dim rs_dados_obra
	Dim rs_dados_rdo
	Dim rs_dados_contrato

	If Not IsNull(cod_empreendimento) And Not IsEmpty(cod_empreendimento) And Not IsNull(dta_rdo) And Not IsEmpty(dta_rdo) Then
		Set rs_dados_obra = Server.CreateObject("ADODB.Recordset")
		rs_dados_obra.ActiveConnection = MM_cpf_STRING
		rs_dados_obra.Source = "SELECT * FROM c_lista_dados_obras WHERE PI = '" & cod_empreendimento & "'"
		rs_dados_obra.CursorType = 0
		rs_dados_obra.CursorLocation = 2
		rs_dados_obra.LockType = 1
		rs_dados_obra.Open()

		dta = Split(dta_rdo,"/")

		Set rs_dados_rdo = Server.CreateObject("ADODB.Recordset")
		rs_dados_rdo.ActiveConnection = MM_cpf_STRING
		rs_dados_rdo.Source = "SELECT * FROM c_lista_acompanhamento WHERE (((c_lista_acompanhamento.[PI])='"& cod_empreendimento &"') AND ((c_lista_acompanhamento.[Data do Registro])=#"& dta(1) & "/" & dta(0) & "/" & dta(2) &"#));"
		rs_dados_rdo.CursorType = 0
		rs_dados_rdo.CursorLocation = 2
		rs_dados_rdo.LockType = 1
		rs_dados_rdo.Open()

		Set rs_dados_contrato = Server.CreateObject("ADODB.Recordset")
		rs_dados_contrato.ActiveConnection = MM_cpf_STRING
		rs_dados_contrato.Source = "SELECT tb_contrato.* FROM tb_pi INNER JOIN (tb_contrato INNER JOIN tb_pi_contrato ON tb_contrato.id = tb_pi_contrato.cod_contrato) ON tb_pi.Código = tb_pi_contrato.cod_empreendimento WHERE (((tb_pi.PI)='"& cod_empreendimento &"'));"
		rs_dados_contrato.CursorType = 0
		rs_dados_contrato.CursorLocation = 2
		rs_dados_contrato.LockType = 1
		rs_dados_contrato.Open()
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
	<script type="text/javascript" src="js/fullscreen.js"></script>
	<script type="text/javascript" src="js/moment.min.js"></script>
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
	<script type="text/javascript" src="js/fancybox/jquery.fancybox.pack.js?v=2.1.5"></script>
	<script type="text/javascript">
		function calcDates() {
			// Dias Corridos
			var dta_assinatura = '<% If Not rs_dados_contrato.EOF Then Response.Write rs_dados_contrato.Fields.Item("dta_assinatura").Value End If %>';

			if (dta_assinatura != ""){
				dta_assinatura = moment(dta_assinatura, "DD/MM/YYYY");
				var dta_rdo = '<%=(dta_rdo)%>';
					dta_rdo = moment(dta_rdo, "DD/MM/YYYY");
				var dias_decorridos = dta_rdo.diff(dta_assinatura, 'days');
				
				// Dias Faltantes
				var dta_assinatura_ = '<%  If Not rs_dados_contrato.EOF Then Response.Write rs_dados_contrato.Fields.Item("dta_assinatura").Value End If %>';
					dta_assinatura_ = moment(dta_assinatura_, "DD/MM/YYYY");
				var prz_execucao = parseInt('<%  If Not rs_dados_contrato.EOF Then Response.Write rs_dados_contrato.Fields.Item("prz_original_execucao_meses").Value End If %>',10);
				var dta_encerramento = dta_assinatura_.add(prz_execucao, 'months');
				var dias_totais = dta_encerramento.diff(dta_assinatura, "days");
				var dias_faltantes = (dias_totais - dias_decorridos);
				
				// % de Andamento
				var perc_andamento = parseInt(((dias_decorridos / dias_totais) * 100), 10);

				$(".dias-decorridos").text(dias_decorridos);
				$(".dias-faltantes").text(dias_faltantes);
				$(".progress-bar").css("width", perc_andamento + "%");
				$(".progress").attr("title", perc_andamento + "%");

				if(perc_andamento > 0 && perc_andamento < 100)
					$(".progress-bar").addClass("progress-bar-warning");
				else if(perc_andamento == 100)
					$(".progress-bar").addClass("progress-bar-success");
				else if(perc_andamento > 100)
					$(".progress-bar").addClass("progress-bar-danger");

				$('[data-toggle="tooltip"]').tooltip();
			}
		}

		$(function(){
			$("li a.print").on("click", function(){
				window.print();
			});

			$(".fancybox").fancybox();

			var cod_empreendimento 	= '<%=(Request.QueryString("cod_empreendimento"))%>';
			var dta_rdo 			= '<%=(Request.QueryString("data"))%>';

			if(cod_empreendimento.length == 0 || dta_rdo.length == 0) {
				$("#modalFieldsFilter").modal("show");
			}
			else
				calcDates();
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
				<a class="navbar-brand" href="#">SIG - RDO</a>
			</div>

			<div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
				<ul class="nav navbar-nav navbar-right">
					<li><a href="javascript:window.history.back();"><i class="fa fa-chevron-left"></i> Voltar</a></li>
					<li><a href="#" class="print"><i class="fa fa-print"></i> Imprimir</a></li>
					<li><a href="<%= MM_Logout %>" class="sign-out"><i class="fa fa-sign-out"></i> Sair do Sistema</a></li>
				</ul>
			</div>
		</div>
	</nav>

	<div class="container">
		<div class="panel panel-default">
			<div class="panel-body">
				<div class="row">
					<div class="col-xs-2 text-center">
						<img src="LogoProjetoAguaLimpa.jpg" class="logo-intermultiplas report"/>
					</div>
					<div class="col-xs-8 text-center">
						<h3>
							<strong>RDO</strong>
							<br/>
							<small>Evolução da Obra</small>
						</h3>
					</div>
					<div class="col-xs-2"></div>
				</div>

				<hr/>

				<%
					If Not rs_dados_rdo.EOF Then
				%>

				<div class="row header-details">
					<div class="col-xs-12 text-left">
						<strong>Município: </strong> <%=(rs_dados_obra.Fields.Item("municipio").Value)%>
					</div>
				</div>

				<div class="row header-details">
					<div class="col-xs-6 text-left">
						<strong>Obra: </strong> <%=(rs_dados_obra.Fields.Item("nome_empreendimento").Value)%>-<%=(rs_dados_obra.Fields.Item("PI").Value)%>
					</div>

					<div class="col-xs-6 text-right">
						<strong><%=(rs_dados_rdo.Fields.Item("Data do Registro").Value)%></strong>
					</div>
				</div>

				<div class="row row-header header-details">
					<div class="col-xs-6 text-left">
						<strong>Relatório Nº:</strong> <%=(rs_dados_rdo.Fields.Item("cod_acompanhamento").Value)%>
					</div>

					<div class="col-xs-6 text-right">
						<strong>Criado por:</strong> <%=(rs_dados_rdo.Fields.Item("nme_interessado").Value)%>
					</div>
				</div>

				<div class="row activities">
					<div class="col-xs-12">
						<table class="table table-condensed table-boxed">
							<thead>
								<tr class="primary">
									<td class="text-center">Registro de Atividades</td>
								</tr>
							</thead>
							<tbody>
								<tr>
									<%
										If Not IsNull(rs_dados_rdo.Fields.Item("Registro").Value) And Not IsEmpty(rs_dados_rdo.Fields.Item("Registro").Value) Then
									%>
									<td><%=(rs_dados_rdo.Fields.Item("Registro").Value)%></td>
									<%
										Else
									%>
									<td class="text-center">N/A</td>
									<%
										End If
									%>
								</tr>
							</tbody>
						</table>
					</div>
				</div>

				<div class="row msst">
					<div class="col-xs-12">
						<table class="table table-condensed table-boxed">
							<thead>
								<tr class="primary">
									<td class="text-center">Segurança do Trabalho</td>
								</tr>
							</thead>
							<tbody>
								<tr>
									<%
										If Not IsNull(rs_dados_rdo.Fields.Item("dsc_situacao_sso").Value) And Not IsEmpty(rs_dados_rdo.Fields.Item("dsc_situacao_sso").Value) Then
									%>
									<td><%=(rs_dados_rdo.Fields.Item("dsc_situacao_sso").Value)%></td>
									<%
										Else
									%>
									<td class="text-center">N/A</td>
									<%
										End If
									%>
								</tr>
							</tbody>
						</table>
					</div>
				</div>

				<div class="row details">
					<div class="col-xs-4">
						<table class="table table-condensed table-boxed">
							<thead>
								<tr class="primary">
									<td colspan="3" class="text-center">Condições Meteorológicas</td>
								</tr>
							</thead>
							<tbody>
								<tr>
									<td><strong>Manhã</strong></td>
									<td class="text-center"><%=(rs_dados_rdo.Fields.Item("clima_manha").Value)%></td>
									<td class="text-center">
										<%
											If Not IsNull(rs_dados_rdo.Fields.Item("clima_manha").Value) And Not IsEmpty(rs_dados_rdo.Fields.Item("clima_manha").Value) Then
												If rs_dados_rdo.Fields.Item("clima_manha").Value = "Bom" Then
										%>
										<i class="fa fa-sun-o text-warning"></i>
										<%
												Else
													If rs_dados_rdo.Fields.Item("clima_manha").Value = "Chuva praticável" Then
										%>
										<i class="fa fa-tint text-info"></i>
										<%
													Else
														If rs_dados_rdo.Fields.Item("clima_manha").Value = "Chuva impraticável" Then
										%>
										<i class="fa fa-cloud text-danger"></i>
										<%
														End If
													End If
												End If
											End If
										%>
									</td>
								</tr>
								<tr>
									<td><strong>Tarde</strong></td>
									<td class="text-center"><%=(rs_dados_rdo.Fields.Item("clima_tarde").Value)%></td>
									<td class="text-center">
										<%
											If Not IsNull(rs_dados_rdo.Fields.Item("clima_tarde").Value) And Not IsEmpty(rs_dados_rdo.Fields.Item("clima_tarde").Value) Then
												If rs_dados_rdo.Fields.Item("clima_tarde").Value = "Bom" Then
										%>
										<i class="fa fa-sun-o text-warning"></i>
										<%
												Else
													If rs_dados_rdo.Fields.Item("clima_tarde").Value = "Chuva praticável" Then
										%>
										<i class="fa fa-tint text-info"></i>
										<%
													Else
														If rs_dados_rdo.Fields.Item("clima_tarde").Value = "Chuva impraticável" Then
										%>
										<i class="fa fa-cloud text-danger"></i>
										<%
														End If
													End If
												End If
											End If
										%>
									</td>
								</tr>
								<tr>
									<td><strong>Noite</strong></td>
									<td class="text-center"><%=(rs_dados_rdo.Fields.Item("clima_noite").Value)%></td>
									<td class="text-center">
										<%
											If Not IsNull(rs_dados_rdo.Fields.Item("clima_noite").Value) And Not IsEmpty(rs_dados_rdo.Fields.Item("clima_noite").Value) Then
												If rs_dados_rdo.Fields.Item("clima_noite").Value = "Bom" Then
										%>
										<i class="fa fa-sun-o text-warning"></i>
										<%
												Else
													If rs_dados_rdo.Fields.Item("clima_noite").Value = "Chuva praticável" Then
										%>
										<i class="fa fa-tint text-info"></i>
										<%
													Else
														If rs_dados_rdo.Fields.Item("clima_noite").Value = "Chuva impraticável" Then
										%>
										<i class="fa fa-cloud text-danger"></i>
										<%
														End If
													End If
												End If
											End If
										%>
									</td>
								</tr>
							</tbody>
						</table>
					</div>

					<div class="col-xs-3">
						<table class="table table-condensed table-boxed">
							<thead>
								<tr class="primary">
									<td colspan="2" class="text-center">Condições da Obra</td>
								</tr>
							</thead>
							<tbody>
								<tr>
									<td><strong>Limpeza</strong></td>
									<td class="text-center"><%=(rs_dados_rdo.Fields.Item("limpeza_obra").Value)%></td>
								</tr>
								<tr>
									<td><strong>Organização</strong></td>
									<td class="text-center"><%=(rs_dados_rdo.Fields.Item("organizacao_obra").Value)%></td>
								</tr>
							</tbody>
						</table>
					</div>

					<div class="col-xs-5">
						<table class="table table-condensed table-boxed">
							<thead>
								<tr class="primary">
									<td class="text-center">Evolução</td>
								</tr>
							</thead>
							<tbody>
								<tr>
									<td class="text-center text-90">
										Decorridos: <span class="dias-decorridos"></span> | Faltantes: <span class="dias-faltantes"></span>
									</td>
								</tr>
								<tr>
									<td>
										<div class="progress progress-striped active" data-toggle="tooltip" data-placement="bottom">
											<div class="progress-bar"></div>
										</div>
									</td>
								</tr>
							</tbody>
						</table>
					</div>
				</div>

				<div class="row histogram">
					<div class="col-xs-6">
						<table class="table table-condensed table-bordered">
							<thead>
								<tr class="primary">
									<td colspan="2" class="text-center">Mão de Obra</td>
								</tr>
							</thead>
							<tbody>
								<tr>
									<th>Nome do Recurso</th>
									<th class="text-right" width="100">Qtde</th>
								</tr>
								<%
									cod_acompanhamento = rs_dados_rdo.Fields.Item("cod_acompanhamento").Value
									strQ = "SELECT * FROM c_lista_histograma WHERE cod_acompanhamento = " & cod_acompanhamento & " AND cod_tipo_recurso = 1"

									Set rs_hist_mao_obra = Server.CreateObject("ADODB.Recordset")
										rs_hist_mao_obra.CursorLocation = 3
										rs_hist_mao_obra.CursorType = 3
										rs_hist_mao_obra.LockType = 1
										rs_hist_mao_obra.Open strQ, objCon, , , &H0001

									qtdRecursoCounter = 0

									If Not rs_hist_mao_obra.EOF Then
										While Not rs_hist_mao_obra.EOF
											qtdRecursoCounter = qtdRecursoCounter + rs_hist_mao_obra.Fields.Item("qtd_recurso").Value
								%>
								<tr>
									<td><%=(rs_hist_mao_obra.Fields.Item("nme_recurso").Value)%></td>
									<td class="text-right"><%=(rs_hist_mao_obra.Fields.Item("qtd_recurso").Value)%></td>
								</tr>
								<%
											rs_hist_mao_obra.MoveNext
										Wend
									End If
								%>
								<tr class="active">
									<th>TOTAL</th>
									<th class="text-right"><%=(qtdRecursoCounter)%></th>
								</tr>
							</tbody>
						</table>
					</div>

					<div class="col-xs-6">
						<table class="table table-condensed table-bordered">
							<thead>
								<tr class="primary">
									<td colspan="2" class="text-center">Equipamentos</td>
								</tr>
							</thead>
							<tbody>
								<tr>
									<th>Equipamento</th>
									<th class="text-right" width="100">Qtde</th>
								</tr>
								<%
									cod_acompanhamento = rs_dados_rdo.Fields.Item("cod_acompanhamento").Value
									strQ = "SELECT * FROM c_lista_histograma WHERE cod_acompanhamento = " & cod_acompanhamento & " AND cod_tipo_recurso = 2"

									Set rs_hist_equip = Server.CreateObject("ADODB.Recordset")
										rs_hist_equip.CursorLocation = 3
										rs_hist_equip.CursorType = 3
										rs_hist_equip.LockType = 1
										rs_hist_equip.Open strQ, objCon, , , &H0001

									qtdRecursoCounter = 0

									If Not rs_hist_equip.EOF Then
										While Not rs_hist_equip.EOF
											qtdRecursoCounter = qtdRecursoCounter + rs_hist_equip.Fields.Item("qtd_recurso").Value
								%>
								<tr>
									<td><%=(rs_hist_equip.Fields.Item("nme_recurso").Value)%></td>
									<td class="text-right"><%=(rs_hist_equip.Fields.Item("qtd_recurso").Value)%></td>
								</tr>
								<%
											rs_hist_equip.MoveNext
										Wend
									End If
								%>
								<tr class="active">
									<th>TOTAL</th>
									<th class="text-right"><%=(qtdRecursoCounter)%></th>
								</tr>
							</tbody>
						</table>
					</div>
				</div>

				<div class="row images-header">
					<div class="col-xs-12">
						<table class="table table-condensed table-boxed">
							<thead>
								<tr class="primary">
									<td class="text-center">Imagens</td>
								</tr>
							</thead>
						</table>
					</div>
				</div>

				<div class="row images-files">
					<%
						cod_acompanhamento = rs_dados_rdo.Fields.Item("cod_acompanhamento").Value
						strQ = "SELECT * FROM tb_acompanhamento_arquivo WHERE cod_referencia = " & cod_acompanhamento

						Set rs_imagens = Server.CreateObject("ADODB.Recordset")
							rs_imagens.CursorLocation = 3
							rs_imagens.CursorType = 3
							rs_imagens.LockType = 1
							rs_imagens.Open strQ, objCon, , , &H0001

						If Not rs_imagens.EOF Then
							While Not rs_imagens.EOF
								pth_url = rs_imagens.Fields.Item("pth_arquivo").Value
								pth_url = Replace(pth_url, "\\10.0.75.125\intermultiplas.net\public\", "")
								pth_url = Replace(pth_url, "e:\home\programaagualimpa\web\", "")
								pth_url = Replace(pth_url, "\", "/")
								img_url = pth_url & rs_imagens.Fields.Item("id_arquivo").Value &"_"& rs_imagens.Fields.Item("nme_arquivo").Value
					%>
					<div class="col-xs-4 col-sm-4 col-md-4">
						<div class="thumbnail">
							<img src="<%=(img_url)%>" alt="">
							<div class="caption">
								<!-- <h4>Thumbnail label</h4> -->
								<p><%=(rs_imagens.Fields.Item("dsc_observacoes").Value)%></p>
							</div>
						</div>
					</div>
					<%
								rs_imagens.MoveNext
							Wend
						Else
					%>
					<div class="col-xs-12 text-center">
						N/A
					</div>
					<%
						End If
					%>
				</div>

				<div class="row signature-lines">
					<div class="col-xs-6 text-center">____________________________________________</div>

					<div class="col-xs-6 text-center">____________________________________________</div>
				</div>

				<div class="row signature-names">
					<div class="col-xs-6 text-center">
						GERENCIADORA
					</div>

					<div class="col-xs-6 text-center">
						CONSTRUTORA
					</div>
				</div>

				<%
					Else
				%>

				<div class="row">
					<div class="col-xs-12 text-center">
						<div class="alert alert-warning"><i class="fa fa-warning"></i> Não há informações para a data selecionada!</div>
					</div>
				</div>

				<%
					End If
				%>
			</div>
		</div>

		<div class="modal fade" id="modalFieldsFilter" tabindex="-1" role="dialog" aria-labelledby="modalFieldsFilterLabel" aria-hidden="true">
			<div class="modal-dialog modal-sm">
				<div class="modal-content">
					<div class="modal-header">
						<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
						<h4 class="modal-title" id="modalFieldsFilterLabel">Selecione as informações que deseja visualizar</h4>
					</div>
					<div class="modal-body">
						<form id="form-filter" class="form" role="form">
							<div class="row">
								<div class="col-xs-12">
									<div class="form-group">
										<label class="control-label">Selecione a Obra</label>
										<select class="form-control">
											<option value=""></option>
										</select>
									</div>
								</div>
							</div>
						</form>
					</div>
				</div>
			</div>
		</div>
	</div>

</body>
</html>