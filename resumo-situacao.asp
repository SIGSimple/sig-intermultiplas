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
	<script type="text/javascript" src="js/jquery.floatThead.min.js"></script>
	<script type="text/javascript" src="js/jquery.table2excel.js"></script>
	<script type="text/javascript" src="js/moment.min.js"></script>
	<script type="text/javascript" src="js/fullscreen.js"></script>
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
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
			$("table.table-data").floatThead({
				scrollingTop: 60
			});

			adjustNumLayout();
			adjustVlrLayout();
			adjustPrcLayout();

			$("li a.print").on("click", function(){
				window.print();
			});

			$("li a.excel").on("click", function(){
				$("table.hide").table2excel({
					name: "Resumo da Situação por Município e Localidade"
				});
			});

			$(".btn-ficha-tecnica").on("click", function() {
				var codEmpreendimento = $(this).closest("tr").data().codEmpreendimento;

				var sql = "SELECT * FROM c_lista_dados_obras WHERE PI = " + codEmpreendimento;

				$.ajax({
					url: "query-to-json-util.asp",
					method: "POST",
					data: {
						sql: sql
					},
					beforeSend: function() {
						$("#modalLoading").modal("show");
					},
					success: function(data, textStatus, jqXHR){
						data = JSON.parse(data);

						if(data.length > 0) {
							var dadosObra = data[0];
							
							console.log(dadosObra);

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
		});
	</script>
	<style type="text/css">
		img.img-governo {
			max-width: 150px;
		}

		img.img-daee {
			max-height: 37px;
		}

		div.row-header {
			margin-bottom: 10px;
		}

		#modalFichaTecnica {
			overflow-y: scroll !important;
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
				<a class="navbar-brand" href="#">
					<%
						If Not IsNull(Request.QueryString("rep_universo_programa")) And Not IsEmpty(Request.QueryString("rep_universo_programa")) Then
							Response.Write "Programa Água Limpa | Universo do Programa"
						Else
							If Not IsNull(Request.QueryString("rep_universo_atendimento_programa")) And Not IsEmpty(Request.QueryString("rep_universo_atendimento_programa")) Then
								Response.Write "Programa Água Limpa | Universo de Atendimento do Programa"
							Else
								Response.Write "SIG - Resumo da Situação por Município e Localidade"
							End If
						End If
					%>
				</a>
			</div>

			<div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
				<ul class="nav navbar-nav navbar-right">
					<li><a href="javascript:window.history.back();"><i class="fa fa-chevron-left"></i> Voltar</a></li>
					<li><a href="#" class="print"><i class="fa fa-print"></i> Imprimir</a></li>
					<li><a href="#" class="excel"><i class="fa fa-file-excel-o"></i> Exportar p/ Excel</a></li>
					<li><a href="#" class="expand"><i class="fa fa-expand"></i>&nbsp;&nbsp;Tela Cheia</a></li>
					<li><a href="<%= MM_Logout %>" class="sign-out"><i class="fa fa-sign-out"></i> Sair do Sistema</a></li>
				</ul>
			</div>
		</div>
	</nav>

	<table class="hide">
		<tr class="info">
			<th class="text-middle text-center" rowspan="2" style="min-width: 180px;">Município</th>
			<th class="text-middle text-center" rowspan="2" style="min-width: 180px;">Localidade</th>
			<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 180px;">Bacia</th>
			<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 100px;">Status</th>
			<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 100px;">IBGE 2010</th>
			<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 100px;">POP 2030</th>
			<th class="text-middle text-center" rowspan="2" style="min-width: 150px;">Situação Atual</th>
			<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 150px;">Investimento Governo SP</th>
			<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" colspan="3">Em Atendimento</th>
			<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 120px;">Concluída<br/>Inaugurada</th>
			<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 120px;">Previsão para<br/>Inauguração</th>
			<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 170px;">Carga Orgânica Retirada<br/>(toneladas/mês)</th>
			<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 250px;">Notas / Observações</th>
		</tr>
		<tr class="info">
			<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="min-width: 120px;">Início das Obras</th>
			<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="min-width: 120px;">Executado</th>
			<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="min-width: 120px;">Previsão de Término</th>
		</tr>
		<%
			strQ = "SELECT * FROM c_lista_dados_obras"
			
			If(Request.QueryString("cod_municipio") <> "") Then
				strQ = strQ & " WHERE id_predio = "& Request.QueryString("cod_municipio")
			End If

			If Request.QueryString("rep_universo_atendimento_programa") <> "" Then
				strQ = strQ & " WHERE cod_status_situacao_interna in (1,2,3,4)"
			End If

			strQ = strQ & " ORDER BY Município, nome_empreendimento;"

			Set rs_lista = Server.CreateObject("ADODB.Recordset")
				rs_lista.CursorLocation = 3
				rs_lista.CursorType = 3
				rs_lista.LockType = 1
				rs_lista.Open strQ, objCon, , , &H0001

			If Not rs_lista.EOF Then
				While Not rs_lista.EOF
		%>
		<tr data-cod-municipio="<%=(rs_lista.Fields.Item("id_predio").Value)%>">
			<td class="text-middle" style="max-width: 180px;"><%=(rs_lista.Fields.Item("Município").Value)%></td>
			<td class="text-middle" style="max-width: 180px;"><%=(rs_lista.Fields.Item("nome_empreendimento").Value)%></td>
			<td class="text-middle text-center" style="max-width: 180px;"><%=( Mid(rs_lista.Fields.Item("bacia_daee").Value, 1, 3) )%></td>					
			<td class="text-middle text-center" style="max-width: 100px;"><%=(rs_lista.Fields.Item("cod_status_situacao_interna").Value)%></td>
			<td class="text-middle text-center num" style="max-width: 100px;"><%=(rs_lista.Fields.Item("qtd_populacao_urbana_2010").Value)%></td>
			<td class="text-middle text-center num" style="max-width: 100px;">
				<%
					If Not IsNull(rs_lista.Fields.Item("qtd_populacao_urbana_2010").Value) Then
						data = rs_lista.Fields.Item("qtd_populacao_urbana_2010").Value
						data = data * 1.25
						a = Round(data/100, 0)
						b = a * 100
						Response.Write b
					End If
				%>
			</td>
			<td class="text-middle text-center" style="max-width: 150px;"><%=(rs_lista.Fields.Item("desc_situacao_interna").Value)%></td>
			<td class="text-middle text-center vlr" style="max-width: 150px;"><%=(rs_lista.Fields.Item("Valor do Contrato").Value)%></td>
			<td class="text-middle text-center" style="max-width: 120px;">
				<%
					If Not IsNull(rs_lista.Fields.Item("mes_inicio_obras").Value) And Not IsEmpty(rs_lista.Fields.Item("mes_inicio_obras").Value) And rs_lista.Fields.Item("mes_inicio_obras").Value <> "" Then
						Response.Write UCase(MonthName(rs_lista.Fields.Item("mes_inicio_obras").Value,True)) & "/" & rs_lista.Fields.Item("ano_inicio_obras").Value
					End If
				%>
			</td>
			<td class="text-middle text-center prc" style="max-width: 120px;">
				<%
					If Not IsNull(rs_lista.Fields.Item("num_percentual_executado").Value) And Not IsEmpty(rs_lista.Fields.Item("num_percentual_executado").Value) Then
						num_percentual_executado = Replace(rs_lista.Fields.Item("num_percentual_executado").Value, ",", ".")

				%>
				<%=(num_percentual_executado)%>
				<%
					End If
				%>
			</td>
			<td class="text-middle text-center" style="max-width: 120px;">
				<%
					If Not IsNull(rs_lista.Fields.Item("mes_previsao_termino").Value) And Not IsEmpty(rs_lista.Fields.Item("mes_previsao_termino").Value) And rs_lista.Fields.Item("mes_previsao_termino").Value <> "" Then
						Response.Write UCase(MonthName(rs_lista.Fields.Item("mes_previsao_termino").Value,True)) & "/" & rs_lista.Fields.Item("ano_previsao_termino").Value
					End If
				%>
			</td>
			<td class="text-middle text-center" style="max-width: 120px;"><%=(rs_lista.Fields.Item("dta_inauguracao").Value)%></td>
			<td class="text-middle text-center" style="max-width: 120px;"><%=(rs_lista.Fields.Item("dta_previsao_inauguracao").Value)%></td>
			<td class="text-middle text-center" style="max-width: 170px;">
				<%
					If Not IsNull(rs_lista.Fields.Item("qtd_populacao_urbana_2010").Value) Then
						If Not IsNull(b) And (rs_lista.Fields.Item("cod_status_situacao_interna").Value) = 1 Then
							' Base de cálculo = qtd_populacao_urbana_2030 * 0,06 * 30 / 1000
							Response.Write b * 0.0018
						End If
					End If
				%>
			</td>
			<td class="text-middle" style="max-width: 180px;"><%=(rs_lista.Fields.Item("dsc_observacoes_relatorio_mensal").Value)%></td>
		</tr>
		<%
					rs_lista.MoveNext
				Wend
			End If
		%>
	</table>

	<%
		If Not IsNull(Request.QueryString("rep_universo_programa")) And Not IsEmpty(Request.QueryString("rep_universo_programa")) Then
	%>
		<div class="container container-box">
	<%
		End If
	%>
		<table class="table table-data table-bordered table-condensed table-hover table-striped table-responsive">
			<thead>
				<tr class="info">
					<%
						If Request.QueryString("rep_universo_programa") = "" Then
					%>
					<th width="50" rowspan="2"></th>
					<%
						End If
					%>
					<th class="text-middle text-center" rowspan="2" style="min-width: 180px;">Município</th>
					<th class="text-middle text-center" rowspan="2" style="min-width: 180px;">Localidade</th>
					<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 180px;">Bacia</th>
					<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 100px;">Status</th>
					<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 100px;">IBGE 2010</th>
					<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 100px;">POP 2030</th>
					<th class="text-middle text-center" rowspan="2" style="min-width: 150px;">Situação Atual</th>
					<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 150px;">Investimento Governo SP</th>
					<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" colspan="3">Em Atendimento</th>
					<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 120px;">Concluída<br/>Inaugurada</th>
					<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 120px;">Previsão para<br/>Inauguração</th>
					<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 170px;">Carga Orgânica Retirada<br/>(toneladas/mês)</th>
					<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" rowspan="2" style="min-width: 250px;">Notas / Observações</th>
					<%
						If Not IsNull(Request.QueryString("rep_universo_programa")) And Not IsEmpty(Request.QueryString("rep_universo_programa")) Then
					%>
					<th width="50" rowspan="2"></th>
					<%
						End If
					%>
				</tr>
				<tr class="info">
					<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="min-width: 120px;">Início das Obras</th>
					<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="min-width: 120px;">Executado</th>
					<th class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="min-width: 120px;">Previsão de Término</th>
				</tr>
			</thead>

			<tbody>
				<%
					strQ = "SELECT * FROM c_lista_dados_obras"
				
					If(Request.QueryString("cod_municipio") <> "") Then
						strQ = strQ & " WHERE id_predio = "& Request.QueryString("cod_municipio")
					End If

					If Request.QueryString("rep_universo_atendimento_programa") <> "" Then
						strQ = strQ & " WHERE cod_status_situacao_interna in (1,2,3,4)"
					End If

					strQ = strQ & " ORDER BY Município, nome_empreendimento;"

					Set rs_lista = Server.CreateObject("ADODB.Recordset")
						rs_lista.CursorLocation = 3
						rs_lista.CursorType = 3
						rs_lista.LockType = 1
						rs_lista.Open strQ, objCon, , , &H0001

					If Not rs_lista.EOF Then
						While Not rs_lista.EOF
				%>
				<tr data-cod-empreendimento="<%=(rs_lista.Fields.Item("PI").Value)%>">
					<%
						If Request.QueryString("rep_universo_programa") = "" Then
					%>
					<td class="text-middle hidden-print">
						<%
							link = ""
							If (rs_lista.Fields.Item("cod_status_situacao_interna").Value) = 1 Then
								link = "ficha-tecnica-obra-concluida.asp"
							Else
								If (rs_lista.Fields.Item("cod_status_situacao_interna").Value) = 2 Then
									link = "ficha-tecnica-obra.asp"
								Else
									If (rs_lista.Fields.Item("cod_status_situacao_interna").Value) = 3 Or (rs_lista.Fields.Item("cod_status_situacao_interna").Value) = 4 Then
										link = "ficha-tecnica-obra-programada.asp"
									Else
										If (rs_lista.Fields.Item("cod_status_situacao_interna").Value) = 5 Then
												link = "ficha-tecnica-obra-potencial.asp"
										End If
									End If
								End If
							End If

							If link <> "" Then
						%>
							<a href="<%=(link)%>?cod_empreendimento=<%=(rs_lista.Fields.Item("PI").Value)%>" class="btn btn-primary btn-sm">Ficha Técnica</a>
						<%
							End If
						%>
					</td>
					<%
						End If
					%>
					<td class="text-middle" style="max-width: 180px;"><%=(rs_lista.Fields.Item("Município").Value)%></td>
					<td class="text-middle" style="max-width: 180px;"><%=(rs_lista.Fields.Item("nome_empreendimento").Value)%></td>
					<td class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="max-width: 180px;"><%=( Mid(rs_lista.Fields.Item("bacia_daee").Value, 1, 3) )%></td>					
					<td class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="max-width: 100px;"><%=(rs_lista.Fields.Item("cod_status_situacao_interna").Value)%></td>
					<td class="text-middle text-center num <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="max-width: 100px;"><%=(rs_lista.Fields.Item("qtd_populacao_urbana_2010").Value)%></td>
					<td class="text-middle text-center num <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="max-width: 100px;">
						<%
							If Not IsNull(rs_lista.Fields.Item("qtd_populacao_urbana_2010").Value) Then
								data = rs_lista.Fields.Item("qtd_populacao_urbana_2010").Value
								data = data * 1.25
								a = Round(data/100, 0)
								b = a * 100
								Response.Write b
							End If
						%>
					</td>
					<td class="text-middle text-center" style="max-width: 150px;"><%=(rs_lista.Fields.Item("desc_situacao_interna").Value)%></td>
					<td class="text-middle text-center vlr <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="max-width: 150px;"><%=(rs_lista.Fields.Item("Valor do Contrato").Value)%></td>
					<td class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="max-width: 120px;">
						<%
							If Not IsNull(rs_lista.Fields.Item("mes_inicio_obras").Value) And Not IsEmpty(rs_lista.Fields.Item("mes_inicio_obras").Value) And rs_lista.Fields.Item("mes_inicio_obras").Value <> "" Then
								Response.Write UCase(MonthName(rs_lista.Fields.Item("mes_inicio_obras").Value,True)) & "/" & rs_lista.Fields.Item("ano_inicio_obras").Value
							End If
						%>
					</td>
					<td class="text-middle text-center prc <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="max-width: 120px;">
						<%
							If Not IsNull(rs_lista.Fields.Item("num_percentual_executado").Value) And Not IsEmpty(rs_lista.Fields.Item("num_percentual_executado").Value) Then
							num_percentual_executado = Replace(rs_lista.Fields.Item("num_percentual_executado").Value, ",", ".")

						%>
						<%=(num_percentual_executado)%>
						<%
							End If
						%>
					</td>
					<td class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="max-width: 120px;">
						<%
							If Not IsNull(rs_lista.Fields.Item("mes_previsao_termino").Value) And Not IsEmpty(rs_lista.Fields.Item("mes_previsao_termino").Value) And rs_lista.Fields.Item("mes_previsao_termino").Value <> "" Then
								Response.Write UCase(MonthName(rs_lista.Fields.Item("mes_previsao_termino").Value,True)) & "/" & rs_lista.Fields.Item("ano_previsao_termino").Value
							End If
						%>
					</td>
					<td class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="max-width: 120px;"><%=(rs_lista.Fields.Item("dta_inauguracao").Value)%></td>
					<td class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="max-width: 120px;"><%=(rs_lista.Fields.Item("dta_previsao_inauguracao").Value)%></td>
					<td class="text-middle text-center <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="max-width: 170px;">
						<%
							If Not IsNull(rs_lista.Fields.Item("qtd_populacao_urbana_2010").Value) Then
								If Not IsNull(b) And (rs_lista.Fields.Item("cod_status_situacao_interna").Value) = 1 Then
									' Base de cálculo = qtd_populacao_urbana_2030 * 0,06 * 30 / 1000
									Response.Write b * 0.0018
								End If
							End If
						%>
					</td>
					<td class="text-middle <% If Request.QueryString("rep_universo_programa") <> "" Then Response.Write "hide" End If %>" style="max-width: 3000px;"><%=(rs_lista.Fields.Item("dsc_observacoes_relatorio_mensal").Value)%></td>
					<%
						If Not IsNull(Request.QueryString("rep_universo_programa")) And Not IsEmpty(Request.QueryString("rep_universo_programa")) Then
					%>
					<td class="text-middle hidden-print">
						<%
							link = ""
							If (rs_lista.Fields.Item("cod_status_situacao_interna").Value) = 1 Then
								link = "ficha-tecnica-obra-concluida.asp"
							Else
								If (rs_lista.Fields.Item("cod_status_situacao_interna").Value) = 2 Then
									link = "ficha-tecnica-obra.asp"
								Else
									If (rs_lista.Fields.Item("cod_status_situacao_interna").Value) = 3 Or (rs_lista.Fields.Item("cod_status_situacao_interna").Value) = 4 Then
										link = "ficha-tecnica-obra-programada.asp"
									Else
										If (rs_lista.Fields.Item("cod_status_situacao_interna").Value) = 5 Then
												link = "ficha-tecnica-obra-potencial.asp"
										End If
									End If
								End If
							End If

							If link <> "" Then
						%>
							<a href="<%=(link)%>?cod_empreendimento=<%=(rs_lista.Fields.Item("PI").Value)%>" class="btn btn-primary btn-sm">Ficha Técnica</a>
						<%
							End If
						%>
					</td>
					<%
						End If
					%>
				</tr>
				<%
							rs_lista.MoveNext
						Wend
					End If
				%>
			</tbody>
		</table>
	<%
		If Request.QueryString("rep_universo_programa") = "" Then
	%>
		</div>
	<%
		End If
	%>

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

	<div class="modal fade" id="modalFichaTecnica" tabindex="-1" role="dialog" aria-labelledby="modalLoadingLabel" aria-hidden="true">
		<div class="modal-dialog">
			<div class="modal-content">
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
					<h4 class="modal-title" id="modalFichaTecnicaLabel"><strong>Programa Água Limpa</strong> | Ficha Técnica da Obra</h4>
				</div>
				<div class="modal-body">
					<div class="row row-header">
						<div class="col-xs-3">
							<img src="img/governo_estado_500.png" class="img-responsive img-governo">
						</div>
						<div class="col-xs-9">
							<img src="logo_daee.jpg" class="img-responsive img-daee">
						</div>
					</div>

					<div class="row row-header">
						<div class="col-xs-12">
							<small><strong>Governo do Estado de São Paulo</strong></small>
							<br/>
							<small>Secretaria de Saneamento e Recursos Hídricos</small>
							<br/>
							<small>Departamento de Águas e Energia Elétrica</small>
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
				</div>
				<div class="modal-footer">
					<button type="button" class="btn btn-default" data-dismiss="modal">Fechar</button>
				</div>
			</div>
		</div>
	</div>
</body>
</html>