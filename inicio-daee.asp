<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="logout.asp" -->
<!--#include file="daee_restrict_access.asp" -->
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
	<script type="text/javascript" src="js/fullscreen.js"></script>
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
	<script type="text/javascript">
		$(function() {
			function clearForm() {
				$("#cod_municipio").val("");
				$("#cod_municipio").trigger("change");

				$("#cod_empreendimento").val("");
				$("#cod_empreendimento").trigger("change");

				$("#general_info").removeAttr("checked")
				$("#situation_overview").removeAttr("checked")
			}

			$("#btnGenerateReport").on("click", function(){
				$(this).button("loading");
				var report_page = $("#modelo_relatorio").val();
				var cod_mun = $("#cod_municipio").val();
				var cod_emp = $("#cod_empreendimento").val();

				if(cod_mun != "")
					report_page += "?cod_municipio=" + cod_mun

				if(cod_emp != "")
					report_page += "&cod_empreendimento=" + cod_emp

				window.location.href = report_page
			});

			$("#btnClearForm").on("click", function(){
				clearForm();
			});

			$("#cod_municipio").on("change", function() {
				var cod_modelo_relatorio = $("#modelo_relatorio").val();
				var cod_mun = $("#cod_municipio").val();
				var url = "inicio-daee.asp";

				if(cod_mun != ""){
					$("#btnClearForm").removeAttr("disabled");
					$("#btnGenerateReport").removeAttr("disabled");
				}

				if((cod_mun != "") && !$("#cod_empreendimento").closest("div.row").hasClass("hide")){
					url += "?cod_modelo_relatorio="+ cod_modelo_relatorio +"&cod_mun=" + cod_mun;

					if(cod_modelo_relatorio == "informacao-obra-andamento.asp")
						url += "&filtra_situacao=sim";
				}
				
				if(!$("#cod_empreendimento").closest("div.row").hasClass("hide"))
					location.replace(url);
			});

			$("#cod_empreendimento").on("change", function() {
				var cod_empreendimento = $("#cod_empreendimento").val();

				if(cod_empreendimento != ""){
					$("#btnClearForm").removeAttr("disabled");
					$("#btnGenerateReport").removeAttr("disabled");
				}
			});

			$("#modelo_relatorio").on("change", function(e) {
				optionData = $("#modelo_relatorio option:selected").data();

				if(optionData.habilitaBotoes) {
					$("#btnClearForm").removeAttr("disabled");
					$("#btnGenerateReport").removeAttr("disabled");
				}
				else {
					if(!$("#btnClearForm").attr("disabled"))
						$("#btnClearForm").attr("disabled","disabled");
					
					if(!$("#btnGenerateReport").attr("disabled"))
						$("#btnGenerateReport").attr("disabled","disabled");
				}

				if(optionData.exibeMunicipio){
					$("#cod_municipio").val("");
					$("#cod_municipio").closest("div.row").removeClass("hide");
				}
				else if(!$("#cod_municipio").closest("div.row").hasClass("hide"))
					$("#cod_municipio").closest("div.row").addClass("hide");

				if(optionData.exibeLocalidade) {
					$("#cod_municipio").val("");
					$("#cod_empreendimento").closest("div.row").removeClass("hide");
				}
				else if(!$("#cod_empreendimento").closest("div.row").hasClass("hide"))
					$("#cod_empreendimento").closest("div.row").addClass("hide");

				if( location.search.indexOf("&filtra_situacao=sim") != -1 ){
					var cod_modelo_relatorio = $("#modelo_relatorio").val();
					var url = location.search.replace("&filtra_situacao=sim","").replace("cod_modelo_relatorio=informacao-obra-andamento.asp", "cod_modelo_relatorio="+cod_modelo_relatorio);
					location.replace("inicio-daee.asp"+url)
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
				<a class="navbar-brand" href="#">Sistema de Informações Gerenciais</a>
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
		<div class="jumbotron">
			<div class="row">
				<div class="col-xs-2">
					<img src="Brasão SP.jpg" class="img-responsive img-topo">
				</div>
				
				<div class="col-xs-8">
					<h3 class="text-center page-title">
						SECRETARIA DE SANEAMENTO E RECURSOS HÍDRICOS
						<br/>
						<strong>DEPARTAMENTO DE ÁGUAS E ENERGIA ELÉTRICA</strong>
						<br/>
						<small>PROGRAMA ÁGUA LIMPA</small>
					</h3>
				</div>

				<div class="col-xs-2">
					<img src="logo_daee.jpg" class="img-responsive img-topo pull-right">
				</div>
			</div>
		</div>
	</div>

	<div class="container container-box">
		<div class="panel panel-primary">
			<div class="panel-heading">
				<h3 class="panel-title">Selecione as informações que deseja visualizar</h3>
			</div>
			<div class="panel-body">
				<div class="row">
					<div class="col-xs-12">	
						<div class="form-group">
							<label class="control-label">Modelo de Relatório:</label>
							<select id="modelo_relatorio" class="form-control">
								<option></option>
								<option 
									<% If Request("cod_modelo_relatorio") <> "" Then If Request("cod_modelo_relatorio") = "resumo-situacao.asp" Then Response.Write "selected='selected'" End If End If %>
									data-habilita-botoes="true" data-exibe-municipio="false" data-exibe-localidade="false" value="resumo-situacao.asp">
									Resumo da Situação por Município e Localidade
								</option>
								<option
									<% If Request("cod_modelo_relatorio") <> "" Then If Request("cod_modelo_relatorio") = "informacao-obra.asp" Then Response.Write "selected='selected'" End If End If %>
									data-habilita-botoes="true" data-exibe-municipio="false" data-exibe-localidade="false" value="informacao-obra.asp">
									Informação Geral das Obras
								</option>
								<option 
									<% If Request("cod_modelo_relatorio") <> "" Then If Request("cod_modelo_relatorio") = "informacao-municipio.asp" Then Response.Write "selected='selected'" End If End If %>
									data-habilita-botoes="false" data-exibe-municipio="true" data-exibe-localidade="false" value="informacao-municipio.asp">
									Dados do Município
								</option>
								<option 
									<% If Request("cod_modelo_relatorio") <> "" Then If Request("cod_modelo_relatorio") = "ficha-tecnica-obra.asp" Then Response.Write "selected='selected'" End If End If %>
									data-habilita-botoes="false" data-exibe-municipio="true" data-exibe-localidade="true" value="ficha-tecnica-obra.asp">
									Ficha Técnica da Obra
								</option>
								<option 
									<% If Request("cod_modelo_relatorio") <> "" Then If Request("cod_modelo_relatorio") = "informacao-obra-andamento.asp" Then Response.Write "selected='selected'" End If End If %>
									data-habilita-botoes="false" data-exibe-municipio="true" data-exibe-localidade="true" value="informacao-obra-andamento.asp">
									Informações de Obra em Andamento
								</option>
							</select>
						</div>
					</div>	
				</div>	

				<div class='row <% If Request("cod_modelo_relatorio") = "" Then Response.Write "hide" End If %>'>
					<div class="col-xs-12">
						<div class="form-group">
							<label class="control-label">Município:</label>
							<select id="cod_municipio" class="form-control">
								<option value=""></option>
								<%
									strQ = "SELECT c_lista_predios.cod_predio, c_lista_predios.nme_municipio FROM c_lista_pi INNER JOIN c_lista_predios ON c_lista_pi.cod_mun = c_lista_predios.id_predio GROUP BY c_lista_predios.cod_predio, c_lista_predios.nme_municipio, c_lista_pi.nme_municipio ORDER BY c_lista_pi.nme_municipio;"

									Set rs_combo = Server.CreateObject("ADODB.Recordset")
										rs_combo.CursorLocation = 3
										rs_combo.CursorType = 3
										rs_combo.LockType = 1
										rs_combo.Open strQ, objCon, , , &H0001

									If Not rs_combo.EOF Then
										While Not rs_combo.EOF
											If Trim(rs_combo.Fields.Item("nme_municipio").Value) <> "" Then
												If Request("cod_mun") <> "" Then
													cod_mun = Request("cod_mun")
													Response.Write "      <OPTION value='" & (rs_combo.Fields.Item("cod_predio").Value) & "'"
													If Lcase(rs_combo.Fields.Item("cod_predio").Value) = Lcase(cod_mun) then
														Response.Write "selected"
													End If
													Response.Write ">" & (rs_combo.Fields.Item("nme_municipio").Value) & "</OPTION>"
												Else
								%>
								<option value="<%=(rs_combo.Fields.Item("cod_predio").Value)%>"><%=(rs_combo.Fields.Item("nme_municipio").Value)%></option>
								<%
												End If
											End If
											rs_combo.MoveNext
										Wend
									End If
								%>
							</select>
						</div>
					</div>
				</div>

				<div class='row <% If Request("cod_mun") = "" Then Response.Write "hide" End If %>'>
					<div class="col-xs-12">
						<div class="form-group">
							<label class="control-label">Localidade:</label>
							<select id="cod_empreendimento" class="form-control">
								<option value=""></option>
								<%
									If Request("cod_mun") <> "" Then
										cod_mun = Request("cod_mun")

										strQ = "SELECT * FROM tb_pi where id_predio = " & cod_mun

										If Request("filtra_situacao") <> "" Then
											strQ = strQ + " AND cod_situacao_externa = 40"
										End If

										Set rs_combo = Server.CreateObject("ADODB.Recordset")
											rs_combo.CursorLocation = 3
											rs_combo.CursorType = 3
											rs_combo.LockType = 1
											rs_combo.Open strQ, objCon, , , &H0001

										If Not rs_combo.EOF Then
											While Not rs_combo.EOF
												If Trim(rs_combo.Fields.Item("PI").Value) <> "" Then
								%>
								<option value="<%=(rs_combo.Fields.Item("PI").Value)%>"><%=(rs_combo.Fields.Item("PI").Value)%> - <%=(rs_combo.Fields.Item("nome_empreendimento").Value)%></option>
								<%
												End If
												rs_combo.MoveNext
											Wend
										End If
									End If
								%>
							</select>
						</div>
					</div>
				</div>
			</div>

			<div class="panel-footer clearfix">
				<div class="pull-right">
					<button type="button" id="btnClearForm" disabled="disabled" class="btn btn-default">
						<i class="fa fa-times-circle"></i> Cancelar
					</button>
					<button type="button" id="btnGenerateReport" disabled="disabled" class="btn btn-primary" data-loading-text="Aguarde...">
						<i class="fa fa-files-o"></i> Gerar Relatório
					</button>
				</div>
			</div>
		</div>
	</div>

</body>
</html>