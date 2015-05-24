<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="logout.asp" -->
<!--#include file="daee_restrict_access.asp" -->
<%
	Response.CharSet = "UTF-8"
	
  	Dim rs
  	Dim cod_municipio

  	Set cod_municipio = Request.QueryString("cod_municipio")

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.ActiveConnection = MM_cpf_STRING
	rs.Source = "SELECT * FROM (tb_responsavel RIGHT JOIN (tb_Construtora RIGHT JOIN (tb_predio LEFT JOIN tb_diretoria ON tb_predio.cod_bacia_secretaria = tb_diretoria.id) ON tb_Construtora.cod_construtora = tb_predio.cod_prefeitura) ON tb_responsavel.cod_fiscal = tb_predio.cod_prefeito) LEFT JOIN tb_partido ON tb_predio.cod_partido = tb_partido.id WHERE id_predio = " + Replace(cod_municipio, "'", "''")
	rs.CursorType = 0
	rs.CursorLocation = 2
	rs.LockType = 1
	rs.Open()
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
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
	<script type="text/javascript">
		$(function(){
			var vlr_lines = $(".vlr");
			$.each(vlr_lines, function(i, item){
				$(item).val($.number($(item).val(), 0, ",", "."));
			});

			$("li a.print").on("click", function(){
				window.print();
			});
		});
	</script>
	<style type="text/css">
		body { padding-top: 70px !important; }
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
				<a class="navbar-brand" href="#">SIG - Informações do Município</a>
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

	<div class="container">
		<div class="panel panel-default">
			<div class="panel-body">
				<form class="form-horizontal">
					<div class="form-group">
						<label for="inputNomePrefeitura" class="col-lg-3 control-label sr-only"></label>
						<div class="col-lg-9">
							<input type="text" class="form-control" readonly="readonly" id="inputNomePrefeitura"  value="<%=(rs.Fields.Item("Construtora").Value)%>">
						</div>
					</div>

					<div class="form-group">
						<label for="inputCNPJ" class="col-lg-3 control-label">CNPJ:</label>
						<div class="col-lg-3">
							<input type="text" class="form-control" readonly="readonly" id="inputCNPJ"  value="<%=(rs.Fields.Item("cnpj_empresa").Value)%>">
						</div>

						<label for="inputEndereco" class="col-lg-1 control-label">Endereço:</label>
						<div class="col-lg-5">
							<input type="text" class="form-control" readonly="readonly" id="inputEndereco"  value="<%=(rs.Fields.Item("Endereço da Construtora").Value)%>">
						</div>
					</div>

					<div class="form-group">
						<label for="inputCEP" class="col-lg-3 control-label">CEP:</label>
						<div class="col-lg-2">
							<input type="text" class="form-control" readonly="readonly" id="inputCEP"  value="<%=(rs.Fields.Item("cep_empresa").Value)%>">
						</div>

						<label for="inputNomeMunicipio" class="col-lg-1 control-label">Município:</label>
						<div class="col-lg-3">
							<input type="text" class="form-control" readonly="readonly" id="inputNomeMunicipio"  value="<%=(rs.Fields.Item("Município").Value)%>">
						</div>

						<label for="inputTelefone" class="col-lg-1 control-label">Telefone:</label>
						<div class="col-lg-2">
							<input type="text" class="form-control" readonly="readonly" id="inputTelefone"  value="<%=(rs.Fields.Item("Fone da Construtora").Value)%>">
						</div>
					</div>

					<div class="form-group">
						<label for="inputSite" class="col-lg-3 control-label">Site:</label>
						<div class="col-lg-4">
							<input type="text" class="form-control" readonly="readonly" id="inputSite"  value="<%=(rs.Fields.Item("site_empresa").Value)%>">
						</div>						

						<label for="inputEmail" class="col-lg-1 control-label">E-mail:</label>
						<div class="col-lg-4">
							<input type="text" class="form-control" readonly="readonly" id="inputEmail"  value="<%=(rs.Fields.Item("email_empresa").Value)%>">
						</div>
					</div>

					<div class="form-group">
						<label for="inputNomeDiretorBaciaDAEE" class="col-lg-3 control-label">Diretoria de Bacia - DAEE:</label>
						<div class="col-lg-4">
							<input type="text" class="form-control" readonly="readonly" id="inputNomeDiretorBaciaDAEE"  value="<%=(rs.Fields.Item("Diretoria de Ensino").Value)%>">
						</div>

						<label for="inputURGHI" class="col-lg-1 control-label">UGRHI:</label>
						<div class="col-lg-4">
							<input type="text" class="form-control" readonly="readonly" id="inputURGHI"  value="<%=(rs.Fields.Item("desc_diretoria").Value)%>">
						</div>
					</div>

					<div class="form-group">
						<label for="inputPopulacao2010" class="col-lg-3 control-label">População Urbana - IBGE(2010) (hab):</label>
						<div class="col-lg-3">
							<input type="text" class="form-control vlr" readonly="readonly" id="inputPopulacao2010"  value="<%=(rs.Fields.Item("qtd_populacao_urbana_2010").Value)%>">
						</div>

						<label for="inputPopulacao2030" class="col-lg-3 control-label">Projeção de População (2030):</label>
						<div class="col-lg-3">
							<input type="text" class="form-control vlr" readonly="readonly" id="inputPopulacao2030"  value="<%=(rs.Fields.Item("qtd_populacao_urbana_2030").Value)%>">
						</div>
					</div>

					<div class="form-group">
						<label for="inputNomePrefeito" class="col-lg-3 control-label">Nome do Prefeito:</label>
						<div class="col-lg-9">
							<input type="text" class="form-control" readonly="readonly" id="inputNomePrefeito"  value="<%=(rs.Fields.Item("Responsável").Value)%>">
						</div>
					</div>

					<div class="form-group">
						<label for="inputTelefonePrefeito" class="col-lg-3 control-label">Telefone (Prefeito):</label>
						<div class="col-lg-2">
							<input type="text" class="form-control" readonly="readonly" id="inputTelefonePrefeito"  value="<%=(rs.Fields.Item("num_telefone").Value)%>">
						</div>

						<label for="inputEmailPrefeito" class="col-lg-2 control-label">E-mail (Prefeito):</label>
						<div class="col-lg-5">
							<input type="text" class="form-control" readonly="readonly" id="inputEmailPrefeito"  value="<%=(rs.Fields.Item("email").Value)%>">
						</div>
					</div>

					<div class="form-group">
						<label for="inputPeriodoAdministracao" class="col-lg-3 control-label">Período de Administração:</label>
						<div class="col-lg-2">
							<input type="text" class="form-control text-center" readonly="readonly" id="inputPeriodoAdministracao"  value="<%=(rs.Fields.Item("ano_inicio_adm").Value)%>">
						</div>

						<label for="inputPeriodoAdministracao" class="col-lg-1 control-label">Até:</label>
						<div class="col-lg-2">
							<input type="text" class="form-control text-center" readonly="readonly" id="inputPeriodoAdministracao"  value="<%=(rs.Fields.Item("ano_fim_adm").Value)%>">
						</div>

						<label for="inputAtendidoPor" class="col-lg-2 control-label">Atendido pela Sabesp:</label>
						<div class="col-lg-2">
							<input type="text" class="form-control text-center" readonly="readonly" id="inputAtendidoPor"  value="<% If rs.Fields.Item("flg_atendido_sabesp").Value = 1 Then Response.Write "Sim" Else Response.Write "Não" End If %>">
						</div>
					</div>
				</form>
			</div>
		</div>
	</div>

</body>
</html>