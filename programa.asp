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
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
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

	<div class="container container-box text-center">
		<a href="atendimento-prefeito.asp" class="btn btn-lg btn-primary">Atendimento a Prefeitos</a>
		<a href="inicio-daee.asp" class="btn btn-lg btn-default">Relatórios Gerenciais</a>
	</div>

</body>
</html>