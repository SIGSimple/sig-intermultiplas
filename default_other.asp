<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
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
<body class="login">
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
					<li class="dropdown">
						<a href="#" class="dropdown-toggle" data-toggle="dropdown">
							<i class="fa fa-sign-in"></i> Entrar no Sistema
						</a>
						<ul class="dropdown-menu" style="padding: 15px;min-width: 250px;">
							<li>
								<div class="row">
									<div class="col-md-12">
										<form class="form" role="form" method="post" action="login" accept-charset="UTF-8" id="login-nav">
											<div class="form-group">
												<label class="sr-only" for="inputEmail">Usuário</label>
												<input type="email" class="form-control" id="inputEmail" placeholder="Usuário" required>
											</div>
											<div class="form-group">
												<label class="sr-only" for="inputSenha">Senha</label>
												<input type="password" class="form-control" id="inputSenha" placeholder="Senha" required>
											</div>
											<div class="form-group">
												<button type="submit" class="btn btn-info btn-block">Entrar</button>
											</div>
										</form>
									</div>
								</div>
							</li>

							<li class="divider"></li>

							<li>
								<button class="btn btn-block btn-default">Esqueci minha senha</button>
							</li>
						</ul>
					</li>
				</ul>
			</div>
		</div>
	</nav>

	<div class="container login text-center">
		<div class="page-header">
			<img src="Cons&oacute;rcio Maior.jpg" class="logo-intermultiplas"/>
		</div>
	</div>
</body>
</html>