<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/cpf.asp" -->
<%
	'*******************************************************************
	' Página gerada pelo sistema Dataform 2 - http://www.dataform.com.br
	'*******************************************************************
	' Altere os valores das variáveis indicadas abaixo se necessário

	'String de conexão para o banco de dados do Microsoft Access
	strCon = "DBQ=C:\inetpub\wwwroot\original\ARQUIVOS\DADOS\bd_fde.mdb;Driver={Microsoft Access Driver (*.mdb)};"
	'strCon = "DBQ=\\10.0.75.124\intermultiplas.net\public\ARQUIVOS\DADOS\bd_fde.mdb;Driver={Microsoft Access Driver (*.mdb)};"
	
	'Nome da página de consulta
	pagina_consulta = "df_consulta.asp"

	'Nome da página de inclusão
	pagina_inclusao = "df_inclusao.asp"

	'*******************************************************************
%>
<html>
	<head>
		<title>Efetuar Login</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
		<meta name="copyright" content="Dataform">
		<meta name="keywords" content="dataform, asp dataform, aspdataform, asp-dataform">
		<meta name="robots" content="ALL">
		<link rel="stylesheet" type="text/css" href="../css/login.css">
		<script type="text/javascript" src="../js/login.js"></script>
	</head>
	<body class="texto_pagina">
		<img src="imagens/login.gif" align="left">
		
		<%
			If Not IsEmpty(Request.Form("enviar")) Then
				Set objCon = Server.CreateObject("ADODB.Connection")
				objCon.Open MM_cpf_STRING
				Set objRS = Server.CreateObject("ADODB.Recordset")
				objRS.CursorLocation = 3
				objRS.CursorType = 3
				objRS.LockType = 1
				login = Trim(Lcase(Request.Form("login")))
				login = Replace(login, "'", "")
				login = Replace(login, "|", "")
				senha = Trim(Lcase(Request.Form("senha")))
				senha = Replace(senha, "'", "")
				senha = Replace(senha, "|", "")
				strQ = "SELECT * FROM login WHERE nome LIKE '" & login & "' AND senha LIKE '" & senha & "'"
				objRS.Open strQ, objCon, , , &H0001
				If Not objRS.EOF Then
					Session("admin") = login
					Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR")
		%>

		<br>
		<br>
		<b>Seja Bem Vindo ao Sistema</b>
		<br>
		Clique no link abaixo para entrar na p&aacute;gina desejada.
		<br>
		<br>
		<br>
		- <A href="<%=pagina_inclusao%>">P&aacute;gina de Inclus&atilde;o</A><br>
		<br>
		- <A href="<%=pagina_consulta%>">P&aacute;gina de Consulta</A><br>
		<br>
		<br>
		<br>
		SEUS DADOS DE ACESSO:
		<br>
		<br>
		Admin: <strong><%=Session("admin")%></strong>
		<br>
		IP: <strong><%=Request.ServerVariables("REMOTE_ADDR")%></strong>
		<br>
		<br>
		<br>
		
		<%
				Else
		%>

		<br><br>
		<strong>Acesso Negado</strong><br>
		O login ou a senha informados n&atilde;o correspondem<br>
		<br>
		<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>">Clique aqui</a> para tentar novamente<br><br>

		<%
				End If
			
				objRS.Close
				Set objRS = Nothing
				
				objCon.Close
				Set objCon = Nothing
			Else
		%>
		
		<br>
		<br>
		<b>Efetuar Login</b>
		<br>
		<br>
		Informe abaixo seu login e senha para ter acesso as p&aacute;ginas protegidas do sistema:
		<br>

		<form name="form_incluir" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" onSubmit="return verifica_form(this);">
			<input type="hidden" name="recordno" value="<%=Request.Form("recordno")%>">
			<input type="hidden" name="strQ" value="<%=Request.Form("strQ")%>">
			
			<table border="0" cellpadding="2" cellspacing="1" class="tabela_formulario">
				<tr class="titulo_campos">
					<td>
						Login
						<br>
						<input style="width: 200px;" type="text" name="login" maxlength="50" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class="campos_formulario">
					</td>
				</tr>
				<tr class=titulo_campos>
					<td>
						Senha
						<br>
						<input style="width: 200px;" type="password" name="senha" maxlength="10" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class="campos_formulario">
					</td>
				</tr>
			</table>

			<input name="enviar" type="submit" class=botao_enviar id="enviar" value="Enviar">
		</form>

		<br>
		<br>
		<br>
		SISTEMA TOTALMENTE SEGURO
		<br>
		<br>
		Protegido por login, senha e endere&ccedil;o de IP da m&aacute;quina.
		<br>
		Ou seja, mesmo quando logado ao sistema, somente seu
		<br>
		computador ter&aacute; acesso &aacute;s p&aacute;ginas protegidas.
		<br>
		<br>
		Seu IP: <%=Request.ServerVariables("REMOTE_ADDR")%>
		<br>
		<br>

		<%
			End If
		%>
	</body>
</html>