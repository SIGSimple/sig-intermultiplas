<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")

If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
	MM_valUsername = CStr(Request.Form("usuario"))
If MM_valUsername <> "" Then
	MM_fldUserAuthorization = "nivel"
	MM_redirectLoginSuccess = "inicio.asp"
	MM_redirectLoginFailed = "default.asp?msgError=Usuário ou Senha inválidos!"
	MM_flag = "ADODB.Recordset"
	
	Set MM_rsUser = Server.CreateObject(MM_flag)
	
	MM_rsUser.ActiveConnection = MM_cpf_STRING
	MM_rsUser.Source = "SELECT idusuario, nome, senha"
	
	If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
		MM_rsUser.Source = MM_rsUser.Source & " FROM login WHERE nome='" & Replace(MM_valUsername,"'","''") &"' AND senha='" & Replace(Request.Form("senha"),"'","''") & "'"
		MM_rsUser.CursorType = 0
		MM_rsUser.CursorLocation = 2
		MM_rsUser.LockType = 3
		MM_rsUser.Open
		
		If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
			' username and password match - this is a valid user
			Session("MM_Username") = MM_valUsername
			Session("MM_Userid") = CStr(MM_rsUser.Fields.Item("idusuario").Value)
			If (MM_fldUserAuthorization <> "") Then
				Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
				If Session("MM_UserAuthorization") = 1 Then
					Session("MM_UserAuthorization_Admin") = True
				Else
					If Session("MM_UserAuthorization") = 8 Then
						MM_redirectLoginSuccess = "inicio-daee.asp"
					Else
						If Session("MM_UserAuthorization") = 9 Then
							MM_redirectLoginSuccess = "atendimento-prefeito.asp"
						Else
							Session("MM_UserAuthorization_Admin") = False
						End If
					End If
				End If
			Else
				Session("MM_UserAuthorization") = ""
		End If

		If CStr(Request.QueryString("accessdenied")) <> "" And true Then
			MM_redirectLoginSuccess = Request.QueryString("accessdenied")
		End If
		
		MM_rsUser.Close
		Response.Redirect(MM_redirectLoginSuccess)
	End If
	
	MM_rsUser.Close
	Response.Redirect(MM_redirectLoginFailed)
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
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
	<script type="text/javascript">
		function showPassword() {
			var key_attr = $('#senha').attr('type');
			if(key_attr != 'text') {
				$('.checkbox').addClass('show');
				$('#senha').attr('type', 'text');
			} else {
				$('.checkbox').removeClass('show');
				$('#senha').attr('type', 'password');
			}	
		}

		$(function() {
			$("#loginform").on("submit", function(e){
				$("#btn-login").button("loading");
				if(e.currentTarget.usuario.value == "" || e.currentTarget.senha.value == ""){
					$(".alert.alert-warning").removeClass("hide");
					$("#btn-login").button("reset");
					return false;
				}
				return true;
			});
		});
	</script>
</head>
<body class="login">
	<section id="login">
		<div class="container container-login development">
			<div class="row">
				<div class="col-xs-12">
					<div class="form-wrap">
						<%
							If Not isInDevelopment Then
						%>
						<h1>Entre com seus dados</h1>
						<%
							End If

							If Request.QueryString("accessdenied") <> "" And True Then
						%>
						<div class="alert alert-danger">
							Você precisa estar logado ou você não tem permissões para ver a tela solicitada!<br/>
							Feche o navegador e tente acessar novamente.
						</div>
						<%
							End If
						%>

						<%
							If Request.QueryString("msgError") <> "" And True Then
						%>
						<div class="alert alert-danger">
							<%=(Request.QueryString("msgError"))%>
						</div>
						<%
							End If
						%>

						<%
							If isInDevelopment Then
						%>
						<div class="alert alert-warning text-center">
							<i class="fa fa-warning fa-3x"></i>
							<br/>
							<%=userFriendlyMessage%>
						</div>
						<%
							Else
						%>

						<div class="alert alert-warning hide"><i class="fa fa-warning"></i> Dados inválidos!</div>

						<form role="form" method="post" id="loginform" autocomplete="off" action="<%=MM_LoginAction%>">
							<div class="form-group">
								<label for="usuario" class="sr-only">Email</label>
								<input type="text" name="usuario" id="usuario" class="form-control" placeholder="Usuário">
							</div>

							<div class="form-group">
								<label for="senha" class="sr-only">Senha</label>
								<input type="password" name="senha" id="senha" class="form-control" placeholder="Senha">
							</div>

							<div class="checkbox">
								<span class="character-checkbox" onclick="showPassword()"></span>
								<span class="label">Exibir senha</span>
							</div>

							<input type="submit" id="btn-login" class="btn btn-custom btn-lg btn-block" data-loading-text="Aguarde..." value="Entrar">
						</form>
						<%
							End If
						%>

						<div class="alert alert-danger text-center">
							<i class="fa fa-database fa-3x"></i>
							<br/>
							ÁREA EXCLUSIVA PARA PROGRAMADORES E EQUIPE DE DESENVOLVIMENTO
						</div>

						<!-- <a href="#" class="forget" data-toggle="modal" data-target=".forget-modal">Esqueci minha senha</a> -->
					</div>
				</div>
			</div>
		</div>
	</section>

	<div class="modal fade forget-modal" tabindex="-1" role="dialog" aria-labelledby="myForgetModalLabel" aria-hidden="true">
		<div class="modal-dialog modal-sm">
			<div class="modal-content">
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal">
						<span aria-hidden="true">×</span>
						<span class="sr-only">Close</span>
					</button>
					<h4 class="modal-title">Recuperar senha</h4>
				</div>
				<div class="modal-body">
					<p>Informe seu e-mail de cadastro</p>
					<input type="email" name="recovery-email" id="recovery-email" class="form-control" autocomplete="off">
				</div>
				<div class="modal-footer">
					<button type="button" class="btn btn-default" data-dismiss="modal">Cancelar</button>
					<button type="button" class="btn btn-custom">Enviar</button>
				</div>
			</div>
		</div>
	</div>
</body>
</html>