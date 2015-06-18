<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="logout.asp" -->
<%
	Response.CharSet = "UTF-8"

	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	Dim dta_filtro

	sql = "SELECT * FROM c_lista_rel_ultimas_ocorrencias "

	If (Request.QueryString("data_inicio") <> "" And Request.QueryString("data_inicio") <> "0" And Request.QueryString("data_fim") <> "" And Request.QueryString("data_fim") <> "0") Then 
	  dta_filtro_inicio = Request.QueryString("data_inicio")
	  dta_filtro_fim 	= Request.QueryString("data_fim")

	  dta_inicio 	= Split(dta_filtro_inicio,"/")
	  dta_fim 		= Split(dta_filtro_fim,"/")

	  sql = sql & "WHERE [Data do Registro] BETWEEN #" & dta_inicio(1) & "/" & dta_inicio(0) & "/" & dta_inicio(2) & "# AND #" & dta_fim(1) & "/" & dta_fim(0) & "/" & dta_fim(2) & "#"

	  If(Request.QueryString("cod_tipo_registro") <> "" And Len(Request.QueryString("cod_tipo_registro") <> "") > 0) Then
	  	sql = sql & " AND cod_tipo_registro = " & Request.QueryString("cod_tipo_registro")
	  End If
	End If

	sql = sql & " ORDER BY municipio ASC, nome_empreendimento ASC, [Data do Registro] ASC"

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
	<script type="text/javascript" src="js/fullscreen.js"></script>
	<script type="text/javascript" src="js/moment.min.js"></script>
	<script type="text/javascript" src="js/jquery.table2excel.js"></script>
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
	<script type="text/javascript" src="js/fancybox/jquery.fancybox.pack.js?v=2.1.5"></script>
	<script type="text/javascript">
		$(function(){
			$("li a.print").on("click", function(){
				window.print();
			});

			$("li a.excel").on("click", function(){
				$("table").table2excel({
					name: "Relatório de Últimas Ocorrências"
				});
			});

			$('[data-toggle="tooltip"]').tooltip();
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
				<a class="navbar-brand" href="#">SIG - Relatório de Últimas Ocorrências</a>
			</div>

			<div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
				<ul class="nav navbar-nav navbar-right">
					<li><a href="javascript:window.close();"><i class="fa fa-times-circle"></i> Fechar Janela</a></li>
					<li><a href="#" class="print"><i class="fa fa-print"></i> Imprimir</a></li>
					<li><a href="#" class="excel"><i class="fa fa-file-excel-o"></i> Exportar p/ Excel</a></li>
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
						<h3><strong>Relatório de Últimas Ocorrências</strong><br/><small>Registros de Obra</small></h3>
					</div>
					<div class="col-xs-2"></div>
				</div>

				<hr/>

				<div class="row">
					<div class="col-xs-12">
						<table class="table table-bordered table-striped table-condensed table-hover">
							<thead>
								<%
									If Session("MM_UserAuthorization") = 1 or Session("MM_UserAuthorization") = 4 Then
								%>
								<th></th>
								<%
									End If
								%>
								<th></th>
								<th class="text-center">Município</th>
								<th class="text-center">Localidade</th>
								<th class="text-center">Responsável</th>
								<th class="text-center">Data Registro</th>
								<th class="text-center">Ação</th>
								<th class="text-center">Situação MSST</th>
								<th class="text-center">Houve Vistoria?</th>
								<th class="text-center">Data Vistoria</th>
								<th class="text-center">Tipo</th>
								<th class="text-center">Pendência?</th>
							</thead>
							<tbody>
								<%
									While (NOT rs_lista.EOF)
								%>
								<tr bgcolor="#F4F4F4">
									<%
										If Session("MM_UserAuthorization") = 1 or Session("MM_UserAuthorization") = 4 Then
									%>
									<td>
										<a class="btn btn-xs btn-default" data-toggle="tooltip" data-placement="right" title="Editar Registro"
											target="_blank" href="altera_acomp.asp?cod_acompanhamento=<%=(rs_lista.Fields.Item("cod_acompanhamento").Value)%>">
											<i class="fa fa-edit"></i>
										</a>
									</td>
									<%
										End If
									%>
									<td>
										<a class="btn btn-xs btn-primary" data-toggle="tooltip" data-placement="right" title="Visualizar Relatório de Obra"
											target="_blank" href="rel_rdo.asp?cod_empreendimento=<%=(rs_lista.Fields.Item("num_autos").Value)%>&data=<%=(rs_lista.Fields.Item("Data do Registro").Value)%>">
											<i class="fa fa-file-text-o"></i>
										</a>
									</td>
									<td width="150">
										<%=(rs_lista.Fields.Item("municipio").Value)%>
									</td>
									<td width="150">
										<%=(rs_lista.Fields.Item("num_autos").Value)%> - <%=(rs_lista.Fields.Item("nome_empreendimento").Value)%>
									</td>
									<td width="150">
										<%=(rs_lista.Fields.Item("nme_responsavel").Value)%>
									</td>
									<td class="text-center">
										<%=(rs_lista.Fields.Item("Data do Registro").Value)%>
									</td>
									<td>
										<%=(rs_lista.Fields.Item("Registro").Value)%>
									</td>
									<td class="text-center">
										<%=(rs_lista.Fields.Item("dsc_situacao_sso").Value)%>
									</td>
									<td class="text-center">
										<%=(rs_lista.Fields.Item("e_vistoria").Value)%>
									</td>
									<td class="text-center">
										<%=(rs_lista.Fields.Item("dt_vistoria").Value)%>
									</td>
									<td class="text-center">
										<%=(rs_lista.Fields.Item("dsc_tipo_registro").Value)%>
									</td>
									<td class="text-center">
										<%=(rs_lista.Fields.Item("e_pendencia").Value)%>
									</td>
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