<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="logout.asp" -->
<%
	Response.CharSet = "UTF-8"

	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

  	cod_empreendimento = Request.QueryString("cod_empreendimento")

  	strQ = "SELECT * FROM c_lista_pi WHERE PI = '"& cod_empreendimento &"'"

	Set rs_dados_obra = Server.CreateObject("ADODB.Recordset")
		rs_dados_obra.CursorLocation = 3
		rs_dados_obra.CursorType = 3
		rs_dados_obra.LockType = 1
		rs_dados_obra.Open strQ, objCon, , , &H0001

	sql = "SELECT * FROM c_lista_todas_fotos_obra WHERE PI = '"& cod_empreendimento &"' ORDER BY id_arquivo DESC"

	Dim rs_fotos

	Set rs_fotos = Server.CreateObject("ADODB.Recordset")
	rs_fotos.PageSize = 12
	rs_fotos.Open sql, objCon, 3, 3

	pg = 0
	rec = 0

	If Not rs_fotos.EOF Then
		If Request.QueryString("pg") = "" Then
			pg = 1
		Else
			If CInt(Request.QueryString("pg")) < 1 Then
				pg = 1
			Else
				pg = Request.QueryString("pg")
			End If
		End If
		
		rs_fotos.AbsolutePage = pg
	End If
%>
<!DOCTYPE html>
<html>
<head>
	<title>:: DAEE ::</title>
	<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
	<link rel="stylesheet" type="text/css" href="css/bootstrap-flaty.min.css">
	<link rel="stylesheet" type="text/css" href="css/daee.css">
	<link rel="stylesheet" href="js/fancybox/jquery.fancybox.css?v=2.1.5" type="text/css" media="screen" />
	<link rel="stylesheet" href="//maxcdn.bootstrapcdn.com/font-awesome/4.3.0/css/font-awesome.min.css">

	<style type="text/css">
		div.thumbnail img {
			min-height: 118px;
			max-height: 118px;
		}

		ul.pagination {
			margin: 0;
		}
		
		img.selected {
			border: 2px solid #2C3E50;
			-webkit-box-shadow: 0px 0px 18px -1px rgba(0,0,0,0.77);
			   -moz-box-shadow: 0px 0px 18px -1px rgba(0,0,0,0.77);
					box-shadow: 0px 0px 18px -1px rgba(0,0,0,0.77);
		}
	</style>

	<script type="text/javascript" src="//code.jquery.com/jquery-1.11.2.min.js"></script>
	<script type="text/javascript" src="js/jquery.number.min.js"></script>
	<script type="text/javascript" src="js/jquery.table2excel.js"></script>
	<script type="text/javascript" src="js/jquery.lazyload.min.js"></script>
	<script type="text/javascript" src="js/moment.min.js"></script>
	<script type="text/javascript" src="js/fullscreen.js"></script>
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
	<script type="text/javascript" src="js/fancybox/jquery.fancybox.pack.js?v=2.1.5"></script>
	<script type="text/javascript">
		$(function(){
			$(".fancybox").fancybox();
			$("img.lazy").lazyload({
				effect : "fadeIn"
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
				<a class="navbar-brand" href="#">SIG - Fotos p/ Relatórios</a>
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
		<div class="panel panel-default">
			<div class="panel-heading">
				<h3 class="panel-title"><i class="fa fa-building-o"></i> Informações da Obra</h3>
			</div>
			
			<div class="panel-body">
				<form class="form-horizontal">
					<div class="form-group">
						<label class="col-lg-1 control-label">Município:</label>
						<div class="col-lg-4">
							<input type="text" class="form-control" readonly="readonly" value="<%=(rs_dados_obra.Fields.Item("nme_municipio").Value)%>">
						</div>

						<label class="col-lg-1 control-label">Localidade:</label>
						<div class="col-lg-3">
							<input type="text" class="form-control" readonly="readonly" value="<%=(rs_dados_obra.Fields.Item("nome_empreendimento").Value)%>">
						</div>

						<label class="col-lg-1 control-label">Nº Autos:</label>
						<div class="col-lg-2">
							<input type="text" class="form-control" readonly="readonly" value="<%=(rs_dados_obra.Fields.Item("PI").Value)%>">
						</div>
					</div>
				</form>
			</div>
		</div>

		<div class="panel panel-default">
			<div class="panel-heading clearfix">
				<div class="pull-left">
					<h3 class="panel-title"><i class="fa fa-picture-o"></i> Ferramenta de Seleção de Fotos p/ Relatórios</h3>
				</div>
				<div class="pull-right">
					<ul class="pagination pagination-sm">
						<%
							i = 1
							While i <= rs_fotos.PageCount
								If CInt(pg) = i Then
						%>
						<li class="active"><a href="?cod_empreendimento=<%=(cod_empreendimento)%>&pg=<%=(i)%>"><%=(i)%></a></li>
						<%
								Else
						%>
						<li><a href="?cod_empreendimento=<%=(cod_empreendimento)%>&pg=<%=(i)%>"><%=(i)%></a></li>
						<%
								End If
								i = (i+1)
							Wend
						%>
					</ul>
				</div>
			</div>

			<div class="panel-body">
				<div class="row">
					<%
						If rs_fotos.EOF Then
					%>
					<div class="col-lg-12">
						Nenhuma foto encontrada!
					</div>
					<%
						End If

						While ((rec < rs_fotos.PageSize) And (Not rs_fotos.EOF))
							pth_url = LCase(rs_fotos.Fields.Item("pth_arquivo").Value)
							pth_url = Replace(pth_url, "\\10.0.75.125\intermultiplas.net\public\", "")
							pth_url = Replace(pth_url, "e:\home\programaagualimpa\web\", "")
							pth_url = Replace(pth_url, "\", "/")
							img_url = pth_url

							If Not rs_fotos.Fields.Item("flg_pmweb_file").Value Then
								img_url = img_url & rs_fotos.Fields.Item("cod_referencia").Value & "_"
							End If

							img_url = img_url & rs_fotos.Fields.Item("nme_arquivo").Value
					%>
					<div class="col-xs-12 col-sm-2 col-md-2">
						<div class="thumbnail">
							<%
								If rs_fotos.Fields.Item("report").Value Then
							%>
							<img src="<%=(img_url)%>" class="lazy selected" alt="">
							<%
								Else
							%>
							<img src="<%=(img_url)%>" class="lazy" alt="">
							<%
								End If
							%>
							<div class="caption">
								<!-- <h4>Thumbnail label</h4> -->
								<!--<p class="thumbnail-label"><%=(rs_fotos.Fields.Item("dsc_observacoes").Value)%></p>-->
								<p>
									<a href="<%=(img_url)%>" rel="group" title="<%=(rs_fotos.Fields.Item("dsc_observacoes").Value)%>" class="btn btn-default btn-block btn-xs fancybox" role="button"><i class="fa fa-expand"></i> Ampliar imagem</a>
								</p>
								<p>
									<form method="post" action="sql_update.asp">
										<input type="hidden" name="sql_query" value="UPDATE tb_acompanhamento_arquivo SET flg_report_file = <% If rs_fotos.Fields.Item("report").Value Then Response.Write "0" Else Response.Write "1" End If %> WHERE id_arquivo = <%=(rs_fotos.Fields.Item("id_arquivo").Value)%>">
										<input type="hidden" name="url_redirect" value="image_tool.asp?<%=(Request.QueryString)%>">
										<%
											If Not rs_fotos.Fields.Item("report").Value Then
										%>
										<button type="submit" class="btn btn-block btn-xs btn-success">Marcar</button>
										<%
											Else
										%>
										<button type="submit" class="btn btn-block btn-xs btn-primary">Desmarcar</button>
										<%
											End If
										%>
									</form>
								</p>
							</div>
						</div>
					</div>
					<%
							rs_fotos.MoveNext()
							rec = (rec+1)
						Wend
					%>
				</div>
			</div>

			<div class="panel-footer text-center">
				<ul class="pagination">
					<%
						i = 1
						While i <= rs_fotos.PageCount
							If CInt(pg) = i Then
					%>
					<li class="active"><a href="?cod_empreendimento=<%=(cod_empreendimento)%>&pg=<%=(i)%>"><%=(i)%></a></li>
					<%
							Else
					%>
					<li><a href="?cod_empreendimento=<%=(cod_empreendimento)%>&pg=<%=(i)%>"><%=(i)%></a></li>
					<%
							End If
							i = (i+1)
						Wend
					%>
				</ul>
			</div>
		</div>
	</div>

</body>
</html>