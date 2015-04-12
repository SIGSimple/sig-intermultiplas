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
	<script type="text/javascript" src="js/jquery.floatThead.min.js"></script>
	<script type="text/javascript" src="js/jquery.table2excel.js"></script>
	<script type="text/javascript" src="js/fullscreen.js"></script>
	<script type="text/javascript">
		$(function(){
			$("table.table").floatThead({
				scrollingTop: 60
			});

			$("li a.print").on("click", function(){
				window.print();
			});

			$("li a.excel").on("click", function(){
				$("table.hide").table2excel({
					name: "Informação Geral das Obras"
				});
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
				<a class="navbar-brand" href="#">SIG - Informação Geral das Obras</a>
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
			<th class="text-middle text-center" rowspan="2" style="min-width: 300px;">Objeto</th>
			<th class="text-middle text-center" colspan="2">Benefício Ambiental</th>
			<th class="text-middle text-center" colspan="9">Informações Complementares</th>
			<th class="text-middle text-center" rowspan="2" style="min-width: 300px;">Notas / Observações</th>
			<th class="text-middle text-center" rowspan="2" style="min-width: 300px;">Benefício Geral da Obra</th>
		</tr>
		<tr class="info">
			<th class="text-middle text-center" style="min-width: 250px;">Bacia Hidrográfica (rio principal)</th>
			<th class="text-middle text-center" style="min-width: 190px;">Manancial de Lançamento</th>
			<th class="text-middle text-center" style="min-width: 190px;">Localização Geográfica</th>
			<th class="text-middle text-center" style="min-width: 190px;">Coletor Tronco (m)</th>
			<th class="text-middle text-center" style="min-width: 115px;">Interceptor (m)</th>
			<th class="text-middle text-center" style="min-width: 250px;">Emissário Efluente Bruto (m)</th>
			<th class="text-middle text-center" style="min-width: 250px;">Estação Elevatória de Esgoto (unid)</th>
			<th class="text-middle text-center" style="min-width: 190px;">Linha de Recalque (m)</th>
			<th class="text-middle text-center" style="min-width: 250px;">Emissário de Efluente Tratado (m)</th>
			<th class="text-middle text-center" style="min-width: 190px;">Tipo de ETE</th>
			<th class="text-middle text-center" style="min-width: 300px;">Estação de Tratamento</th>
		</tr>
		<%
			strQ = "SELECT tb_PI.*, tb_tipo_empreendimento.desc_tipo AS tipo_empreendimento, [tb_depto].[sigla]+' - '+[tb_depto].[desc_depto] AS programa, tb_predio.Município, tb_responsavel.Responsável AS eng_obras_consorcio, tb_responsavel_1.Responsável AS eng_daee, tb_responsavel_2.Responsável AS eng_plan_consorcio, tb_responsavel_3.Responsável AS fiscal_consorcio, tb_responsavel_4.Responsável AS eng_obras_construtora, [num_autos]+' - '+[num_convenio] AS convenio, tb_situacao_pi.desc_situacao AS desc_situacao_interna, tb_situacao_pi_1.desc_situacao AS desc_situacao_externa, tb_predio.[Diretoria de Ensino] AS bacia_daee, tb_bacia_hidrografica.nme_bacia_hidrografica, tb_manancial_lancamento.nme_manancial, tb_tipo_ete.nme_tipo_ete, IIf([flg_estudo_elaborado_daee]=1,'Sim','Não') AS estudo_elaborado_daee FROM tb_manancial_lancamento RIGHT JOIN (tb_bacia_hidrografica RIGHT JOIN (tb_tipo_ete RIGHT JOIN (((tb_convenio RIGHT JOIN (tb_responsavel AS tb_responsavel_4 RIGHT JOIN (tb_responsavel AS tb_responsavel_3 RIGHT JOIN (tb_responsavel AS tb_responsavel_2 RIGHT JOIN (tb_responsavel AS tb_responsavel_1 RIGHT JOIN (tb_depto RIGHT JOIN (tb_tipo_empreendimento RIGHT JOIN (tb_responsavel RIGHT JOIN (tb_predio RIGHT JOIN tb_PI ON tb_predio.cod_predio = tb_PI.cod_predio) ON tb_responsavel.cod_fiscal = tb_PI.cod_fiscal) ON tb_tipo_empreendimento.id = tb_PI.cod_tipo_empreendimento) ON tb_depto.cod_depto = tb_PI.cod_programa) ON tb_responsavel_1.cod_fiscal = tb_PI.cod_engenheiro_daee) ON tb_responsavel_2.cod_fiscal = tb_PI.cod_engenheiro_plan_consorcio) ON tb_responsavel_3.cod_fiscal = tb_PI.cod_fiscal_consorcio) ON tb_responsavel_4.cod_fiscal = tb_PI.cod_engenheiro_construtora) ON tb_convenio.id = tb_PI.cod_convênio) LEFT JOIN tb_situacao_pi ON tb_PI.cod_situacao = tb_situacao_pi.cod_situacao) LEFT JOIN tb_situacao_pi AS tb_situacao_pi_1 ON tb_PI.cod_situacao_externa = tb_situacao_pi_1.cod_situacao) ON tb_tipo_ete.id = tb_PI.cod_tipo_ete) ON tb_bacia_hidrografica.id = tb_PI.cod_bacia_hidrografica) ON tb_manancial_lancamento.id = tb_PI.cod_manancial_lancamento ORDER BY tb_predio.Município, tb_PI.nome_empreendimento;"

			Set rs_lista = Server.CreateObject("ADODB.Recordset")
				rs_lista.CursorLocation = 3
				rs_lista.CursorType = 3
				rs_lista.LockType = 1
				rs_lista.Open strQ, objCon, , , &H0001

			If Not rs_lista.EOF Then
				While Not rs_lista.EOF
		%>
		<tr>
			<td class="text-middle" style="max-width: 180px"><%=(rs_lista.Fields.Item("Município").Value)%></td>
			<td class="text-middle" style="max-width: 180px"><%=(rs_lista.Fields.Item("nome_empreendimento").Value)%></td>
			<td class="text-middle" style="max-width: 300px"><%=(rs_lista.Fields.Item("Descrição da Intervenção FDE").Value)%></td>					
			<td class="text-middle text-center" style="max-width: 250px;"><%=(rs_lista.Fields.Item("nme_bacia_hidrografica").Value)%></td>
			<td class="text-middle text-center" style="max-width: 190px;"><%=(rs_lista.Fields.Item("nme_manancial").Value)%></td>
			<td class="text-middle text-center" style="max-width: 190px;"><%=(rs_lista.Fields.Item("latitude_longitude").Value)%></td>
			<td class="text-middle text-center" style="max-width: 190px;"><%=(rs_lista.Fields.Item("qtd_metragem_coletor_tronco").Value)%></td>
			<td class="text-middle text-center" style="max-width: 115px;"><%=(rs_lista.Fields.Item("qtd_metragem_interceptor").Value)%></td>
			<td class="text-middle text-center" style="max-width: 250px;"><%=(rs_lista.Fields.Item("qtd_metragem_emissario_fluente_bruto").Value)%></td>
			<td class="text-middle text-center" style="max-width: 250px;"><%=(rs_lista.Fields.Item("qtd_eee").Value)%></td>
			<td class="text-middle text-center" style="max-width: 190px;"><%=(rs_lista.Fields.Item("qtd_metragem_linha_recalque").Value)%></td>
			<td class="text-middle text-center" style="max-width: 250px;"><%=(rs_lista.Fields.Item("qtd_metragem_emissario_efluente_tratado").Value)%></td>
			<td class="text-middle text-center" style="max-width: 190px;"><%=(rs_lista.Fields.Item("nme_tipo_ete").Value)%></td>
			<td class="text-middle text-center" style="max-width: 300px;"><%=(rs_lista.Fields.Item("dsc_estacao_tratamento").Value)%></td>
			<td class="text-middle" style="max-width: 300px;"><%=(rs_lista.Fields.Item("dsc_observacoes_obra").Value)%></td>
			<td class="text-middle" style="max-width: 300px;"><%=(rs_lista.Fields.Item("dsc_resultado_obtido").Value)%></td>
		</tr>
		<%
					rs_lista.MoveNext
				Wend
			End If
		%>
	</table>

	<table class="table table-bordered table-condensed table-hover table-striped table-responsive">
		<thead>
			<tr class="info">
				<th class="text-middle text-center" rowspan="2" style="min-width: 180px;">Município</th>
				<th class="text-middle text-center" rowspan="2" style="min-width: 180px;">Localidade</th>
				<th class="text-middle text-center" rowspan="2" style="min-width: 400px;">Objeto</th>
				<th class="text-middle text-center" colspan="2">Benefício Ambiental</th>
				<th class="text-middle text-center" colspan="9">Informações Complementares</th>
				<th class="text-middle text-center" rowspan="2" style="min-width: 300px;">Notas / Observações</th>
				<th class="text-middle text-center" rowspan="2" style="min-width: 300px;">Benefício Geral da Obra</th>
			</tr>
			<tr class="info">
				<th class="text-middle text-center" style="min-width: 150px;">Bacia Hidrográfica<br/>(rio principal)</th>
				<th class="text-middle text-center" style="min-width: 120px;">Manancial<br/>de Lançamento</th>
				<th class="text-middle text-center" style="min-width: 190px;">Localização Geográfica</th>
				<th class="text-middle text-center" style="min-width: 120px;">Coletor Tronco<br/>(m)</th>
				<th class="text-middle text-center" style="min-width: 115px;">Interceptor<br/>(m)</th>
				<th class="text-middle text-center" style="min-width: 150px;">Emissário de<br/>Efluente Bruto (m)</th>
				<th class="text-middle text-center" style="min-width: 50px;">EEE (unid)</th>
				<th class="text-middle text-center" style="min-width: 120px;">Linha de<br/>Recalque (m)</th>
				<th class="text-middle text-center" style="min-width: 150px;">Emissário de<br/>Efluente Tratado (m)</th>
				<th class="text-middle text-center" style="min-width: 190px;">Tipo de ETE</th>
				<th class="text-middle text-center" style="min-width: 300px;">Estação de Tratamento</th>
			</tr>
		</thead>

		<tbody>
			<%
				strQ = "SELECT tb_PI.*, tb_tipo_empreendimento.desc_tipo AS tipo_empreendimento, [tb_depto].[sigla]+' - '+[tb_depto].[desc_depto] AS programa, tb_predio.Município, tb_responsavel.Responsável AS eng_obras_consorcio, tb_responsavel_1.Responsável AS eng_daee, tb_responsavel_2.Responsável AS eng_plan_consorcio, tb_responsavel_3.Responsável AS fiscal_consorcio, tb_responsavel_4.Responsável AS eng_obras_construtora, [num_autos]+' - '+[num_convenio] AS convenio, tb_situacao_pi.desc_situacao AS desc_situacao_interna, tb_situacao_pi_1.desc_situacao AS desc_situacao_externa, tb_predio.[Diretoria de Ensino] AS bacia_daee, tb_bacia_hidrografica.nme_bacia_hidrografica, tb_manancial_lancamento.nme_manancial, tb_tipo_ete.nme_tipo_ete, IIf([flg_estudo_elaborado_daee]=1,'Sim','Não') AS estudo_elaborado_daee FROM tb_manancial_lancamento RIGHT JOIN (tb_bacia_hidrografica RIGHT JOIN (tb_tipo_ete RIGHT JOIN (((tb_convenio RIGHT JOIN (tb_responsavel AS tb_responsavel_4 RIGHT JOIN (tb_responsavel AS tb_responsavel_3 RIGHT JOIN (tb_responsavel AS tb_responsavel_2 RIGHT JOIN (tb_responsavel AS tb_responsavel_1 RIGHT JOIN (tb_depto RIGHT JOIN (tb_tipo_empreendimento RIGHT JOIN (tb_responsavel RIGHT JOIN (tb_predio RIGHT JOIN tb_PI ON tb_predio.cod_predio = tb_PI.cod_predio) ON tb_responsavel.cod_fiscal = tb_PI.cod_fiscal) ON tb_tipo_empreendimento.id = tb_PI.cod_tipo_empreendimento) ON tb_depto.cod_depto = tb_PI.cod_programa) ON tb_responsavel_1.cod_fiscal = tb_PI.cod_engenheiro_daee) ON tb_responsavel_2.cod_fiscal = tb_PI.cod_engenheiro_plan_consorcio) ON tb_responsavel_3.cod_fiscal = tb_PI.cod_fiscal_consorcio) ON tb_responsavel_4.cod_fiscal = tb_PI.cod_engenheiro_construtora) ON tb_convenio.id = tb_PI.cod_convênio) LEFT JOIN tb_situacao_pi ON tb_PI.cod_situacao = tb_situacao_pi.cod_situacao) LEFT JOIN tb_situacao_pi AS tb_situacao_pi_1 ON tb_PI.cod_situacao_externa = tb_situacao_pi_1.cod_situacao) ON tb_tipo_ete.id = tb_PI.cod_tipo_ete) ON tb_bacia_hidrografica.id = tb_PI.cod_bacia_hidrografica) ON tb_manancial_lancamento.id = tb_PI.cod_manancial_lancamento ORDER BY tb_predio.Município, tb_PI.nome_empreendimento;"

				Set rs_lista = Server.CreateObject("ADODB.Recordset")
					rs_lista.CursorLocation = 3
					rs_lista.CursorType = 3
					rs_lista.LockType = 1
					rs_lista.Open strQ, objCon, , , &H0001

				If Not rs_lista.EOF Then
					While Not rs_lista.EOF
			%>
			<tr>
				<td class="text-middle" style="max-width: 180px"><%=(rs_lista.Fields.Item("Município").Value)%></td>
				<td class="text-middle" style="max-width: 180px"><%=(rs_lista.Fields.Item("nome_empreendimento").Value)%></td>
				<td class="text-middle" style="max-width: 400px"><%=(rs_lista.Fields.Item("Descrição da Intervenção FDE").Value)%></td>					
				<td class="text-middle text-center" style="max-width: 150px;"><%=(rs_lista.Fields.Item("nme_bacia_hidrografica").Value)%></td>
				<td class="text-middle text-center" style="max-width: 120px;"><%=(rs_lista.Fields.Item("nme_manancial").Value)%></td>
				<td class="text-middle text-center" style="max-width: 190px;"><%=(rs_lista.Fields.Item("latitude_longitude").Value)%></td>
				<td class="text-middle text-center" style="max-width: 120px;"><%=(rs_lista.Fields.Item("qtd_metragem_coletor_tronco").Value)%></td>
				<td class="text-middle text-center" style="max-width: 115px;"><%=(rs_lista.Fields.Item("qtd_metragem_interceptor").Value)%></td>
				<td class="text-middle text-center" style="max-width: 150px;"><%=(rs_lista.Fields.Item("qtd_metragem_emissario_fluente_bruto").Value)%></td>
				<td class="text-middle text-center" style="max-width: 50px;"><%=(rs_lista.Fields.Item("qtd_eee").Value)%></td>
				<td class="text-middle text-center" style="max-width: 120px;"><%=(rs_lista.Fields.Item("qtd_metragem_linha_recalque").Value)%></td>
				<td class="text-middle text-center" style="max-width: 150px;"><%=(rs_lista.Fields.Item("qtd_metragem_emissario_efluente_tratado").Value)%></td>
				<td class="text-middle text-center" style="max-width: 190px;"><%=(rs_lista.Fields.Item("nme_tipo_ete").Value)%></td>
				<td class="text-middle text-center" style="max-width: 300px;"><%=(rs_lista.Fields.Item("dsc_estacao_tratamento").Value)%></td>
				<td class="text-middle" style="max-width: 300px;"><%=(rs_lista.Fields.Item("dsc_observacoes_obra").Value)%></td>
				<td class="text-middle" style="max-width: 300px;"><%=(rs_lista.Fields.Item("dsc_resultado_obtido").Value)%></td>
			</tr>
			<%
						rs_lista.MoveNext
					Wend
				End If
			%>
		</tbody>
	</table>

</body>
</html>