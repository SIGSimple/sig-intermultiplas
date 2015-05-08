<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
	Response.CharSet = "UTF-8"
	
	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	If Not IsEmpty(Request.Form) Then
		strQ = "SELECT * FROM tb_contrato_cronograma Where 1 <> 1"

		Set rs_update = Server.CreateObject("ADODB.Recordset")
			rs_update.CursorLocation = 3
			rs_update.CursorType = 0
			rs_update.LockType = 3
			rs_update.Open strQ, objCon, , , &H0001
			rs_update.Addnew()
			
			' INÍCIO CAMPOS
			rs_update("cod_contrato") 	= Trim(Request.Form("cod_contrato"))
			' FIM CAMPOS
			
			rs_update.Update
	End If
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
		<title>Untitled Document</title>
		<style type="text/css">
			<!--
				.style5 {font-family: Arial, Helvetica, sans-serif; font-size: 12px;}
				.style7 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold;}
				.style17 {font-family: Arial, Helvetica, sans-serif; font-size: 16px;}
				.style22 {font-family: Arial, Helvetica, sans-serif; font-size: 9;}
				.style23 {font-size: 9}
			-->
		</style>
		<link rel="stylesheet" href="//code.jquery.com/ui/1.11.3/themes/smoothness/jquery-ui.css">
		<script src="//code.jquery.com/jquery-1.11.2.min.js"></script>
		<script type="text/javascript" src="//code.jquery.com/ui/1.11.4/jquery-ui.min.js"></script>
		<script type="text/javascript" src="js/datepicker-pt-BR.js"></script>
		<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.9.0/moment.min.js"></script>
		<script type="text/javascript" src="js/jquery.number.min.js"></script>
		<script type="text/javascript" src="js/moment.min.js"></script>
		<script type="text/javascript" src="js/underscore-min.js"></script>
		<script type="text/javascript">
			function adjustVlrLayout() {
				$.each($(".vlr"), function(i, item){
					// $(item).val($.number($(item).val(), 0, ",", "."));
					if($(item).text() != "")
						$(item).text("R$ " + $.number($(item).text(), 2, ",", "."));
				});
			}

			$(function(){
				adjustVlrLayout();

				var arr_cronograma_data_values = [];

				$.each($(".tr-values"), function(i, item){
					arr_cronograma_data_values.push($(this).data());
				});

				$.each(arr_cronograma_data_values, function(i, item) {
					var dta_inicio_atual 	= moment(item.dtaInicio, "DD/MM/YYYY");
					var dta_termino_atual 	= moment(item.dtaTermino, "DD/MM/YYYY");
					
					arr_cronograma_data_values[i].dtaInicio = dta_inicio_atual;
					arr_cronograma_data_values[i].dtaTermino = dta_termino_atual;
					arr_cronograma_data_values[i].codCronogramaAnterior = 0;
					arr_cronograma_data_values[i].qtdMesesAnterior 		= 0;
					arr_cronograma_data_values[i].vlrTotalAnterior 		= 0;
					arr_cronograma_data_values[i].qtdMesesAtual 		= dta_termino_atual.diff(dta_inicio_atual, 'months');

					if(i > 0) {
						var dta_inicio_anterior 	= moment(arr_cronograma_data_values[i-1].dtaInicio, "DD/MM/YYYY");
						var dta_termino_anterior 	= moment(arr_cronograma_data_values[i-1].dtaTermino, "DD/MM/YYYY");

						arr_cronograma_data_values[i].codCronogramaAnterior = arr_cronograma_data_values[i-1].codCronograma;
						arr_cronograma_data_values[i].qtdMesesAnterior 		= dta_termino_anterior.diff(dta_inicio_anterior, 'months');
						arr_cronograma_data_values[i].vlrTotalAnterior 		= parseFloat(arr_cronograma_data_values[i-1].vlrTotal);
					}
				});

				var cod_contrato = <%=(Request.QueryString("cod_contrato"))%>;

				$.ajax({
					url: "query-to-json-util.asp",
					method: "POST",
					data: {
						sql: 'SELECT * FROM c_lista_items_cronogramas_contrato WHERE cod_contrato = '+ cod_contrato
					},
					success: function(data, textStatus, jqXHR){
						data = JSON.parse(data);

						if(data.length > 0) {
							var arr_items_value = data;
							arr_items_value = _.groupBy(arr_items_value, 'id');

							var x = 0;
							$.each(arr_items_value, function(i, item){
								var codCronograma 		= parseInt(i,10);
								
								var dtaInicio 				= _.findWhere(arr_cronograma_data_values, {codCronograma: codCronograma}).dtaInicio;
								var dtaTermino 				= _.findWhere(arr_cronograma_data_values, {codCronograma: codCronograma}).dtaTermino;
								var qtdMesesAtual 			= _.findWhere(arr_cronograma_data_values, {codCronograma: codCronograma}).qtdMesesAtual;
								var vlrTotalAtual 			= _.findWhere(arr_cronograma_data_values, {codCronograma: codCronograma}).vlrTotal;

								var dsc_tipo_cronograma = "";
								
								var flg_aditamento_valor 	= false,
									flg_aditamento_prazo 	= false,
									flg_replanilhamento 	= false;

								if(x > 0){
									var qtdMesesAnterior 		= _.findWhere(arr_cronograma_data_values, {codCronograma: codCronograma}).qtdMesesAnterior;
									var vlrTotalAnterior 		= _.findWhere(arr_cronograma_data_values, {codCronograma: codCronograma}).vlrTotalAnterior;
									
									var codCronogramaAnterior	= _.findWhere(arr_cronograma_data_values, {codCronograma: codCronograma}).codCronogramaAnterior;

									if(vlrTotalAtual > vlrTotalAnterior){
										flg_aditamento_valor = true;
										dsc_tipo_cronograma = "Aditamento de Valor";
									}
									else {
										var itens = _.values(_.pick(arr_items_value, codCronogramaAnterior))[0];

										$.each(item, function(x, xitem){
											$.each(itens, function(y, yitem){
												if(
													(xitem.cod_item_planejamento == yitem.cod_item_planejamento) 
													&&
													(xitem.SomaDevlr_planejamento > yitem.SomaDevlr_planejamento)
												){
													flg_replanilhamento = true;
												}
											});
										});
									}

									if(qtdMesesAtual > qtdMesesAnterior){
										flg_aditamento_prazo = true;
										dsc_tipo_cronograma = "Aditamento de Prazo";
									}

									if(flg_replanilhamento)
										dsc_tipo_cronograma = "Replanilhamento";
									else {
										if(flg_aditamento_valor && flg_aditamento_prazo)
											dsc_tipo_cronograma = "Aditamento de Prazo e Valor";
										else if((!flg_aditamento_valor) && (!flg_aditamento_prazo))
											dsc_tipo_cronograma = "Redistribuição de Valores";
									}
								}
								else
									dsc_tipo_cronograma = "Cronograma Inicial";

								var dsc_vigencia = dtaInicio.format("MMM/YYYY") +" a "+ dtaTermino.format("MMM/YYYY");

								$("tr[data-cod-cronograma="+ codCronograma +"]").find("span.dsc_tipo_cronograma").text(dsc_tipo_cronograma);
								$("tr[data-cod-cronograma="+ codCronograma +"]").find("span.prz_total").text(dsc_vigencia);
								$("tr[data-cod-cronograma="+ codCronograma +"]").find("span.qtd_meses").text(qtdMesesAtual + " Meses");

								x++;
							});
						}
					},
					error: function(jqXHR, textStatus, errorThrown){
						console.log(jqXHR, textStatus, errorThrown);
					}
				});
			});
		</script>
	</head>

	<body>
		<p align="center">
			<strong>
				<span class="style17">Cadastro de Cronogramas do Contrato</span>
			</strong>
		</p>

		<form method="post" name="form1">
			<input type="hidden" name="cod_contrato" value="<%=(Request.QueryString("cod_contrato"))%>">
			<table align="center">
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Nº Autos - Contrato:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<span class="style22"><%=(Request.QueryString("num_autos"))%></span>
					</td>
				</tr>
				<tr valign="baseline">
					<td bgcolor="#CCCCCC" colspan="2" align="center">
						<input type="submit" value="Criar novo cronograma">
					</td>
				</tr>
			</table>
		</form>

		<div align="center">
			<table border="0">
				<tr bgcolor="#999999">
					<td>&nbsp;</td>
					<td align="center"><span class="style7">Nº Revisão</span></td>
					<td align="center"><span class="style7">Vigência</span></td>
					<td align="center"><span class="style7">Prazo Total</span></td>
					<td align="center"><span class="style7">Valor Planejado</span></td>
					<td align="center"><span class="style7">Tipo Cronograma</span></td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				<%
					cod_contrato = Request.QueryString("cod_contrato")
					strQ = "SELECT * FROM c_lista_cronogramas_contrato WHERE cod_contrato = " & cod_contrato & " ORDER BY id ASC"

					Set rs_lista = Server.CreateObject("ADODB.Recordset")
						rs_lista.CursorLocation = 3
						rs_lista.CursorType = 3
						rs_lista.LockType = 1
						rs_lista.Open strQ, objCon, , , &H0001

					If Not rs_lista.EOF Then
						i = 0

						While Not rs_lista.EOF
							cod_cronograma 	= rs_lista.Fields.Item("id").Value
							vlr_total 		= Replace(rs_lista.Fields.Item("vlr_total").Value, ",", ".")
							dta_inicio 		= rs_lista.Fields.Item("dta_inicio").Value
							dta_termino 	= rs_lista.Fields.Item("dta_termino").Value
				%>
				<tr class="tr-values" bgcolor="#CCCCCC" 
					data-cod-cronograma="<%=(cod_cronograma)%>"
					data-vlr-total="<%=(vlr_total)%>"
					data-dta-inicio="<%=(dta_inicio)%>"
					data-dta-termino="<%=(dta_termino)%>">
					<td>
						<a href="delete_contrato_cronograma.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>&cod_contrato=<%=(Request.QueryString("cod_contrato"))%>&num_autos=<%=(Request.QueryString("num_autos"))%>&num_revisao=<%=(i)%>">
							<img src="const/imagens/delete.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td align="center">
						<span class="style5"><%=(i)%></span>
					</td>
					<td align="center">
						<span class="style5 prz_total"></span>
					</td>
					<td align="center">
						<span class="style5 qtd_meses"></span>
					</td>
					<td>
						<span class="style5 vlr"><%=(vlr_total)%></span>
					</td>
					<td align="center">
						<span class="style5 dsc_tipo_cronograma"></span>
					</td>
					<td>
						<%
							If Not rs_lista.Fields.Item("flg_valido").Value Then
						%>
							<form method="post" action="sql_update.asp">
								<input type="hidden" name="sql_query" value="UPDATE tb_contrato_cronograma SET flg_valido = 1 WHERE id = <%=(rs_lista.Fields.Item("id").Value)%>">
								<input type="hidden" name="url_redirect" value="cad_contrato_cronograma.asp?<%=(Request.QueryString)%>">
								<button type="submit">Marcar como válido</button>
							</form>
						<%
							End If
						%>
					</td>
					<td>
						<a href="cad_contrato_cronograma_item.asp?cod_cronograma=<%=(rs_lista.Fields.Item("id").Value)%>&cod_contrato=<%=(Request.QueryString("cod_contrato"))%>&num_autos=<%=(Request.QueryString("num_autos"))%>&num_revisao=<%=(i)%>">
							<span class="style5">
								Planejamento
							</span>
						</a>
					</td>
				</tr>
				<%
							rs_lista.MoveNext
							i = i + 1
						Wend
					End If
				%>
			</table>
		</div>
	</body>
</html>