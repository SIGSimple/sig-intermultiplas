<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="functions.asp" -->
<%
	Response.CharSet = "UTF-8"
	
	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	If Not IsEmpty(Request.Form) Then
		strQ = "SELECT * FROM tb_contrato_medicao_item Where 1 <> 1"

		Set rs_update = Server.CreateObject("ADODB.Recordset")
			rs_update.CursorLocation = 3
			rs_update.CursorType = 0
			rs_update.LockType = 3
			rs_update.Open strQ, objCon, , , &H0001
			rs_update.Addnew()
			
			' INÍCIO CAMPOS
			rs_update("cod_contrato") 		= Trim(Request.Form("cod_contrato"))
			rs_update("cod_medicao") 		= Trim(Request.Form("cod_medicao"))
			rs_update("cod_item") 			= Trim(Request.Form("cod_item"))
			rs_update("dta_medicao") 		= Trim(Request.Form("dta_medicao"))
			rs_update("vlr_medido") 		= Trim(Request.Form("vlr_medido"))
			' FIM CAMPOS
			
			rs_update.Update
	End If

	Dim rs
	Dim rs_numRows

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.ActiveConnection = MM_cpf_STRING
	rs.Source = "SELECT * FROM tb_contrato_medicao_item WHERE cod_medicao = " & Request.QueryString("cod_medicao")
	rs.CursorType = 0
	rs.CursorLocation = 2
	rs.LockType = 1
	rs.Open()

	rs_numRows = 0
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
				$(".datepicker").datepicker($.datepicker.regional["pt-BR"]);
				$(".datepicker").datepicker("option", "dateFormat", "mm/yy");
			});
		</script>
	</head>

	<body>
		<p align="center">
			<strong>
				<span class="style17">Medição de Itens do Contrato</span>
			</strong>
		</p>

		<form method="post" name="form1">
			<input type="hidden" name="cod_contrato" value="<%=(Request.QueryString("cod_contrato"))%>">
			<input type="hidden" name="cod_medicao" value="<%=(Request.QueryString("cod_medicao"))%>">
			<input type="hidden" name="dta_medicao" value="<%=(Request.QueryString("dta_medicao"))%>">
			<table align="center">
				<tr valign="baseline">
					<td align="center" nowrap bgcolor="#CCCCCC" colspan="2">
						<span class="style22">
							<strong>Nº Autos Contrato: </strong><%=(Request.QueryString("num_autos"))%>
						</span>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="center" nowrap bgcolor="#CCCCCC" colspan="2">
						<span class="style22">
							<strong>Medição: </strong><%=(Request.QueryString("mes_ano_medicao"))%>
						</span>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Item:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_item">
							<option value=""></option>
							<%
								cod_contrato = Request.QueryString("cod_contrato")
								strQ = "SELECT * FROM tb_contrato_item WHERE cod_contrato = "& cod_contrato &" ORDER BY dsc_item ASC"

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("dsc_item").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("id").Value)%>"><%=(rs_combo.Fields.Item("dsc_item").Value)%></option>
							<%
										End If
										rs_combo.MoveNext
									Wend
								End If
							%>
						</select>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Valor Medido:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="vlr_medido" value="" size="15">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">&nbsp;</td>
					<td bgcolor="#CCCCCC">
						<input type="submit" value="Salvar">
					</td>
				</tr>
			</table>
		</form>

		<div align="center">
			<table border="0">
				<tr bgcolor="#999999">
					<td>&nbsp;</td>
					<td align="center">
						<span class="style7">Item</span>
					</td>
					<td align="center">
						<span class="style7">Data</span>
					</td>
					<td align="center">
						<span class="style7">Valor Medido</span>
					</td>
				</tr>
				<%
					cod_medicao = Request.QueryString("cod_medicao")
					strQ = "SELECT * FROM c_lista_itens_medicao WHERE cod_medicao = " & cod_medicao & " ORDER BY dta_medicao ASC, dsc_item ASC"

					Set rs_lista = Server.CreateObject("ADODB.Recordset")
						rs_lista.CursorLocation = 3
						rs_lista.CursorType = 3
						rs_lista.LockType = 1
						rs_lista.Open strQ, objCon, , , &H0001

					If Not rs_lista.EOF Then
						While Not rs_lista.EOF
				%>
				<tr bgcolor="#CCCCCC">
					<td>
						<a href="delete_contrato_medicao_item.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>&cod_medicao=<%=(Request.QueryString("cod_medicao"))%>&dta_medicao=<%=(Request.QueryString("dta_medicao"))%>&mes_ano_medicao=<%=(Request.QueryString("mes_ano_medicao"))%>&cod_contrato=<%=(Request.QueryString("cod_contrato"))%>&num_autos=<%=(Request.QueryString("num_autos"))%>">
							<img src="const/imagens/delete.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dsc_item").Value)%></span>
					</td>
					<td align="center">
						<span class="style5">
							<%
								mes = CaptalizeText(MonthName(rs_lista.Fields.Item("mes_medicao").Value,True))
								ano = rs_lista.Fields.Item("ano_medicao").Value
								Response.Write mes & "/" & ano
							%>
						</span>
					</td>
					<td align="right">
						<span class="style5 vlr"><%=(rs_lista.Fields.Item("vlr_medido").Value)%></span>
					</td>
				</tr>
				<%
							rs_lista.MoveNext
						Wend
					End If
				%>
			</table>
		</div>
	</body>
</html>