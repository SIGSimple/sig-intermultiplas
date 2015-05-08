<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="functions.asp" -->
<%
	Response.CharSet = "UTF-8"
	
	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	If Not IsEmpty(Request.Form) Then
		strQ = "SELECT * FROM tb_contrato_pagamento Where 1 <> 1"

		Set rs_update = Server.CreateObject("ADODB.Recordset")
			rs_update.CursorLocation = 3
			rs_update.CursorType = 0
			rs_update.LockType = 3
			rs_update.Open strQ, objCon, , , &H0001
			rs_update.Addnew()
			
			' INÍCIO CAMPOS
			rs_update("cod_contrato") 		= Trim(Request.Form("cod_contrato"))
			rs_update("cod_medicao") 		= Trim(Request.Form("cod_medicao"))
			rs_update("dta_pagamento") 		= Trim(Request.Form("dta_pagamento"))
			rs_update("vlr_pagamento") 		= Trim(Request.Form("vlr_pagamento"))
			' FIM CAMPOS

			rs_update.Update

			' append the query string to the redirect URL
			redirectUrl = "cad_contrato_medicao.asp"
			If (redirectUrl <> "" And Request.QueryString <> "") Then
				If (InStr(1, redirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
					redirectUrl = redirectUrl & "?" & Request.QueryString
				Else
					redirectUrl = redirectUrl & "&" & Request.QueryString
				End If
			End If

			If (redirectUrl <> "") Then
				Response.Redirect(redirectUrl)
			End If
	End If

	Dim rs
	Dim rs_numRows

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.ActiveConnection = MM_cpf_STRING
	rs.Source = "SELECT * FROM tb_contrato_pagamento WHERE cod_medicao = " & Request.QueryString("cod_medicao")
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
				<span class="style17">Registro de Pagamento de Medição de Contrato</span>
			</strong>
		</p>

		<form method="post" name="form1">
			<input type="hidden" name="cod_contrato" value="<%=(Request.QueryString("cod_contrato"))%>">
			<input type="hidden" name="cod_medicao" value="<%=(Request.QueryString("cod_medicao"))%>">
			<input type="hidden" name="dta_pagamento" value="<%=(Request.QueryString("dta_medicao"))%>">
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
						<span class="style22">Valor do Pagamento:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="vlr_pagamento" value="" size="15">
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
	</body>
</html>