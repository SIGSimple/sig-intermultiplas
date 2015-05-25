<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
	Response.CharSet = "UTF-8"
	
	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	If Not IsEmpty(Request.Form) Then
		strQ = "SELECT * FROM tb_histograma Where 1 <> 1"

		Set rs_update = Server.CreateObject("ADODB.Recordset")
			rs_update.CursorLocation = 3
			rs_update.CursorType = 0
			rs_update.LockType = 3
			rs_update.Open strQ, objCon, , , &H0001
			rs_update.Addnew()
			
			' INÍCIO CAMPOS
			rs_update("cod_acompanhamento") = Trim(Request.Form("cod_acompanhamento"))
			rs_update("cod_recurso") 		= Trim(Request.Form("cod_recurso"))
			rs_update("qtd_recurso") 		= Trim(Request.Form("qtd_recurso"))
			rs_update("dsc_observacoes") 	= Trim(Request.Form("dsc_observacoes"))
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
		<script type="text/javascript">
			$(function() {
				$(".datepicker").datepicker($.datepicker.regional["pt-BR"]);
			});
		</script>
	</head>

	<body>
		<p align="center">
			<strong>
				<span class="style17">Histograma de Recursos</span>
			</strong>
		</p>

		<h4 align="center">Dados do Acompanhamento</h4>

		<table align="center">
			<%
				cod_acompanhamento = Request.QueryString("cod_acompanhamento")
				strQ = "SELECT tb_Acompanhamento.*, tb_responsavel.Responsável, tb_pi.PI, tb_pi.nome_empreendimento FROM tb_pi RIGHT JOIN (tb_Acompanhamento LEFT JOIN tb_responsavel ON tb_Acompanhamento.cod_fiscal = tb_responsavel.cod_fiscal) ON tb_pi.PI = tb_Acompanhamento.PI WHERE (((tb_Acompanhamento.[cod_acompanhamento])="& cod_acompanhamento &"));"

				Set rs_data = Server.CreateObject("ADODB.Recordset")
					rs_data.CursorLocation = 3
					rs_data.CursorType = 3
					rs_data.LockType = 1
					rs_data.Open strQ, objCon, , , &H0001

				If Not rs_data.EOF Then
					While Not rs_data.EOF
						cod_empreendimento = rs_data.Fields.Item("tb_pi.PI").Value
			%>
			<tr valign="baseline">
				<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
					<span class="style22">Município:</span>
				</td>
				<td bgcolor="#CCCCCC">
					<%=(Request.QueryString("nome_municipio"))%>
				</td>
			</tr>
			<tr valign="baseline">
				<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
					<span class="style22">Empreendimento:</span>
				</td>
				<td bgcolor="#CCCCCC">
					<%=(rs_data.Fields.Item("nome_empreendimento").Value)%>
				</td>
			</tr>
			<tr valign="baseline">
				<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
					<span class="style22">Data do Registro:</span>
				</td>
				<td bgcolor="#CCCCCC">
					<%=(rs_data.Fields.Item("Data do Registro").Value)%>
				</td>
			</tr>
			<tr valign="baseline">
				<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
					<span class="style22">Responsável:</span>
				</td>
				<td bgcolor="#CCCCCC">
					<%=(rs_data.Fields.Item("Responsável").Value)%>
				</td>
			</tr>
			<%
						rs_data.MoveNext
					Wend
				End If
			%>
		</table>

		<h4 align="center">Novo registro</h4>

		<form method="post" name="form1">
			<input type="hidden" name="cod_acompanhamento" value="<%=(Request.QueryString("cod_acompanhamento"))%>"/>
			<table align="center">
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Recurso:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<select name="cod_recurso">
							<option value=""></option>
							<%
								strQ = "SELECT * FROM tb_recurso "

								Set rs_combo = Server.CreateObject("ADODB.Recordset")
									rs_combo.CursorLocation = 3
									rs_combo.CursorType = 3
									rs_combo.LockType = 1
									rs_combo.Open strQ, objCon, , , &H0001

								If Not rs_combo.EOF Then
									While Not rs_combo.EOF
										If Trim(rs_combo.Fields.Item("nme_recurso").Value) <> "" Then
							%>
							<option value="<%=(rs_combo.Fields.Item("id").Value)%>"><%=(rs_combo.Fields.Item("nme_recurso").Value)%></option>
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
						<span class="style22">Quantidade:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="qtd_recurso" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Observações:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<textarea name="dsc_observacoes" cols="25"></textarea>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">&nbsp;</td>
					<td bgcolor="#CCCCCC">
						<input type="submit" value="Salvar">
						 ou 
						<a href="rep_histograma.asp?cod_empreendimento=<%=(cod_empreendimento)%>&nome_municipio=<%=(Request.QueryString("nome_municipio"))%>&cod_acompanhamento=<%=(Request.QueryString("cod_acompanhamento"))%>">
							Replicar Último Histograma
						</a>
					</td>
				</tr>
			</table>
		</form>

		<div align="center">
			<table border="0">
				<tr bgcolor="#999999">
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>
						<span class="style7">Recurso</span>
					</td>
					<td>
						<span class="style7">Tipo de Recurso</span>
					</td>
					<td>
						<span class="style7">Quantidade</span>
					</td>
					<td>
						<span class="style7">Observações</span>
					</td>
				</tr>
				<%
					cod_acompanhamento = Request.QueryString("cod_acompanhamento")
					strQ = "SELECT tb_histograma.*, tb_tipo_recurso.nme_tipo_recurso, tb_recurso.nme_recurso FROM (tb_histograma INNER JOIN tb_recurso ON tb_histograma.cod_recurso = tb_recurso.id) INNER JOIN tb_tipo_recurso ON tb_recurso.cod_tipo_recurso = tb_tipo_recurso.id where cod_acompanhamento = " & cod_acompanhamento

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
						<a href="altera_histograma.asp?id=<%=(rs_lista.Fields.Item("id"))%>&<%=(Request.QueryString)%>">
							<img src="const/imagens/edit.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<a href="delete_histograma.asp?id=<%=(rs_lista.Fields.Item("id"))%>&<%=(Request.QueryString)%>">
							<img src="const/imagens/delete.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("nme_recurso").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("nme_tipo_recurso").Value)%></span>
					</td>
					<td align="center">
						<span class="style5"><%=(rs_lista.Fields.Item("qtd_recurso").Value)%></span>
					</td>
					<td align="center">
						<span class="style5"><%=(rs_lista.Fields.Item("dsc_observacoes").Value)%></span>
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