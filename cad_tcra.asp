<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
	Response.CharSet = "UTF-8"
	
	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	If Not IsEmpty(Request.Form) Then
		strQ = "SELECT * FROM tb_tcra Where 1 <> 1"

		Set rs_update = Server.CreateObject("ADODB.Recordset")
			rs_update.CursorLocation = 3
			rs_update.CursorType = 0
			rs_update.LockType = 3
			rs_update.Open strQ, objCon, , , &H0001
			rs_update.Addnew()
			
			' INÍCIO CAMPOS
			rs_update("cod_empreendimento") = Trim(Request.Form("cod_empreendimento"))
			rs_update("cod_tcra") 			= Trim(Request.Form("cod_tcra"))
			
			If(Request.Form("dta_concessao") <> "") Then
				rs_update("dta_concessao") 	= Request.Form("dta_concessao")
			End If
			
			If(Request.Form("dta_vencimento") <> "") Then
				rs_update("dta_vencimento") = Request.Form("dta_vencimento")
			End If

			rs_update("dsc_observacoes") 	= Trim(Request.Form("dsc_observacoes"))
			rs_update("flg_receber_notificacoes") 	= Request.Form("flg_receber_notificacoes")
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
				<span class="style17">Cadastro de TCRAs</span>
			</strong>
		</p>

		<form method="post" name="form1">
			<input type="hidden" name="cod_empreendimento" value="<%=(Request.QueryString("cod_empreendimento"))%>" />
			<%
				Dim rs_dados_empreendimento

				strQ = "SELECT * FROM c_lista_dados_obras WHERE PI = '"& Request.QueryString("cod_empreendimento") &"'"

				Set rs_dados_empreendimento = Server.CreateObject("ADODB.Recordset")
				rs_dados_empreendimento.CursorLocation = 3
				rs_dados_empreendimento.CursorType = 3
				rs_dados_empreendimento.LockType = 1
				rs_dados_empreendimento.Open strQ, objCon, , , &H0001
			%>
			<table align="center">
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Município:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<%=(rs_dados_empreendimento.Fields.Item("municipio").Value)%>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Empreendimento:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<%=(rs_dados_empreendimento.Fields.Item("PI").Value)%> - <%=(rs_dados_empreendimento.Fields.Item("nome_empreendimento").Value)%>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Código:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" name="cod_tcra" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data de Concessão:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_concessao" value="" size="10">
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Data de Vencimento:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input type="text" class="datepicker" name="dta_vencimento" value="" size="10">
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
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Receber Notificações?</span>
					</td>
					<td bgcolor="#CCCCCC">
						<input name="flg_receber_notificacoes" type="checkbox" value="1" />
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
					<td>&nbsp;</td>
					<td>
						<span class="style7">Código</span>
					</td>
					<td>
						<span class="style7">Data de Concessão</span>
					</td>
					<td>
						<span class="style7">Data de Vencimento</span>
					</td>
					<td>
						<span class="style7">Observações</span>
					</td>
					<td>
						<span class="style7">Receber Notificações?</span>
					</td>
					<td>&nbsp;</td>
				</tr>
				<%
					cod_empreendimento = Request.QueryString("cod_empreendimento")
					strQ = "SELECT * from tb_tcra WHERE cod_empreendimento = " & cod_empreendimento

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
						<a href="altera_tcra.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>&nme_municipio=<%=(Request.QueryString("nme_municipio"))%>&cod_empreendimento=<%=(Request.QueryString("cod_empreendimento"))%>&nme_empreendimento=<%=(Request.QueryString("nme_empreendimento"))%>">
							<img src="const/imagens/edit.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<a href="delete_tcra.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>&nme_municipio=<%=(Request.QueryString("nme_municipio"))%>&cod_empreendimento=<%=(Request.QueryString("cod_empreendimento"))%>&nme_empreendimento=<%=(Request.QueryString("nme_empreendimento"))%>">
							<img src="const/imagens/delete.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("cod_tcra").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_concessao").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_vencimento").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dsc_observacoes").Value)%></span>
					</td>
					<td align="center">
						<span class="style5">
							<%
								If rs_lista.Fields.Item("flg_receber_notificacoes").Value Then
									Response.Write "Sim"
								Else
									Response.Write "Não"
								End If
							%>
						</span>
					</td>
					<td>
						<form id="form-upload" method="post" enctype="multipart/form-data"
							action="novo_upload.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>&folder=TCRA&retUrl=<%=(Request.ServerVariables("URL"))%>?<%=(Request.QueryString)%>">
							<input type="file" name="blob">
							<br/>
							<input type="submit" id="btnSubmit" value="Upload">
						</form>

						<%
							cod_convenio = rs_lista.Fields.Item("id").Value
							strF = "SELECT * FROM tb_tcra_arquivo WHERE cod_referencia = " & cod_convenio

							Set rs_files = Server.CreateObject("ADODB.Recordset")
								rs_files.CursorLocation = 3
								rs_files.CursorType = 3
								rs_files.LockType = 1
								rs_files.Open strF, objCon, , , &H0001

							If Not rs_files.EOF Then
								While Not rs_files.EOF
						%>
							<ul>
								<li>
									<a href="download.asp?path=/ARQUIVOS/TCRA&filename=<%=(rs_lista.Fields.Item("id").Value)%>_<%=(rs_files.Fields.Item("nme_arquivo").Value)%>">
										<%=(rs_files.Fields.Item("nme_arquivo").Value)%>
									</a>
								</li>
							</ul>
						<%
									rs_files.MoveNext
								Wend
							End If
						%>
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