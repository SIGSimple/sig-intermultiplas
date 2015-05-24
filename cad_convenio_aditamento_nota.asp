<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
	Response.CharSet = "UTF-8"
	
	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
  		objCon.Open MM_cpf_STRING

	If Not IsEmpty(Request.Form) Then
		strQ = "SELECT * FROM tb_convenio_aditamento_nota Where 1 <> 1"

		Set rs_update = Server.CreateObject("ADODB.Recordset")
			rs_update.CursorLocation = 3
			rs_update.CursorType = 0
			rs_update.LockType = 3
			rs_update.Open strQ, objCon, , , &H0001
			rs_update.Addnew()
			
			' INÍCIO CAMPOS
			rs_update("cod_convenio_aditamento")	= Trim(Request.Form("cod_convenio_aditamento"))
			rs_update("cod_usuario")				= Trim(Request.Form("cod_usuario"))
			rs_update("dsc_nota") 					= Trim(Request.Form("dsc_nota"))
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
	</head>

	<body>
		<p align="center">
			<strong>
				<span class="style17">Aditamentos de Convênio</span>
			</strong>
		</p>

		<form method="post" name="form1">
			<input type="hidden" name="cod_convenio_aditamento" value="<%=(Request.QueryString("cod_convenio_aditamento"))%>">
			<input type="hidden" name="cod_usuario" value="<%=(Session("MM_Userid"))%>">
			<table align="center">
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Núm. Termo:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<span class="style22"><%=(Request.QueryString("num_termo_aditamento"))%></span>
					</td>
				</tr>
				<tr valign="middle">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">
						<span class="style22">Nota:</span>
					</td>
					<td bgcolor="#CCCCCC">
						<textarea name="dsc_nota" cols="25"></textarea>
					</td>
				</tr>
				<tr valign="baseline">
					<td align="right" nowrap bgcolor="#CCCCCC" class="style7">&nbsp;</td>
					<td bgcolor="#CCCCCC">
						<input type="submit" value="Salvar" style="width: 100%;">
					</td>
				</tr>
			</table>
		</form>

		<div align="center">
			<table border="0">
				<tr bgcolor="#999999">
					<!-- <td>&nbsp;</td>
					<td>&nbsp;</td> -->
					<td>
						<span class="style7">Nota</span>
					</td>
					<td align="center">
						<span class="style7">Usuário</span>
					</td>
					<td align="center">
						<span class="style7">Data de Registro</span>
					</td>
					<td align="center">
						<span class="style7">Upload de Arquivos (Máx. 2MB)</span>
					</td>
				</tr>
				<%
					cod_convenio_aditamento = Request.QueryString("cod_convenio_aditamento")
					strQ = "SELECT * FROM tb_convenio_aditamento_nota INNER JOIN login ON login.idusuario = tb_convenio_aditamento_nota.cod_usuario WHERE tb_convenio_aditamento_nota.cod_convenio_aditamento = " & cod_convenio_aditamento

					Set rs_lista = Server.CreateObject("ADODB.Recordset")
						rs_lista.CursorLocation = 3
						rs_lista.CursorType = 3
						rs_lista.LockType = 1
						rs_lista.Open strQ, objCon, , , &H0001

					If Not rs_lista.EOF Then
						While Not rs_lista.EOF
				%>
				<tr bgcolor="#CCCCCC">
					<!-- <td>
						<a href="altera_convenio.asp?cod_convenio=cod_convenio">
							<img src="const/imagens/edit.gif" width="16" height="15" border="0" />
						</a>
					</td>
					<td>
						<a href="del_convenio.asp?cod_convenio=cod_convenio">
							<img src="const/imagens/delete.gif" width="16" height="15" border="0" />
						</a>
					</td> -->
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dsc_nota").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("nome").Value)%></span>
					</td>
					<td>
						<span class="style5"><%=(rs_lista.Fields.Item("dta_registro").Value)%></span>
					</td>
					<td>
						<form id="form-upload" method="post" enctype="multipart/form-data"
							action="novo_upload.asp?id=<%=(rs_lista.Fields.Item("id").Value)%>&folder=NOTA&retUrl=<%=(Request.ServerVariables("URL"))%>?<%=(Request.QueryString)%>">
							<input type="file" name="blob">
							<br/>
							<input type="submit" id="btnSubmit" value="Upload">
						</form>

						<%
							cod_convenio_aditamento_nota = rs_lista.Fields.Item("id").Value
							strF = "SELECT * FROM tb_convenio_aditamento_nota_arquivo WHERE cod_referencia = " & cod_convenio_aditamento_nota

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
									<a href="download.asp?path=/ARQUIVOS/NOTA&filename=<%=(rs_lista.Fields.Item("id").Value)%>_<%=(rs_files.Fields.Item("nme_arquivo").Value)%>">
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
		
		<script src="js/jquery-1.10.2.js"></script>
		<script src="js/ui/1.11.3/jquery-ui.js"></script>
		<script src="js/datepicker-pt-BR.js"></script>
		<script src="js/upload.js"></script>
		<script>
			$(function() {
				$(".datepicker").datepicker($.datepicker.regional["pt-BR"]);
			});
		</script>
	</body>
</html>