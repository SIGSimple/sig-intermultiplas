<!--#include file="Connections/cpf.asp" -->
<%
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'			REMOVE O ARQUIVO DO BANCO DE DADOS				'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	Dim returnurl, fileid, foldername, table, sql, cmd

	Set returnurl 	= Request.QueryString("returnurl")
	Set fileid 		= Request.QueryString("fileid")
	Set foldername 	= Request.QueryString("foldername")

	Select Case foldername
		Case "ACOMPANHAMENTO"
			table = "tb_acompanhamento_arquivo"
		Case "NOTA"
			table = "tb_convenio_aditamento_nota_arquivo"
		Case "ADITAMENTO"
			table = "tb_convenio_aditamento_arquivo"
		Case "CONVENIO"
			table = "tb_convenio_arquivo"
		Case "LICENCA"
			table = "tb_licenca_ambiental_arquivo"
		Case "OUTORGA"
			table = "tb_outorga_arquivo"
		Case "TCRA"
			table = "tb_tcra_arquivo"
		Case "APP"
			table = "tb_app_arquivo"
		Case "CONTRATO"
			table = "tb_contrato_arquivo"
		Case "LICITACAO"
			table = "tb_licitacao_arquivo"
	End Select

	sql = "DELETE FROM " & table & " WHERE id_arquivo = " & fileid

	Set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = MM_cpf_STRING
		cmd.CommandText = sql
		cmd.Execute
		cmd.ActiveConnection.Close

	Set cmd = Nothing

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'			REMOVE O ARQUIVO DO SISTEMA DE ARQUIVOS			'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	Dim fs, filename

	Set filename = Request.QueryString("filename")

	fullfilename = "/ARQUIVOS/" & foldername & "/" & filename
	filepath = Server.MapPath(fullfilename)

	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	
	If fs.FileExists(filepath) Then
		fs.DeleteFile(filepath)
	End If

	Set fs = Nothing

	Response.Redirect(returnurl)
%>