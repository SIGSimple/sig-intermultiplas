upload<%@ ENABLESESSIONSTATE = False %>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="upload_fun.asp"-->
<%
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = -1
	Response.Buffer = True
	Response.AddHeader "Pragma", "no-cache"
	Response.AddHeader "Content-language", "pt-BR"
	Response.AddHeader "Content-Type", "text/html; charset=ISO-8859-1"

	'Response.AddHeader "cache-control", "no-store"
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'				FAZENDO UPLOAD DO ARQUIVO				'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	' Author Philippe Collignon
	' Email PhCollignon@email.com
	Dim UploadRequest
	Dim byteCount, RequestBin
	Dim sFullFilePath, sDestinationPath, sPathEnd
	Dim sContentType, sFilePathName, sFileName, sValue
	Dim oFile, oFSO
	Dim i
	Dim itemId, folderName

	Set itemId = Request.QueryString("id")
	Set folderName = Request.QueryString("folder")

	Response.Expires = 0
	Response.Buffer = True

	byteCount = Request.TotalBytes

	If( CInt(byteCount / 1024) > 400 ) Then
		Response.Write "Arquivo maior que o limite esperado (4MB)"
		Response.Flush
	End If

	RequestBin = Request.BinaryRead(byteCount)

	Set UploadRequest = CreateObject("Scripting.Dictionary")

	BuildUploadRequest RequestBin

	' This will place the uploaded file into the root directory of the web site - 
	' Modify this path as needed.
	sDestinationPath = Server.MapPath("/ARQUIVOS/" & folderName & "/")
	If Not Right(sDestinationPath, 1) = "\" Then
		sDestinationPath = sDestinationPath & "\"
	End If

	sFilePathName 	= UploadRequest.Item("blob").Item("FileName")
	sFileName 		= itemId & "_" & Right(sFilePathName,Len(sFilePathName)-InstrRev(sFilePathName,"\"))
	sValue 			= UploadRequest.Item("blob").Item("Value")
	sFullFilePath 	= sDestinationPath & sFileName

	'Create FileSytemObject Component
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set oFile = oFSO.CreateTextFile(sFullFilePath, True) 'overwrite existing file'
	 
	For i = 1 to LenB(sValue)
	    oFile.Write chr(AscB(MidB(sValue,i,1)))
	Next
	 
	oFile.Close

	Set oFile = Nothing
	Set oFSO = Nothing
	Set UploadRequest = Nothing

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'				GRAVANDO NO BANCO DE DADOS				'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	Dim sqlTable, sqlInsert, sqlFields, sqlValues

	Select Case folderName
		Case "NOTA"
			sqlTable = "tb_convenio_aditamento_nota_arquivo"
		Case "ADITAMENTO"
			sqlTable = "tb_convenio_aditamento_arquivo"
		Case "CONVENIO"
			sqlTable = "tb_convenio_arquivo"
	End Select 
	
	sqlFields = "cod_referencia, nme_arquivo, pth_arquivo"
	sqlValues = itemId & ",'" & sFileName & "', '" & sDestinationPath & "'"
	sqlInsert = "INSERT INTO "& sqlTable &" ("& sqlFields &") VALUES ("& sqlValues &")"

	Set cmdInsert = Server.CreateObject("ADODB.Command")
		cmdInsert.ActiveConnection = MM_cpf_STRING
		cmdInsert.CommandText = sqlInsert
		cmdInsert.Execute
		cmdInsert.ActiveConnection.Close

	Response.Redirect(Request.QueryString("retUrl"))
	Response.Flush
%>