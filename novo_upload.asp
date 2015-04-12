<!--#include file="Connections/cpf.asp" -->
<%
On Error Resume Next
 
Server.ScriptTimeOut = 90

Dim sDestinationPath, itemId, folderName, dscObservacoes

Set itemId = Request.QueryString("id")
Set folderName = Request.QueryString("folder")

sDestinationPath = Server.MapPath("/ARQUIVOS/" & folderName & "/")
If Not Right(sDestinationPath, 1) = "\" Then
	sDestinationPath = sDestinationPath & "\"
End If

Set objUpload = Server.CreateObject("Dundas.Upload.2")
objUpload.MaxFileSize = 4194304 '4MB'
objUpload.UseUniqueNames = False
objUpload.SaveToMemory

If Err.Number <> 0 Then
	Response.Redirect "erro_upload.asp"
Else
	dscObservacoes = objUpload.Form("dsc_observacoes")

	For Each objUploadedFile in objUpload.Files
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'						GRAVANDO NO DISCO							'
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		sFileName = objUpload.GetFileName(objUploadedFile.OriginalPath)
		objUploadedFile.SaveAs sDestinationPath & itemId & "_" & sFileName

		If InStr(1,objUploadedFile.ContentType,"octet-stream") Then
			objUploadedFile.Delete
		End If

		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'					GRAVANDO NO BANCO DE DADOS						'
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		Dim cmdInsert
		Dim sqlTable, sqlInsert, sqlFields, sqlValues

		Select Case folderName
			CASE "ACOMPANHAMENTO"
				sqlTable = "tb_acompanhamento_arquivo"
			Case "NOTA"
				sqlTable = "tb_convenio_aditamento_nota_arquivo"
			Case "ADITAMENTO"
				sqlTable = "tb_convenio_aditamento_arquivo"
			Case "CONVENIO"
				sqlTable = "tb_convenio_arquivo"
			Case "LICENCA"
				sqlTable = "tb_licenca_ambiental_arquivo"
			Case "OUTORGA"
				sqlTable = "tb_outorga_arquivo"
			Case "TCRA"
				sqlTable = "tb_tcra_arquivo"
			Case "APP"
				sqlTable = "tb_app_arquivo"
		End Select 

		sqlFields = "cod_referencia, nme_arquivo, pth_arquivo"
		sqlValues = itemId & ",'" & sFileName & "', '" & sDestinationPath & "'"

		If len(dscObservacoes) > 0 Then
			sqlFields = sqlFields + ", dsc_observacoes"
			sqlValues = sqlValues + ", '"& dscObservacoes &"'"
		End If

		sqlInsert = "INSERT INTO "& sqlTable &" ("& sqlFields &") VALUES ("& sqlValues &")"

		Set cmdInsert = Server.CreateObject("ADODB.Command")
			cmdInsert.ActiveConnection = MM_cpf_STRING
			cmdInsert.CommandText = sqlInsert
			cmdInsert.Execute
			cmdInsert.ActiveConnection.Close

		Set cmdInsert = Nothing
	Next
End If

Set objUpload = Nothing

Response.Redirect(Request.QueryString("retUrl"))

%>