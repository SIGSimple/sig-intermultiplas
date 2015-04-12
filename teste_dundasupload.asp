<!--#include file="Connections/cpf.asp" -->
<%
On Error Resume Next
 
Server.ScriptTimeOut = 90

Dim sDestinationPath, itemId, folderName

Set itemId = Request.QueryString("id")
Set folderName = Request.QueryString("folder")

sDestinationPath = Server.MapPath("/ARQUIVOS/" & folderName & "/")
If Not Right(sDestinationPath, 1) = "\" Then
	sDestinationPath = sDestinationPath & "\"
End If

Set objUpload = Server.CreateObject("Dundas.Upload.2")
objUpload.MaxFileSize = 2097152 '2MB'
objUpload.UseUniqueNames = False
objUpload.SaveToMemory

If Err.Number <> 0 Then
	Response.Redirect "erro_upload.asp"
Else
	For Each objUploadedFile in objUpload.Files
		sFileName = objUpload.GetFileName(objUploadedFile.OriginalPath)
		sFileName = itemId & "_" & sFileName
		objUploadedFile.SaveAs sDestinationPath & sFileName

		If InStr(1,objUploadedFile.ContentType,"octet-stream") Then
			objUploadedFile.Delete
		End If

		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'					GRAVANDO NO BANCO DE DADOS						'
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		Dim cmdInsert
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

		Set cmdInsert = Nothing
	Next
End If

Set objUpload = Nothing

Response.Redirect(Request.QueryString("retUrl"))

%>