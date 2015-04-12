<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%

Dim folderName

Set folderName = Request.QueryString("folder")

Select Case folderName
	Case "NOTA"
		Response.Write("gravar na tabela tb_convenio_aditamento_nota_arquivo")
	Case "ADITAMENTO"
		Response.Write("gravar na tabela tb_convenio_aditamento_arquivo")
	Case "CONVENIO"
		Response.Write("gravar na tabela tb_convenio_arquivo")
End Select

%>