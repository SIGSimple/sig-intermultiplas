<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
Response.CharSet = "UTF-8"

Set objCon = Server.CreateObject("ADODB.Connection")
	objCon.Open MM_cpf_STRING

cod_acompanhamento 	= Request.QueryString("cod_acompanhamento")
cod_empreendimento 	= Request.QueryString("cod_empreendimento")
nome_municipio	 	= Request.QueryString("nome_municipio")

selectQuery = "SELECT tb_histograma.cod_recurso, tb_histograma.qtd_recurso, tb_histograma.dsc_observacoes FROM c_max_acompanhamento_histograma INNER JOIN tb_histograma ON [c_max_acompanhamento_histograma].[MáxDecod_acompanhamento] = tb_histograma.cod_acompanhamento WHERE c_max_acompanhamento_histograma.PI = '"& cod_empreendimento &"';"

Set selectRS = Server.CreateObject("ADODB.Recordset")
	selectRS.CursorLocation = 3
	selectRS.CursorType = 3
	selectRS.LockType = 1
	selectRS.Open selectQuery, objCon, , , &H0001

If selectRS.RecordCount > 0 Then
	arrData 	= selectRS.getRows()
	numColums 	= Ubound(arrData, 1)
	numRows 	= Ubound(arrData, 2)

	For rowCounter = 0 To numRows
		cod_recurso 		= arrData(0, rowCounter)
		qtd_recurso 		= arrData(1, rowCounter)
		dsc_observacoes 	= arrData(2, rowCounter)

		insertQuery	= "INSERT INTO tb_histograma (cod_acompanhamento, cod_recurso, qtd_recurso, dsc_observacoes) VALUES ("& cod_acompanhamento &","& cod_recurso &","& qtd_recurso &",'"& dsc_observacoes &"');"

		Set updateCommand = Server.CreateObject("ADODB.Command")
			updateCommand.ActiveConnection = MM_cpf_STRING
			updateCommand.CommandText = insertQuery
			updateCommand.Execute
			updateCommand.ActiveConnection.Close
	Next

	Response.Redirect("cad_histograma.asp?cod_acompanhamento="& cod_acompanhamento &"&nome_municipio="& nome_municipio)
End If

%>