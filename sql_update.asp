<!--#include file="Connections/cpf.asp" -->
<%
	Response.CharSet = "UTF-8"

	If Not IsEmpty(Request.Form) Then
		strQuery = Request.Form("sql_query")

		Set updateCommand = Server.CreateObject("ADODB.Command")
	    updateCommand.ActiveConnection = MM_cpf_STRING
	    updateCommand.CommandText = strQuery
	    updateCommand.Execute
	    updateCommand.ActiveConnection.Close

	    Response.Redirect Request.Form("url_redirect")
	End If
%>