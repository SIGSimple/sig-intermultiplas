<!--#include file="Connections/cpf.asp" -->
<%
	Response.CharSet = "UTF-8"

	If Not IsEmpty(Request.Form) Then
		strQuery 	= Request.Form("sql_query")
		urlRedirect = Request.Form("url_redirect")

		If IsEmpty(strQuery) Or IsNull(strQuery) Then
			strQuery = Request.QueryString("sql_query")
		End If

		If IsEmpty(urlRedirect) Or IsNull(urlRedirect) Then
			urlRedirect = Request.QueryString("url_redirect")
		End If

		Set updateCommand = Server.CreateObject("ADODB.Command")
		updateCommand.ActiveConnection = MM_cpf_STRING
		updateCommand.CommandText = strQuery
		updateCommand.Execute
		updateCommand.ActiveConnection.Close

		Response.Redirect(urlRedirect)
	End If
%>