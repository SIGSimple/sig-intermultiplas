<!--#include file="Connections/cpf.asp" -->
<%
	Set conn = Server.CreateObject("ADODB.Connection")
		conn.open MM_cpf_STRING

	Response.Write "Banco conectado!"

	conn.close()

	Set conn = Nothing
%>