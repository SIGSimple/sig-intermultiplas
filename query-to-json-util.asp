<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="JSON_2.0.4.asp"-->
<!--#include file="JSON_UTIL_0.1.1.asp"-->
<!--#include file="functions.asp" -->
<%
	Response.CharSet = "UTF-8"

	Set objCon = Server.CreateObject("ADODB.Connection")
		objCon.Open MM_cpf_STRING

	sqlQuery = Request.Form("sql")

	If sqlQuery = "" Then
		sqlQuery = Request.QueryString("sql")
	End If

	' Response.Write sqlQuery

	QueryToJSON(objCon, sqlQuery).Flush
%>