<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="JSON_2.0.4.asp"-->
<!--#include file="JSON_UTIL_0.1.1.asp"-->
<!--#include file="functions.asp" -->
<%

	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
		objCon.Open MM_cpf_STRING

	Dim sqlQuery
		sqlQuery = Request.Form("sql")

	If sqlQuery = "" Then
		sqlQuery = Request.QueryString("sql")
	End If

	QueryToJSON(objCon, sqlQuery).Flush
%>