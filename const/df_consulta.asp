<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/cpf.asp" -->
<%

'*******************************************************************
' Página gerada pelo sistema Dataform 2 - http://www.dataform.com.br
'*******************************************************************
' Altere os valores das variáveis indicadas abaixo se necessário

strCon = "DBQ=C:\inetpub\wwwroot\original\ARQUIVOS\DADOS\bd_fde.mdb;Driver={Microsoft Access Driver (*.mdb)};"
'strCon = "DBQ=d:\dominios\terkoisolacao.com\www\DB\fde.mdb;Driver={Microsoft Access Driver (*.mdb)};"

'Número total de registros a serem exibidos por página
Const RegPorPag = 15

'Número de páginas a ser exibido no índice de paginação
VarPagMax = 10

'Cor da linha selecionada na tabela de registros
cor_linha_selecionada = "gainsboro"

'Nome da página de consulta
pagina_consulta = "df_consulta.asp"

'Nome da página de alteração
pagina_alteracao = "df_alteracao.asp"

'Nome da página de inclusão
pagina_inclusao = "df_inclusao.asp"

'Nome da página de login
pagina_login = "df_login.asp"

'*******************************************************************

%>

<HTML>
<HEAD>
<TITLE>Consultar Registros</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="copyright" content="Dataform">
<meta name="keywords" content="dataform, asp dataform, aspdataform, asp-dataform">
<meta name="robots" content="ALL">
<style type="text/css">
<!--
.texto_pagina
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
color: dimgray;
}

.tabela_registros
{
width: 100%;
background-color: white;
}

.titulos_registros
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
color: white;
background-color: gray;
}

.exibe_registros
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
width: 100%;
color: dimgray;
background-color: whitesmoke;
}

.tabela_paginacao
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
width: 100%;
color: gray;
border-top: 1px solid gainsboro;
background-color: gainsboro;
}

.links_paginacao
{
color: dimgray;
text-decoration: none;
}

.links_paginacao:hover
{
color: gray;
text-decoration: underline;
}
-->
</style>
<SCRIPT language="JavaScript">
<!--
function abre_foto(width, height, nome) {
	var top; var left;
	top = ( (screen.height/2) - (height/2) )
	left = ( (screen.width/2) - (width/2) )
	window.open('',nome,'width='+width+',height='+height+',scrollbars=yes,toolbar=no,location=no,status=no,menubar=no,resizable=no,left='+left+',top='+top);
}
function confirm_delete(form) {
	if (confirm("Tem certeza que deseja excluir o registro?")) {
	document[form].action = '<%=Request.ServerVariables("SCRIPT_NAME")%>';
	document[form].submit();
	}
}
//-->
</SCRIPT>
</HEAD>
<body class="texto_pagina">
Links: <a href="<%=pagina_consulta%>" class="texto_pagina">Página de Consulta</a> | <a href="<%=pagina_inclusao%>" class="texto_pagina">Página de Inclusão<hr size=1 color=gainsboro></a><br/>

<%
If Request.QueryString("PagAtual") = "" Then
	PagAtual = 1
	NumPagMax = VarPagMax
Else
	NumPagMax = CInt(Request.QueryString("NumPagMax"))
	PagAtual = CInt(Request.QueryString("PagAtual"))
	Select Case Request.QueryString("Submit")
		Case "Anterior" : PagAtual = PagAtual - 1
		Case "Proxima" : PagAtual = PagAtual + 1
		Case "Menos" : NumPagMax = NumPagMax - VarPagMax
		Case "Mais" : NumPagMax = NumPagMax + VarPagMax
		Case Else : PagAtual = CInt(Request.QueryString("Submit"))
	End Select
	If NumPagMax < PagAtual then
		NumPagMax = NumPagMax + VarPagMax
	End If
	If NumPagMax - (VarPagMax - 1) > PagAtual then
		NumPagMax = NumPagMax - VarPagMax
	End If
End If

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open MM_cpf_STRING

	If Session("admin") <> "" And Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR") Then
	If Request.Form("recordno") <> "" Then
		Set objRS_delete = Server.CreateObject("ADODB.Recordset")
		objRS_delete.CursorLocation = 3
		objRS_delete.CursorType = 0
		objRS_delete.LockType = 3

		strQ_delete = Request.Form("strQ")
		indice = Trim(Request.Form("indice"))
		If indice <> "" Then strQ_delete = " SELECT * FROM tb_Construtora WHERE " & indice

		objRS_delete.Open strQ_delete, objCon, , , &H0001
		If indice = "" Then objRS_delete.Move Request.Form("recordno") - 1
		If Not objRS_delete.EOF Then
			objRS_delete.Delete
			objRS_delete.UpdateBatch
		End IF

		objRS_delete.Close
		Set objRS_delete = Nothing
		Set strQ_delete = Nothing
	End If
	End If

Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.CursorLocation = 3
objRS.CursorType = 2
objRS.LockType = 1
objRS.CacheSize = RegPorPag
strQ = "SELECT * FROM tb_Construtora LEFT JOIN tb_Municipios ON tb_Municipios.cod_mun = tb_Construtora.cod_municipio_empresa "

If Trim(Request("string_busca")) <> "" Then
	If Trim(Request("campo_busca")) <> "" Then
		strQ = strQ & " Where " & Trim(Request("campo_busca")) & " LIKE '%" & Trim(Request("string_busca")) & "%'"
	Else
		strQ = strQ & " Where 1 <> 1"
		strQ = strQ & " Or Construtora LIKE '%" & Trim(Request("string_busca")) & "%'"
		strQ = strQ & " Or Endereço da Construtora LIKE '%" & Trim(Request("string_busca")) & "%'"
		strQ = strQ & " Or Engenheiro responsável LIKE '%" & Trim(Request("string_busca")) & "%'"
		strQ = strQ & " Or Fone da Construtora LIKE '%" & Trim(Request("string_busca")) & "%'"
		strQ = strQ & " Or Número do CREA LIKE '%" & Trim(Request("string_busca")) & "%'"
	End If
End If

If Trim(Request.QueryString("Ordem")) <> "" Then
	strQ = strQ & " ORDER BY " & Request.QueryString("Ordem")
End If
objRS.Open strQ, objCon, , , &H0001
objRS.PageSize = RegPorPag

Set objRS_indice = Server.CreateObject("ADODB.Recordset")
objRS_indice.CursorLocation = 2
objRS_indice.CursorType = 0
objRS_indice.LockType = 2
strQ_indice = "SELECT * FROM tb_Construtora WHERE 1 <> 1"
objRS_indice.Open strQ_indice, objCon, , , &H0001
indice = ""
For Each item In objRS_indice.Fields
	If item.properties("IsAutoIncrement") = True Then
		indice = item.name
		Exit For
	End If
Next
objRS_indice.Close
Set objRS_indice = Nothing
Set strQ_indice = Nothing

Set objRS.ActiveConnection = Nothing
objCon.Close
Set objCon = Nothing
%>

<b>Consultar Registros</b>
<br/>
Visualize os registros da  tabela abaixo:<br/>
<form name="form_busca" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>">
	Pesquizar por
	<input type="text" name="string_busca" value="<%=Request("string_busca")%>" class="texto_pagina">
	em
	<select name="campo_busca" class="texto_pagina">
		<option value="" selected>
		</option>
		<option value="Construtora" <% If Trim(Request("campo_busca")) = Trim("Construtora") Then : Response.Write "selected" : End If %>>
			Empresa
		</option>
		<option value="Endereço da Construtora" <% If Trim(Request("campo_busca")) = Trim("Endereço da Construtora") Then : Response.Write "selected" : End If %>>
			Endereço da Empresa
		</option>
		<option value="Engenheiro responsável" <% If Trim(Request("campo_busca")) = Trim("Engenheiro responsável") Then : Response.Write "selected" : End If %>>
			Engenheiro responsável
		</option>
		<option value="Fone da Construtora" <% If Trim(Request("campo_busca")) = Trim("Fone da Construtora") Then : Response.Write "selected" : End If %>>
			Fone da Empresa
		</option>
		<option value="Número do CREA" <% If Trim(Request("campo_busca")) = Trim("Número do CREA") Then : Response.Write "selected" : End If %>>
			Número do CREA
		</option>
	</select>
	<input type="submit" name="submit" value="ok" class=texto_pagina style="color: black">
</form>

<%
If Not(objRS.EOF) Then
	objRS.AbsolutePage = PagAtual
	TotPag = objRS.PageCount
%>

Foram encontrados <%= objRS.RecordCount%> registros
<br/>
<br/>

<table border="0" cellpadding="2" cellspacing="1" class="tabela_registros">
	<tr class="titulos_registros">
		<%
			If Session("admin") <> "" And Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR") Then
				Response.Write "<td align=""center"" style=""background-color: crimson; color: white"" width=""1%"" nowrap><b>Editar</b></TD>"
			End IF

			If Right(Request.QueryString("Ordem"), 3) = "asc" Then
				Ordem = "desc"
			Else
				Ordem = "asc"
			End IF
		%>

		<td style="cursor: hand;" valign="top nowrap" onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=Construtora+<%=Ordem%>', '_self')">
			<%If Left(Request.QueryString("Ordem"), 11) = "Construtora" Then : Response.Write "<img src=""imagens/ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
			<b>Empresa</b>
		</td>
		<td style="cursor: hand;" valign="top nowrap" onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=Endereço da Construtora+<%=Ordem%>', '_self')">
			<%If Left(Request.QueryString("Ordem"), 23) = "Endereço da Construtora" Then : Response.Write "<img src=""imagens/ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
			<b>Endereço da Empresa</b>
		</td>
		
		<td style="cursor: hand;" valign="top nowrap" onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=cep_empresa+<%=Ordem%>', '_self')">
			<%If Left(Request.QueryString("Ordem"), 23) = "cep_empresa" Then : Response.Write "<img src=""imagens/ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
			<b>CEP da Empresa</b>
		</td>
		<td style="cursor: hand;" valign="top nowrap" onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=cod_municipio_empresa+<%=Ordem%>', '_self')">
			<%If Left(Request.QueryString("Ordem"), 23) = "cod_municipio_empresa" Then : Response.Write "<img src=""imagens/ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
			<b>Município da Empresa</b>
		</td>
		<td style="cursor: hand;" valign="top nowrap" onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=email_empresa+<%=Ordem%>', '_self')">
			<%If Left(Request.QueryString("Ordem"), 23) = "email_empresa" Then : Response.Write "<img src=""imagens/ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
			<b>Email da Empresa</b>
		</td>
		<td style="cursor: hand;" valign="top nowrap" onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=site_empresa+<%=Ordem%>', '_self')">
			<%If Left(Request.QueryString("Ordem"), 23) = "site_empresa" Then : Response.Write "<img src=""imagens/ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
			<b>Site da Empresa</b>
		</td>
		<td style="cursor: hand;" valign="top nowrap" onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=cnpj_empresa+<%=Ordem%>', '_self')">
			<%If Left(Request.QueryString("Ordem"), 23) = "cnpj_empresa" Then : Response.Write "<img src=""imagens/ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
			<b>CNPJ da Empresa</b>
		</td>
		<td style="cursor: hand;" valign="top nowrap" onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=Fone da Construtora+<%=Ordem%>', '_self')">
			<%If Left(Request.QueryString("Ordem"), 19) = "Fone da Construtora" Then : Response.Write "<img src=""imagens/ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
			<b>Fone da Empresa</b>
		</td>
		<td style="cursor: hand;" valign="top nowrap" onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=Engenheiro responsável+<%=Ordem%>', '_self')">
			<%If Left(Request.QueryString("Ordem"), 22) = "Engenheiro responsável" Then : Response.Write "<img src=""imagens/ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
			<b>Engenheiro responsável</b>
		</td>
		<td style="cursor: hand;" valign="top nowrap" onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=Número do CREA+<%=Ordem%>', '_self')">
			<%If Left(Request.QueryString("Ordem"), 14) = "Número do CREA" Then : Response.Write "<img src=""imagens/ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
			<b>Número do CREA</b>
		</td>
		<td style="cursor: hand;" valign="top nowrap" onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=telefone_responsavel+<%=Ordem%>', '_self')">
			<%If Left(Request.QueryString("Ordem"), 22) = "telefone_responsavel" Then : Response.Write "<img src=""imagens/ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
			<b>Telefone responsável</b>
		</td>
		<td style="cursor: hand;" valign="top nowrap" onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=email_responsavel+<%=Ordem%>', '_self')">
			<%If Left(Request.QueryString("Ordem"), 14) = "email_responsavel" Then : Response.Write "<img src=""imagens/ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
			<b>Email responsável</b>
		</td>
	</tr>

	<%
		For Cont = 1 to objRS.PageSize
	%>

	<tr class="exibe_registros" onMouseOver="this.style.backgroundColor='<%=cor_linha_selecionada%>';" onMouseOut="this.style.backgroundColor='';">

	<%
		If Session("admin") <> "" And Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR") Then
			Response.Write "<form name=""form_edit_" & Cont & """ action=""" & pagina_alteracao & """ method=post>"
			Response.Write "<td  align=""center"" nowrap style=""background-color: gainsboro""  nowrap>&nbsp;"

			If indice <> "" Then Response.Write "<input type=""hidden"" name=""indice"" value=""" & indice & "=" & objRS.Fields.Item(indice).Value & """>"
			
			Response.Write "<input type='hidden' name='recordno' value=""" & (objRS.AbsolutePosition) & """>"
			Response.Write "<input type='hidden' name='strQ' value=""" & strQ & """>"
			Response.Write "<input type='image' src=""imagens/edit.gif"" alt=""Alterar Registro"" name=alterar value=alterar>"
			
			If Session("admin") <> "" And Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR") Then
				Response.Write "&nbsp;<img src=""imagens/delete.gif"" alt=""Excluir Registro"" name='delete' border='0' style=""cursor:hand"" OnClick=""confirm_delete('form_edit_" & Cont & "')"">"
			End If
			
			Response.Write "&nbsp;</td>"
			Response.Write "</form>"
		End If
	%>

		<td><%=(objRS.Fields.Item("Construtora").Value)%></td>
		<td><%=(objRS.Fields.Item("Endereço da Construtora").Value)%></td>
		<td><%=(objRS.Fields.Item("cep_empresa").Value)%></td>
		<td><%=(objRS.Fields.Item("Municipios").Value)%></td>
		<td><%=(objRS.Fields.Item("email_empresa").Value)%></td>
		<td><%=(objRS.Fields.Item("site_empresa").Value)%></td>
		<td><%=(objRS.Fields.Item("cnpj_empresa").Value)%></td>
		<td><%=(objRS.Fields.Item("Fone da Construtora").Value)%></td>
		<td><%=(objRS.Fields.Item("Engenheiro responsável").Value)%></td>
		<td><%=(objRS.Fields.Item("Número do CREA").Value)%></td>
		<td><%=(objRS.Fields.Item("telefone_responsavel").Value)%></td>
		<td><%=(objRS.Fields.Item("email_responsavel").Value)%></td>
	</tr>

	<%
			objRS.MoveNext
			If objRS.Eof then Exit For
		Next
		Set Cont = Nothing
	%>
	<tr>
		<td colspan="12"><% LinksNavegacao() %></td>
	</tr>

</table>

<%
	If indice = "" Then
		Response.Write "<br/><b>ATENÇÃO:</b> Crie um campo do tipo <i>AutoIncrement</i> com qualquer nome em sua tabela para evitar erros na alteração dos dados. "
		Response.Write "<a href=""http://www.dataform.com.br/criar_campo_autoincrement.asp"" target=""_blank"">Clique aqui</a> para mais detalhes."
	End If
	objRS.Close
	Set objRS = Nothing
Else
%>

<br/>
<b>Nenhum registro foi encontrado</b>
<br/>
<br/>

<% End If %>

</body>
</html>

<%
Sub LinksNavegacao()
	'O código a seguir insere uma tabela com todos os links de navegação das páginas
	Response.Write "<table border=0 cellPadding=2 cellSpacing=0 class=tabela_paginacao>"
	Response.Write "<tr><td align='center' vAlign='top noWrap' colspan='5'>"
	Response.Write "Página " & PagAtual & " de " & TotPag
	Response.Write "</td></tr><tr><td width='33%' align='right' vAlign='top noWrap'>"
	If PagAtual > 1 Then
		Response.Write "<A href=""" & Request.ServerVariables("SCRIPT_NAME") & "?PagAtual=" &  PagAtual &"&VarPagMax=" & VarPagMax & "&NumPagMax=" & NumPagMax & "&Submit=Anterior&Ordem=" & Request.QueryString("Ordem")& "&string_busca=" & Server.URLEncode(Request("string_busca")) & "&campo_busca=" & Server.URLEncode(Request("campo_busca"))  & """ class=links_paginacao>&lt; Anterior</A>"
	End If
	Response.Write "</td><td width='33%' align='middle' vAlign='top noWrap'>"
	If NumPagMax - VarPagMax <> 0 then
		Response.Write "&nbsp;<A href=""" & Request.ServerVariables("SCRIPT_NAME") & "?PagAtual=" & NumPagMax - VarPagMax & "&VarPagMax=" & VarPagMax & "&NumPagMax=" & NumPagMax - VarPagMax & "&Submit=Menos&Ordem=" & Request.QueryString("Ordem") & "&string_busca=" & Server.URLEncode(Request("string_busca")) & "&campo_busca=" & Server.URLEncode(Request("campo_busca")) & """ class=links_paginacao>&lt;&lt;</A>&nbsp;&nbsp;"
	End If
	for i = NumPagMax - (VarPagMax - 1) to NumPagMax
		If i <= TotPag then
			If i <> CInt(PagAtual) then
				Response.Write "&nbsp;<A href=""" & Request.ServerVariables("SCRIPT_NAME") & "?PagAtual=" & PagAtual & "&VarPagMax=" & VarPagMax & "&NumPagMax=" & NumPagMax & "&Submit=" & i & "&Ordem=" & Request.QueryString("Ordem") & "&string_busca=" & Server.URLEncode(Request("string_busca")) & "&campo_busca=" & Server.URLEncode(Request("campo_busca")) & """ class=links_paginacao>" & i & "</A>&nbsp;"
			Else
				If PagAtual <> TotPag Then
					Response.Write "&nbsp;" & i & "&nbsp;"
				End If
			End If
		End If
	Next
	If NumPagMax  < TotPag then
		Response.Write "&nbsp;&nbsp;<A href=""" & Request.ServerVariables("SCRIPT_NAME") & "?PagAtual=" & NumPagMax + 1 & "&VarPagMax=" & VarPagMax & "&NumPagMax=" & NumPagMax + VarPagMax & "&Submit=Mais&Ordem=" & Request.QueryString("Ordem") & "&string_busca=" & Server.URLEncode(Request("string_busca")) & "&campo_busca=" & Server.URLEncode(Request("campo_busca")) & """ class=links_paginacao>&gt;&gt;</A>"
	End If
	Response.Write "</td><td width=33% align=left vAlign=top noWrap>"
	If PagAtual <> TotPag Then
		Response.Write "&nbsp;&nbsp;<A href=""" & Request.ServerVariables("SCRIPT_NAME") & "?PagAtual=" & PagAtual & "&VarPagMax=" & VarPagMax & "&NumPagMax=" & NumPagMax & "&Submit=Proxima&Ordem=" & Request.QueryString("Ordem") & "&string_busca=" & Server.URLEncode(Request("string_busca")) & "&campo_busca=" & Server.URLEncode(Request("campo_busca")) & """ class=links_paginacao>Proxima &gt;</A>"
	End If
	Response.Write "</td></tr></table>"
End Sub
%>
