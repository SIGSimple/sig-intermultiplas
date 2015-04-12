<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/cpf.asp" -->
<%
'*******************************************************************
' P�gina gerada pelo sistema Dataform 2 - http://www.dataform.com.br
'*******************************************************************
' Altere os valores das vari�veis indicadas abaixo se necess�rio

'String de conex�o para o banco de dados do Microsoft Access
strCon = "DBQ=C:\inetpub\wwwroot\original\ARQUIVOS\DADOS\bd_fde.mdb;Driver={Microsoft Access Driver (*.mdb)};"
'strCon = "DBQ=\\10.0.75.124\intermultiplas.net\public\ARQUIVOS\DADOS\bd_fde.mdb;Driver={Microsoft Access Driver (*.mdb)};"

'N�mero total de registros a serem exibidos por p�gina
Const RegPorPag = 15

'N�mero de p�ginas a ser exibido no �ndice de pagina��o
VarPagMax = 10

'Cor da linha selecionada na tabela de registros
cor_linha_selecionada = "gainsboro"

'Nome da p�gina de consulta
pagina_consulta = "df_consulta.asp"

'Nome da p�gina de altera��o
pagina_alteracao = "df_alteracao.asp"

'Nome da p�gina de inclus�o
pagina_inclusao = "df_inclusao.asp"

'Nome da p�gina de login
pagina_login = "df_login.asp"

'*******************************************************************

%>

<HTML>
<HEAD>
<TITLE>Consultar Registros</TITLE>
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
<BODY class=texto_pagina>
Links: <a href="<%=pagina_consulta%>" class="texto_pagina">P�gina de Consulta</a> | <a href="<%=pagina_inclusao%>" class="texto_pagina">P�gina de Inclus�o<hr size=1 color=gainsboro></a><br>

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
    If indice <> "" Then strQ_delete = " SELECT * FROM tb_predio WHERE " & indice

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
strQ = "SELECT * FROM tb_predio"

If Trim(Request("string_busca")) <> "" Then
  If Trim(Request("campo_busca")) <> "" Then
    strQ = strQ & " Where " & Trim(Request("campo_busca")) & " LIKE '%" & Trim(Request("string_busca")) & "%'"
  Else
    strQ = strQ & " Where 1 <> 1"
    strQ = strQ & " Or cod_predio LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or Nome_Unidade LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or Endere�o LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or Complemento LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or CEP LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or �rea Constru�da LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or Diretoria de Ensino LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or Munic�pio LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or �rea Total LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or �rea Ocupada LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or N�mero de Pavimentos LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or N�mero de Salas LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or Observa��o sobre o Pr�dio LIKE '%" & Trim(Request("string_busca")) & "%'"
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
strQ_indice = "SELECT * FROM tb_predio WHERE 1 <> 1"
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

<B>Consultar Registros</B><BR>Visualize os registros da 
tabela abaixo:<BR>
<FORM name="form_busca" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>">
Pesquizar por <INPUT type=text name=string_busca value="<%=Request("string_busca")%>" class=texto_pagina> em
<SELECT name=campo_busca class=texto_pagina>
  <OPTION value="" selected></OPTION>
  <OPTION value="cod_predio" <% If Trim(Request("campo_busca")) = Trim("cod_predio") Then : Response.Write "selected" : End If %>>cod_predio</OPTION>
  <OPTION value="Nome_Unidade" <% If Trim(Request("campo_busca")) = Trim("Nome_Unidade") Then : Response.Write "selected" : End If %>>Nome_Unidade</OPTION>
  <OPTION value="Endere�o" <% If Trim(Request("campo_busca")) = Trim("Endere�o") Then : Response.Write "selected" : End If %>>Endere�o</OPTION>
  <OPTION value="Complemento" <% If Trim(Request("campo_busca")) = Trim("Complemento") Then : Response.Write "selected" : End If %>>Complemento</OPTION>
  <OPTION value="CEP" <% If Trim(Request("campo_busca")) = Trim("CEP") Then : Response.Write "selected" : End If %>>CEP</OPTION>
  <OPTION value="�rea Constru�da" <% If Trim(Request("campo_busca")) = Trim("�rea Constru�da") Then : Response.Write "selected" : End If %>>�rea Constru�da</OPTION>
  <OPTION value="Diretoria de Ensino" <% If Trim(Request("campo_busca")) = Trim("Diretoria de Ensino") Then : Response.Write "selected" : End If %>>Diretoria de Ensino</OPTION>
  <OPTION value="Munic�pio" <% If Trim(Request("campo_busca")) = Trim("Munic�pio") Then : Response.Write "selected" : End If %>>Munic�pio</OPTION>
  <OPTION value="�rea Total" <% If Trim(Request("campo_busca")) = Trim("�rea Total") Then : Response.Write "selected" : End If %>>�rea Total</OPTION>
  <OPTION value="�rea Ocupada" <% If Trim(Request("campo_busca")) = Trim("�rea Ocupada") Then : Response.Write "selected" : End If %>>�rea Ocupada</OPTION>
  <OPTION value="N�mero de Pavimentos" <% If Trim(Request("campo_busca")) = Trim("N�mero de Pavimentos") Then : Response.Write "selected" : End If %>>N�mero de Pavimentos</OPTION>
  <OPTION value="N�mero de Salas" <% If Trim(Request("campo_busca")) = Trim("N�mero de Salas") Then : Response.Write "selected" : End If %>>N�mero de Salas</OPTION>
  <OPTION value="Observa��o sobre o Pr�dio" <% If Trim(Request("campo_busca")) = Trim("Observa��o sobre o Pr�dio") Then : Response.Write "selected" : End If %>>Observa��o sobre o Pr�dio</OPTION>
</SELECT>
<INPUT type="submit" name="submit" value="ok" class=texto_pagina style="color: black">
</FORM>

<%
If Not(objRS.EOF) Then
  objRS.AbsolutePage = PagAtual
  TotPag = objRS.PageCount
%>

Foram encontrados <%= objRS.RecordCount%> registros<BR><BR>
<table border=0 cellpadding=2 cellspacing=1 class=tabela_registros>
  <tr class=titulos_registros>
    <%
If Session("admin") <> "" And Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR") Then
  Response.Write "<TD align=""center"" style=""background-color: crimson; color: white"" width=""1%"" nowrap><b>Editar</b></TD>"
End IF

If Right(Request.QueryString("Ordem"), 3) = "asc" Then
  Ordem = "desc"
Else
  Ordem = "asc"
End IF
%>
    <td style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=cod_predio+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 10) = "cod_predio" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
        <b>cod_predio</b></td>
    <td style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=Nome_Unidade+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 12) = "Nome_Unidade" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
        <b>Nome_Unidade</b></td>
    <td style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=Endere�o+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 8) = "Endere�o" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
        <b>Endere�o</b></td>
    <td style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=Complemento+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 11) = "Complemento" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
        <b>Complemento</b></td>
    <td style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=CEP+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 3) = "CEP" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
        <b>CEP</b></td>
    <td style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=�rea Constru�da+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 15) = "�rea Constru�da" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
        <b>�rea Constru�da</b></td>
    <td style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=Diretoria de Ensino+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 19) = "Diretoria de Ensino" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
        <b>Diretoria de Ensino</b></td>
	   <td style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=Munic�pio+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 19) = "Munic�pio" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
        <b>Munic�pio</b></td>	
    <td style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=N�mero de Pavimentos+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 19) = "N�mero de Pavimentos" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%>
        <b>N�mero de Pavimentos</b></td>
  </tr>
  <%
For Cont = 1 to objRS.PageSize
%>
  <tr class=exibe_registros onMouseOver="this.style.backgroundColor='<%=cor_linha_selecionada%>';" onMouseOut="this.style.backgroundColor='';">
    <%
If Session("admin") <> "" And Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR") Then
  Response.Write "<FORM name=""form_edit_" & Cont & """ action=""" & pagina_alteracao & """ method=post>"
  Response.Write "<TD  align=""center"" nowrap style=""background-color: gainsboro""  nowrap>&nbsp;"
  If indice <> "" Then Response.Write "<input type=""hidden"" name=""indice"" value=""" & indice & "=" & objRS.Fields.Item(indice).Value & """>"
  Response.Write "<INPUT type=hidden name=recordno value=""" & (objRS.AbsolutePosition) & """>"
  Response.Write "<INPUT type=hidden name=strQ value=""" & strQ & """>"
  Response.Write "<INPUT type=image src=""imagens\edit.gif"" alt=""Alterar Registro"" name=alterar value=alterar>"
  If Session("admin") <> "" And Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR") Then
  Response.Write "&nbsp;<IMG src=""imagens\delete.gif"" alt=""Excluir Registro"" name=delete border=0 style=""cursor:hand"" OnClick=""confirm_delete('form_edit_" & Cont & "')"">"
  End If
  Response.Write "&nbsp;</TD>"
  Response.Write "</FORM>"
End If
%>
    <td><%=(objRS.Fields.Item("cod_predio").Value)%></td>
    <td><%=(objRS.Fields.Item("Nome_Unidade").Value)%></td>
    <td><%=(objRS.Fields.Item("Endere�o").Value)%></td>
    <td><%=(objRS.Fields.Item("Complemento").Value)%></td>
    <td><%=(objRS.Fields.Item("CEP").Value)%></td>
    <td><%=(objRS.Fields.Item("�rea Constru�da").Value)%></td>
    <td><%=(objRS.Fields.Item("Diretoria de Ensino").Value)%></td>
    <td><%=(objRS.Fields.Item("Munic�pio").Value)%></td>
	<td><%=(objRS.Fields.Item("N�mero de pavimentos").Value)%></td>
    
  </tr>
  <%
  objRS.MoveNext
  If objRS.Eof then Exit For
Next
Set Cont = Nothing
%>
  <tr>
    <td colspan="8"><%LinksNavegacao()%></td>
  </tr>
</table>
'
<%
'  If indice = "" Then
'    Response.Write "<BR><B>ATEN��O:</B> Crie um campo do tipo <i>AutoIncrement</i> com qualquer nome em sua tabela para evitar erros na altera��o dos dados. "
'    Response.Write "<a href=""http://www.dataform.com.br/criar_campo_autoincrement.asp"" target=""_blank"">Clique aqui</a> para mais detalhes."
'  End If
'  objRS.Close
'  Set objRS = Nothing
'Else
'%>

<BR><B>Nenhum registro foi encontrado</B><BR><BR>

<%
End If
%>

</BODY>
</HTML>

<%
Sub LinksNavegacao()
'O c�digo a seguir insere uma tabela com todos os links de navega��o das p�ginas
Response.Write "<TABLE border=0 cellPadding=2 cellSpacing=0 class=tabela_paginacao>"
Response.Write "<TR><TD align=center vAlign=top noWrap colspan=5>"
Response.Write "P�gina " & PagAtual & " de " & TotPag
Response.Write "</TD></TR><TR><TD width=33% align=right vAlign=top noWrap>"
If PagAtual > 1 Then
  Response.Write "<A href=""" & Request.ServerVariables("SCRIPT_NAME") & "?PagAtual=" &  PagAtual &"&VarPagMax=" & VarPagMax & "&NumPagMax=" & NumPagMax & "&Submit=Anterior&Ordem=" & Request.QueryString("Ordem")& "&string_busca=" & Server.URLEncode(Request("string_busca")) & "&campo_busca=" & Server.URLEncode(Request("campo_busca"))  & """ class=links_paginacao>&lt; Anterior</A>"
End If
Response.Write "</TD><TD width=33% align=middle vAlign=top noWrap>"
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
Response.Write "</TD><TD width=33% align=left vAlign=top noWrap>"
If PagAtual <> TotPag Then
  Response.Write "&nbsp;&nbsp;<A href=""" & Request.ServerVariables("SCRIPT_NAME") & "?PagAtual=" & PagAtual & "&VarPagMax=" & VarPagMax & "&NumPagMax=" & NumPagMax & "&Submit=Proxima&Ordem=" & Request.QueryString("Ordem") & "&string_busca=" & Server.URLEncode(Request("string_busca")) & "&campo_busca=" & Server.URLEncode(Request("campo_busca")) & """ class=links_paginacao>Proxima &gt;</A>"
End If
Response.Write "</TD></TR></TABLE>"
End Sub
%>
