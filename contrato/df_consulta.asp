<% @ LANGUAGE="VBSCRIPT" %>
<%
'*******************************************************************
' Página gerada pelo sistema Dataform 2 - http://www.dataform.com.br
'*******************************************************************
' Altere os valores das variáveis indicadas abaixo se necessário

'String de conexão para o banco de dados do Microsoft Access
strCon = "DBQ=C:\Inetpub\wwwroot\procentro\db\fde.mdb;Driver={Microsoft Access Driver (*.mdb)};"
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
Links: <a href="<%=pagina_consulta%>" class="texto_pagina">Página de Consulta</a> | <a href="<%=pagina_inclusao%>" class="texto_pagina">Página de Inclusão<hr size=1 color=gainsboro></a><br>

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
    If indice <> "" Then strQ_delete = " SELECT * FROM tb_contrato WHERE " & indice

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
strQ = "SELECT * FROM tb_contrato"

If Trim(Request("string_busca")) <> "" Then
  If Trim(Request("campo_busca")) <> "" Then
    strQ = strQ & " Where " & Trim(Request("campo_busca")) & " LIKE '%" & Trim(Request("string_busca")) & "%'"
  Else
    strQ = strQ & " Where 1 <> 1"
    strQ = strQ & " Or area_construida LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or area_gerenciadora LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or cod_contrato LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or cod_predio LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or cod_reg LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or Construtora LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or crit_calculo LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or crit_reajuste LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or desc_interv_gerenciadora LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or Desc_intervencao_FDE LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or dig_contrato LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or dt_abertura LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or dt_assinatura LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or dt_base LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or dt_CI LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or dt_impressao_ois LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or dt_termino LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or dt_TRD LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or dt_TRP LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or fator_reducao LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or fical LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or gerenc_mede LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or Obra_pi LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or orc_FDE LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or orgao LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or PI LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or prz_contrato LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or reducao LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or solic_aditamento LIKE '%" & Trim(Request("string_busca")) & "%'"
    strQ = strQ & " Or vl_contrato LIKE '%" & Trim(Request("string_busca")) & "%'"
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
strQ_indice = "SELECT * FROM tb_contrato WHERE 1 <> 1"
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
  <OPTION value="" selected>Registros</OPTION>
  <OPTION value="area_construida" <% If Trim(Request("campo_busca")) = Trim("area_construida") Then : Response.Write "selected" : End If %>>area_construida</OPTION>
  <OPTION value="area_gerenciadora" <% If Trim(Request("campo_busca")) = Trim("area_gerenciadora") Then : Response.Write "selected" : End If %>>area_gerenciadora</OPTION>
  <OPTION value="cod_contrato" <% If Trim(Request("campo_busca")) = Trim("cod_contrato") Then : Response.Write "selected" : End If %>>cod_contrato</OPTION>
  <OPTION value="cod_predio" <% If Trim(Request("campo_busca")) = Trim("cod_predio") Then : Response.Write "selected" : End If %>>cod_predio</OPTION>
  <OPTION value="cod_reg" <% If Trim(Request("campo_busca")) = Trim("cod_reg") Then : Response.Write "selected" : End If %>>cod_reg</OPTION>
  <OPTION value="Construtora" <% If Trim(Request("campo_busca")) = Trim("Construtora") Then : Response.Write "selected" : End If %>>Construtora</OPTION>
  <OPTION value="crit_calculo" <% If Trim(Request("campo_busca")) = Trim("crit_calculo") Then : Response.Write "selected" : End If %>>crit_calculo</OPTION>
  <OPTION value="crit_reajuste" <% If Trim(Request("campo_busca")) = Trim("crit_reajuste") Then : Response.Write "selected" : End If %>>crit_reajuste</OPTION>
  <OPTION value="desc_interv_gerenciadora" <% If Trim(Request("campo_busca")) = Trim("desc_interv_gerenciadora") Then : Response.Write "selected" : End If %>>desc_interv_gerenciadora</OPTION>
  <OPTION value="Desc_intervencao_FDE" <% If Trim(Request("campo_busca")) = Trim("Desc_intervencao_FDE") Then : Response.Write "selected" : End If %>>Desc_intervencao_FDE</OPTION>
  <OPTION value="dig_contrato" <% If Trim(Request("campo_busca")) = Trim("dig_contrato") Then : Response.Write "selected" : End If %>>dig_contrato</OPTION>
  <OPTION value="dt_abertura" <% If Trim(Request("campo_busca")) = Trim("dt_abertura") Then : Response.Write "selected" : End If %>>dt_abertura</OPTION>
  <OPTION value="dt_assinatura" <% If Trim(Request("campo_busca")) = Trim("dt_assinatura") Then : Response.Write "selected" : End If %>>dt_assinatura</OPTION>
  <OPTION value="dt_base" <% If Trim(Request("campo_busca")) = Trim("dt_base") Then : Response.Write "selected" : End If %>>dt_base</OPTION>
  <OPTION value="dt_CI" <% If Trim(Request("campo_busca")) = Trim("dt_CI") Then : Response.Write "selected" : End If %>>dt_CI</OPTION>
  <OPTION value="dt_impressao_ois" <% If Trim(Request("campo_busca")) = Trim("dt_impressao_ois") Then : Response.Write "selected" : End If %>>dt_impressao_ois</OPTION>
  <OPTION value="dt_termino" <% If Trim(Request("campo_busca")) = Trim("dt_termino") Then : Response.Write "selected" : End If %>>dt_termino</OPTION>
  <OPTION value="dt_TRD" <% If Trim(Request("campo_busca")) = Trim("dt_TRD") Then : Response.Write "selected" : End If %>>dt_TRD</OPTION>
  <OPTION value="dt_TRP" <% If Trim(Request("campo_busca")) = Trim("dt_TRP") Then : Response.Write "selected" : End If %>>dt_TRP</OPTION>
  <OPTION value="fator_reducao" <% If Trim(Request("campo_busca")) = Trim("fator_reducao") Then : Response.Write "selected" : End If %>>fator_reducao</OPTION>
  <OPTION value="fical" <% If Trim(Request("campo_busca")) = Trim("fical") Then : Response.Write "selected" : End If %>>fical</OPTION>
  <OPTION value="gerenc_mede" <% If Trim(Request("campo_busca")) = Trim("gerenc_mede") Then : Response.Write "selected" : End If %>>gerenc_mede</OPTION>
  <OPTION value="Obra_pi" <% If Trim(Request("campo_busca")) = Trim("Obra_pi") Then : Response.Write "selected" : End If %>>Obra_pi</OPTION>
  <OPTION value="orc_FDE" <% If Trim(Request("campo_busca")) = Trim("orc_FDE") Then : Response.Write "selected" : End If %>>orc_FDE</OPTION>
  <OPTION value="orgao" <% If Trim(Request("campo_busca")) = Trim("orgao") Then : Response.Write "selected" : End If %>>orgao</OPTION>
  <OPTION value="PI" <% If Trim(Request("campo_busca")) = Trim("PI") Then : Response.Write "selected" : End If %>>PI</OPTION>
  <OPTION value="prz_contrato" <% If Trim(Request("campo_busca")) = Trim("prz_contrato") Then : Response.Write "selected" : End If %>>prz_contrato</OPTION>
  <OPTION value="reducao" <% If Trim(Request("campo_busca")) = Trim("reducao") Then : Response.Write "selected" : End If %>>reducao</OPTION>
  <OPTION value="solic_aditamento" <% If Trim(Request("campo_busca")) = Trim("solic_aditamento") Then : Response.Write "selected" : End If %>>solic_aditamento</OPTION>
  <OPTION value="vl_contrato" <% If Trim(Request("campo_busca")) = Trim("vl_contrato") Then : Response.Write "selected" : End If %>>vl_contrato</OPTION>
</SELECT>
<INPUT type="submit" name="submit" value="ok" class=texto_pagina style="color: black">
</FORM>

<%
If Not(objRS.EOF) Then
  objRS.AbsolutePage = PagAtual
  TotPag = objRS.PageCount
%>

Foram encontrados <%= objRS.RecordCount%> registros<BR><BR>

<TABLE border=0 cellpadding=2 cellspacing=1 class=tabela_registros>
  <TR class=titulos_registros>

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

  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=area_construida+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 15) = "area_construida" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>area_construida</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=area_gerenciadora+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 17) = "area_gerenciadora" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>area_gerenciadora</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=cod_contrato+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 12) = "cod_contrato" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>cod_contrato</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=cod_predio+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 10) = "cod_predio" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>cod_predio</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=cod_reg+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 7) = "cod_reg" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>cod_reg</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=Construtora+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 11) = "Construtora" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>Construtora</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=crit_calculo+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 12) = "crit_calculo" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>crit_calculo</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=crit_reajuste+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 13) = "crit_reajuste" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>crit_reajuste</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=desc_interv_gerenciadora+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 24) = "desc_interv_gerenciadora" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>desc_interv_gerenciadora</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=Desc_intervencao_FDE+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 20) = "Desc_intervencao_FDE" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>Desc_intervencao_FDE</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=dig_contrato+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 12) = "dig_contrato" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>dig_contrato</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=dt_abertura+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 11) = "dt_abertura" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>dt_abertura</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=dt_assinatura+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 13) = "dt_assinatura" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>dt_assinatura</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=dt_base+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 7) = "dt_base" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>dt_base</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=dt_CI+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 5) = "dt_CI" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>dt_CI</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=dt_impressao_ois+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 16) = "dt_impressao_ois" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>dt_impressao_ois</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=dt_termino+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 10) = "dt_termino" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>dt_termino</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=dt_TRD+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 6) = "dt_TRD" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>dt_TRD</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=dt_TRP+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 6) = "dt_TRP" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>dt_TRP</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=fator_reducao+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 13) = "fator_reducao" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>fator_reducao</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=fical+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 5) = "fical" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>fical</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=gerenc_mede+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 11) = "gerenc_mede" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>gerenc_mede</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=Obra_pi+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 7) = "Obra_pi" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>Obra_pi</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=orc_FDE+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 7) = "orc_FDE" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>orc_FDE</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=orgao+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 5) = "orgao" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>orgao</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=PI+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 2) = "PI" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>PI</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=prz_contrato+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 12) = "prz_contrato" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>prz_contrato</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=reducao+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 7) = "reducao" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>reducao</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=solic_aditamento+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 16) = "solic_aditamento" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>solic_aditamento</b></TD>
  <TD style="cursor: hand" valign=top nowrap onClick="window.open('<%=Request.ServerVariables("SCRIPT_NAME")%>?Ordem=vl_contrato+<%=Ordem%>', '_self')"><%If Left(Request.QueryString("Ordem"), 11) = "vl_contrato" Then : Response.Write "<img src=""imagens\ordem_" & Ordem & ".gif"" width=9 height=10>&nbsp;" : End If%><b>vl_contrato</b></TD>
  </TR>

<%
For Cont = 1 to objRS.PageSize
%>

  <TR class=exibe_registros onMouseOver="this.style.backgroundColor='<%=cor_linha_selecionada%>';" onMouseOut="this.style.backgroundColor='';">

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

    <TD><%=(objRS.Fields.Item("area_construida").Value)%></TD>
    <TD><%=(objRS.Fields.Item("area_gerenciadora").Value)%></TD>
    <TD><%=(objRS.Fields.Item("cod_contrato").Value)%></TD>
    <TD><%=(objRS.Fields.Item("cod_predio").Value)%></TD>
    <TD><%=(objRS.Fields.Item("cod_reg").Value)%></TD>
    <TD><%=(objRS.Fields.Item("Construtora").Value)%></TD>
    <TD><%=(objRS.Fields.Item("crit_calculo").Value)%></TD>
    <TD><%=(objRS.Fields.Item("crit_reajuste").Value)%></TD>
    <TD><%=(objRS.Fields.Item("desc_interv_gerenciadora").Value)%></TD>
    <TD><%=(objRS.Fields.Item("Desc_intervencao_FDE").Value)%></TD>
    <TD><%=(objRS.Fields.Item("dig_contrato").Value)%></TD>
    <TD><%=(objRS.Fields.Item("dt_abertura").Value)%></TD>
    <TD><%=(objRS.Fields.Item("dt_assinatura").Value)%></TD>
    <TD><%=(objRS.Fields.Item("dt_base").Value)%></TD>
    <TD><%=(objRS.Fields.Item("dt_CI").Value)%></TD>
    <TD><%=(objRS.Fields.Item("dt_impressao_ois").Value)%></TD>
    <TD><%=(objRS.Fields.Item("dt_termino").Value)%></TD>
    <TD><%=(objRS.Fields.Item("dt_TRD").Value)%></TD>
    <TD><%=(objRS.Fields.Item("dt_TRP").Value)%></TD>
    <TD><%=(objRS.Fields.Item("fator_reducao").Value)%></TD>
    <TD><%=(objRS.Fields.Item("fical").Value)%></TD>
    <TD><% if Left(Lcase(objRS.Fields.Item("gerenc_mede").Value),1) = "f" then : Response.Write "Não" : Else : Response.Write "Sim" End If %></TD>
    <TD><% if Left(Lcase(objRS.Fields.Item("Obra_pi").Value),1) = "f" then : Response.Write "Não" : Else : Response.Write "Sim" End If %></TD>
    <TD><%=(objRS.Fields.Item("orc_FDE").Value)%></TD>
    <TD><%=(objRS.Fields.Item("orgao").Value)%></TD>
    <TD><%=(objRS.Fields.Item("PI").Value)%></TD>
    <TD><%=(objRS.Fields.Item("prz_contrato").Value)%></TD>
    <TD><%=(objRS.Fields.Item("reducao").Value)%></TD>
    <TD><% if Left(Lcase(objRS.Fields.Item("solic_aditamento").Value),1) = "f" then : Response.Write "Não" : Else : Response.Write "Sim" End If %></TD>
    <TD><%=(objRS.Fields.Item("vl_contrato").Value)%></TD>
  </TR>

<%
  objRS.MoveNext
  If objRS.Eof then Exit For
Next
Set Cont = Nothing
%>

<TR>
  <TD colspan="31"><%LinksNavegacao()%></TD>
</TR>

</TABLE>

<%
  If indice = "" Then
    Response.Write "<BR><B>ATENÇÃO:</B> Crie um campo do tipo <i>AutoIncrement</i> com qualquer nome em sua tabela para evitar erros na alteração dos dados. "
    Response.Write "<a href=""http://www.dataform.com.br/criar_campo_autoincrement.asp"" target=""_blank"">Clique aqui</a> para mais detalhes."
  End If
  objRS.Close
  Set objRS = Nothing
Else
%>

<BR><B>Nenhum registro foi encontrado</B><BR><BR>

<%
End If
%>

</BODY>
</HTML>

<%
Sub LinksNavegacao()
'O código a seguir insere uma tabela com todos os links de navegação das páginas
Response.Write "<TABLE border=0 cellPadding=2 cellSpacing=0 class=tabela_paginacao>"
Response.Write "<TR><TD align=center vAlign=top noWrap colspan=5>"
Response.Write "Página " & PagAtual & " de " & TotPag
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
