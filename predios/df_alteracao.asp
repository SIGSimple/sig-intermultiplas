<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/cpf.asp" -->
<%
'*******************************************************************
' Página gerada pelo sistema Dataform 2 - http://www.dataform.com.br
'*******************************************************************
' Altere os valores das variáveis indicadas abaixo se necessário

'String de conexão para o banco de dados do Microsoft Access
strCon = "DBQ=C:\inetpub\wwwroot\original\ARQUIVOS\DADOS\bd_fde.mdb;Driver={Microsoft Access Driver (*.mdb)};"
'strCon = "DBQ=\\10.0.75.124\intermultiplas.net\public\ARQUIVOS\DADOS\bd_fde.mdb;Driver={Microsoft Access Driver (*.mdb)};"

'Nome da página de consulta
pagina_consulta = "df_consulta.asp"

'Nome da página de alteração
pagina_alteracao = "df_alteracao.asp"

'Nome da página de inclusão
pagina_inclusao = "df_inclusao.asp"

'Nome da página de login
pagina_login = "df_login.asp"

'*******************************************************************


If Session("admin") <> "" And Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR") Then
%>

<HTML>
<HEAD>
<TITLE>Alterar Registro</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="copyright" content="Dataform">
<meta name="keywords" content="dataform, asp dataform, aspdataform, asp-dataform">
<meta name="robots" content="ALL">
<style type="text/css">
<!--
.campo_alerta
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
border: 1px solid black;
background-color: #ffff99;
}
.texto_pagina
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
color: dimgray;
}

.tabela_formulario
{
width: 200;
background-color: white;
}

.titulo_campos
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
color: dimgray;
background-color: whitesmoke;
}

.campos_formulario
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
color: dimgray;
background-color: gainsboro;
border: 1px inset;
}

.botao_enviar
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
color: white;
background-color: gray;
}
-->
</style>
<script language="JavaScript" type="text/javascript">
<!--
function abre_janela(width, height, nome) {
var top; var left;
top = ( (screen.height/2) - (height/2) )
left = ( (screen.width/2) - (width/2) )
window.open('',nome,'width='+width+',height='+height+',scrollbars=no,toolbar=no,location=no,status=no,menubar=no,resizable=no,left='+left+',top='+top);
}
function recebe_imagem(campo, imagem){
var foto = 'img_' + campo
document.form_incluir[campo].value = imagem;
document.form_incluir[foto].src = imagem;
}
function verifica_form(form) {
var passed = false;
var ok = false
var campo
for (i = 0; i < form.length; i++) {
  campo = form[i].name;
  if (form[i].df_verificar == "sim") {
    if (form[i].type == "text"  | form[i].type == "textarea" | form[i].type == "select-one") {
      if (form[i].value == "" | form[i].value == "http://") {
		form[campo].className='campo_alerta'
        form[campo].focus();
        alert("Preencha corretamente o campo");
        return passed;
        stop;
      }
    }
    else if (form[i].type == "radio") {
      for (x = 0; x < form[campo].length; x++) {
        ok = false;
        if (form[campo][x].checked) {
          ok = true;
          break;
        }
      }
      if (ok == false) {
        form[campo][0].focus();
		form[campo][0].select();
        alert("Informe uma das opcões");
        return passed;
        stop;
      }
    }
    var msg = ""
    if (form[campo].df_validar == "cpf") msg = checa_cpf(form[campo].value);
    if (form[campo].df_validar == "cnpj") msg = checa_cnpj(form[campo].value);
    if (form[campo].df_validar == "cpf_cnpj") {
	  msg = checa_cpf(form[campo].value);
	  if (msg != "") msg = checa_cnpj(form[campo].value);
	}
    if (form[campo].df_validar == "email") msg = checa_email(form[campo].value);
    if (form[campo].df_validar == "numerico") msg = checa_numerico(form[campo].value);
    if (msg != "") {
	  if (form[campo].df_validar == "cpf_cnpj") msg = "informe corretamente o número do CPF ou CNPJ";
	  form[campo].className='campo_alerta'
      form[campo].focus();
      form[campo].select();
      alert(msg);
      return passed;
      stop;
    }
  }
}
passed = true;
return passed;
}
function desabilita_cor(campo) {
campo.className='campos_formulario'
}
function checa_numerico(String) {
var mensagem = "Este campo aceita somente números"
var msg = "";
if (isNaN(String)) msg = mensagem;
return msg;
}
function checa_email(campo) {
var mensagem = "Informe corretamente o email"
var msg = "";
var email = campo.match(/(\w+)@(.+)\.(\w+)$/);
if (email == null){
  msg = mensagem;
  }
return msg;
}
function checa_cpf(CPF) {
var mensagem = "informe corretamente o número do CPF"
var msg = "";
if (CPF.length != 11 || CPF == "00000000000" || CPF == "11111111111" ||
  CPF == "22222222222" ||	CPF == "33333333333" || CPF == "44444444444" ||
  CPF == "55555555555" || CPF == "66666666666" || CPF == "77777777777" ||
  CPF == "88888888888" || CPF == "99999999999")
msg = mensagem;
soma = 0;
for (y=0; y < 9; y ++)
soma += parseInt(CPF.charAt(y)) * (10 - y);
resto = 11 - (soma % 11);
if (resto == 10 || resto == 11)resto = 0;
if (resto != parseInt(CPF.charAt(9)))
  msg = mensagem; soma = 0;
for (y = 0; y < 10; y ++)
  soma += parseInt(CPF.charAt(y)) * (11 - y);
resto = 11 - (soma % 11);
if (resto == 10 || resto == 11) resto = 0;
if (resto != parseInt(CPF.charAt(10)))
  msg = mensagem;
return msg;
}
function checa_cnpj(s) {
var mensagem = "informe corretamente o número do CNPJ"
var msg = "";
var y;
var c = s.substr(0,12);
var dv = s.substr(12,2);
var d1 = 0;
for (y = 0; y < 12; y++)
{
d1 += c.charAt(11-y)*(2+(y % 8));
}
if (d1 == 0) msg = mensagem;
d1 = 11 - (d1 % 11);
if (d1 > 9) d1 = 0;
if (dv.charAt(0) != d1)msg = mensagem;
d1 *= 2;
for (y = 0; y < 12; y++)
{
d1 += c.charAt(11-y)*(2+((y+1) % 8));
}
d1 = 11 - (d1 % 11);
if (d1 > 9) d1 = 0;
if (dv.charAt(1) != d1) msg = mensagem;
return msg;
}
function mascara_data(data){ 
var mydata = ''; 
mydata = mydata + data; 
if (mydata.length == 2){ 
mydata = mydata + '/'; 
} 
if (mydata.length == 5){ 
mydata = mydata + '/'; 
} 
return mydata; 
} 
function verifica_data(data) { 
if (data.value != "") {
dia = (data.value.substring(0,2));
mes = (data.value.substring(3,5)); 
ano = (data.value.substring(6,10)); 
situacao = ""; 
if ((dia < 01)||(dia < 01 || dia > 30) && (  mes == 04 || mes == 06 || mes == 09 || mes == 11 ) || dia > 31) { 
situacao = "falsa"; 
} 
if (mes < 01 || mes > 12 ) { 
situacao = "falsa"; 
}
if (mes == 2 && ( dia < 01 || dia > 29 || ( dia > 28 && (parseInt(ano / 4) != ano / 4)))) { 
situacao = "falsa"; 
} 
if (situacao == "falsa") { 
data.focus();
data.select();
alert("Data inválida!"); 
}
} 
}
//-->
</script>
</HEAD>
<BODY class=texto_pagina>
Links: <a href="<%=pagina_consulta%>" class="texto_pagina">Página de Consulta</a> | <a href="<%=pagina_inclusao%>" class="texto_pagina">Página de Inclusão<hr size=1 color=gainsboro></a><br>

<%
If Not IsEmpty(Request.Form("salvar")) Then
  Set objCon = Server.CreateObject("ADODB.Connection")
  objCon.Open MM_cpf_STRING
  Set objRS = Server.CreateObject("ADODB.Recordset")
  objRS.CursorLocation = 3
  objRS.CursorType = 0
  objRS.LockType = 3

  strQ = Request.Form("strQ")
  indice = Trim(Request.Form("indice"))
  If indice <> "" Then strQ = " SELECT * FROM tb_predio WHERE " & indice

  objRS.Open strQ, objCon, , , &H0001
  If indice = "" Then objRS.Move Request.Form("recordno") - 1

  If objRS.Fields("id_predio").properties("IsAutoIncrement") = False Then
    objRS("id_predio") = Trim(Request.Form("id_predio"))
  End If
  If objRS.Fields("Diretoria de Ensino").properties("IsAutoIncrement") = False Then
    objRS("Diretoria de Ensino") = Trim(Request.Form("Diretoria de Ensino"))
  End If
  If objRS.Fields("cod_bacia_secretaria").properties("IsAutoIncrement") = False Then
    objRS("cod_bacia_secretaria") = Trim(Request.Form("cod_bacia_secretaria"))
  End If
  If objRS.Fields("Município").properties("IsAutoIncrement") = False Then
    objRS("Município") = Trim(Request.Form("Municipio"))
  End If

  ' NOVOS CAMPOS
  If objRS.Fields("cod_prefeitura").properties("IsAutoIncrement") = False Then
    objRS("cod_prefeitura") = Trim(Request.Form("cod_prefeitura"))
  End If
  If objRS.Fields("cod_prefeito").properties("IsAutoIncrement") = False Then
    objRS("cod_prefeito") = Trim(Request.Form("cod_prefeito"))
  End If
  If objRS.Fields("ano_inicio_adm").properties("IsAutoIncrement") = False Then
    objRS("ano_inicio_adm") = Trim(Request.Form("ano_inicio_adm"))
  End If
  If objRS.Fields("ano_fim_adm").properties("IsAutoIncrement") = False Then
    objRS("ano_fim_adm") = Trim(Request.Form("ano_fim_adm"))
  End If
  If objRS.Fields("cod_partido").properties("IsAutoIncrement") = False Then
    objRS("cod_partido") = Trim(Request.Form("cod_partido"))
  End If
  If objRS.Fields("qtd_populacao_urbana_2010").properties("IsAutoIncrement") = False Then
    objRS("qtd_populacao_urbana_2010") = Trim(Request.Form("qtd_populacao_urbana_2010"))
  End If
  If objRS.Fields("qtd_populacao_urbana_2030").properties("IsAutoIncrement") = False Then
    objRS("qtd_populacao_urbana_2030") = Trim(Request.Form("qtd_populacao_urbana_2030"))
  End If
  If objRS.Fields("flg_atendido_sabesp").properties("IsAutoIncrement") = False Then
    flg_atendido_sabesp = Trim(Request.Form("flg_atendido_sabesp"))
    If flg_atendido_sabesp = "on" Then
      flg_atendido_sabesp = True
    Else
      flg_atendido_sabesp = False
    End If
    objRS("flg_atendido_sabesp") = flg_atendido_sabesp
  End If
  If objRS.Fields("cod_concessao").properties("IsAutoIncrement") = False Then
    objRS("cod_concessao") = Trim(Request.Form("cod_concessao"))
  End If
  If objRS.Fields("latitude_longitude").properties("IsAutoIncrement") = False Then
    objRS("latitude_longitude") = Trim(Request.Form("latitude_longitude"))
  End If
  ' FIM NOVOS CAMPOS

  If objRS.Fields("Observação sobre o Prédio").properties("IsAutoIncrement") = False Then
    objRS("Observação sobre o Prédio") = Trim(Request.Form("Observação sobre o Prédio"))
  End If
  On Error Resume Next
  objRS.UpdateBatch
  objRS.Close
  Set objRS = Nothing
  objCon.Close
  Set objCon = Nothing
%>

<BR><B>Registro alterado</B><BR>O registro foi alterado 
com sucesso.<BR><BR>

<%
Else
  If Not IsEmpty(Request.Form("recordno")) Then
    Set objCon = Server.CreateObject("ADODB.Connection")
    objCon.Open MM_cpf_STRING

    Set objRS = Server.CreateObject("ADODB.Recordset")
    objRS.CursorLocation = 2
    objRS.CursorType = 0
    objRS.LockType = 3


  strQ = Request.Form("strQ")
  indice = Trim(Request.Form("indice"))
  If indice <> "" Then strQ = " SELECT * FROM tb_predio WHERE " & indice

    objRS.Open strQ, objCon, , , &H0001
  If indice = "" Then objRS.Move Request.Form("recordno") - 1
%>

<B>Alterar dados do Registro</B><BR>Altere os dados 
necessários abaixo:<BR>
<form name="form_incluir" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" onSubmit="return verifica_form(this);">
<INPUT type=hidden name=recordno value="<%=Request.Form("recordno")%>">
<INPUT type=hidden name=strQ value="<%=Request.Form("strQ")%>">
<INPUT type="hidden" name="indice" value="<%=indice%>">
<TABLE border=0 cellpadding=2 cellspacing=1 class=tabela_formulario>
  <TR class=titulo_campos><TD>id<br>
<%If objRS.Fields("id_predio").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="id_predio" maxlength="255" value="<%=(objRS.Fields.Item("id_predio").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
<%
Else
  Response.Write "<B>" & (objRS.Fields.Item("id_predio").Value) & "</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>Municipio<BR>
    <SELECT style="width=350" name="Municipio" df_verificar="sim" onChange="desabilita_cor(this)" class=campos_formulario>
      <OPTION value=""></OPTION>

<%
  Set objRS2 = Server.CreateObject("ADODB.Recordset")
  objRS2.CursorLocation = 3
  objRS2.CursorType = 3
  objRS2.LockType = 1

  strQ = "SELECT Municipios FROM tb_Municipios ORDER BY Municipios ASC"
  objRS2.Open strQ, objCon, , , &H0001
  If Not objRS2.EOF Then
    While Not objRS2.EOF
    If Trim(objRS2.Fields.Item("Municipios").Value) <> "" Then
      Response.Write "      <OPTION value='" & (objRS2.Fields.Item("Municipios").Value) & "'"
      If Lcase(objRS2.Fields.Item("Municipios").Value) = Lcase(objRS.Fields.Item("Município").Value) then
        Response.Write "selected"
      End If
      Response.Write ">" & (objRS2.Fields.Item("Municipios").Value) & "</OPTION>"
    End If
      objRS2.MoveNext
    Wend
  End If
  Response.Write("ok")
%>

    </SELECT>
  </TD></TR>
  <TR class=titulo_campos><TD>Bacia DAEE<BR>
    <SELECT style="width=350" name="Diretoria de Ensino" df_verificar="sim" onChange="desabilita_cor(this)" class=campos_formulario>
      <OPTION value=""></OPTION>

<%
  Set objRS2 = Server.CreateObject("ADODB.Recordset")
  objRS2.CursorLocation = 3
  objRS2.CursorType = 3
  objRS2.LockType = 1

  strQ = "SELECT desc_diretoria FROM tb_diretoria ORDER BY desc_diretoria ASC"
  objRS2.Open strQ, objCon, , , &H0001
  If Not objRS2.EOF Then
    While Not objRS2.EOF
	  If Trim(objRS2.Fields.Item("desc_diretoria").Value) <> "" Then
	    Response.Write "      <OPTION value='" & (objRS2.Fields.Item("desc_diretoria").Value) & "'"
	    If Lcase(objRS2.Fields.Item("desc_diretoria").Value) = Lcase(objRS.Fields.Item("Diretoria de Ensino").Value) then
	      Response.Write "selected"
	    End If
	    Response.Write ">" & (objRS2.Fields.Item("desc_diretoria").Value) & "</OPTION>"
	  End If
      objRS2.MoveNext
    Wend
  End If
  Response.Write("ok")
%>

    </SELECT>
  </TD></TR>
  <TR class=titulo_campos><TD>Bacia Secretaria<BR>
    <SELECT style="width=350" name="cod_bacia_secretaria" df_verificar="sim" onChange="desabilita_cor(this)" class=campos_formulario>
      <OPTION value=""></OPTION>

<%
  Set objRS2 = Server.CreateObject("ADODB.Recordset")
  objRS2.CursorLocation = 3
  objRS2.CursorType = 3
  objRS2.LockType = 1
  strQ = "SELECT id, desc_diretoria FROM tb_diretoria ORDER BY desc_diretoria ASC"
  objRS2.Open strQ, objCon, , , &H0001
  If Not objRS2.EOF Then
    While Not objRS2.EOF
    If Trim(objRS2.Fields.Item("desc_diretoria").Value) <> "" Then
      Response.Write "      <OPTION value='" & (objRS2.Fields.Item("id").Value) & "'"
      If Lcase(objRS2.Fields.Item("id").Value) = Lcase(objRS.Fields.Item("cod_bacia_secretaria").Value) then
        Response.Write "selected"
      End If
      Response.Write ">" & (objRS2.Fields.Item("desc_diretoria").Value) & "</OPTION>"
    End If
      objRS2.MoveNext
    Wend
  End If
  Response.Write("ok")
%>

    </SELECT>
  </TD></TR>

  <!-- NOVOS CAMPOS -->
  <TR class=titulo_campos><TD>Nome da Prefeitura<BR>
    <SELECT style="width=350" name="cod_prefeitura" df_verificar="sim" onChange="desabilita_cor(this)" class=campos_formulario>
      <OPTION value=""></OPTION>

<%
  Set objRS2 = Server.CreateObject("ADODB.Recordset")
  objRS2.CursorLocation = 3
  objRS2.CursorType = 3
  objRS2.LockType = 1
  strQ = "SELECT * FROM tb_Construtora ORDER BY Construtora ASC"
  objRS2.Open strQ, objCon, , , &H0001
  If Not objRS2.EOF Then
    While Not objRS2.EOF
    If Trim(objRS2.Fields.Item("Construtora").Value) <> "" Then
      Response.Write "      <OPTION value='" & (objRS2.Fields.Item("cod_construtora").Value) & "'"
      If Lcase(objRS2.Fields.Item("cod_construtora").Value) = Lcase(objRS.Fields.Item("cod_prefeitura").Value) then
        Response.Write "selected"
      End If
      Response.Write ">" & (objRS2.Fields.Item("Construtora").Value) & "</OPTION>"
    End If
      objRS2.MoveNext
    Wend
  End If
  Response.Write("ok")
%>

    </SELECT>
  </TD></TR>
  <TR class=titulo_campos><TD>Nome do Prefeito<BR>
    <SELECT style="width=350" name="cod_prefeito" df_verificar="sim" onChange="desabilita_cor(this)" class=campos_formulario>
      <OPTION value=""></OPTION>

<%
  Set objRS2 = Server.CreateObject("ADODB.Recordset")
  objRS2.CursorLocation = 3
  objRS2.CursorType = 3
  objRS2.LockType = 1
  strQ = "SELECT * FROM tb_responsavel ORDER BY Responsável ASC"
  objRS2.Open strQ, objCon, , , &H0001
  If Not objRS2.EOF Then
    While Not objRS2.EOF
    If Trim(objRS2.Fields.Item("Responsável").Value) <> "" Then
      Response.Write "      <OPTION value='" & (objRS2.Fields.Item("cod_fiscal").Value) & "'"
      If Lcase(objRS2.Fields.Item("cod_fiscal").Value) = Lcase(objRS.Fields.Item("cod_prefeito").Value) then
        Response.Write "selected"
      End If
      Response.Write ">" & (objRS2.Fields.Item("Responsável").Value) & "</OPTION>"
    End If
      objRS2.MoveNext
    Wend
  End If
  Response.Write("ok")
%>

    </SELECT>
  </TD></TR>
  <TR class=titulo_campos>
    <TD>
      Período de Administração
      <br>
      <%
        If objRS.Fields("ano_inicio_adm").properties("IsAutoIncrement") = False Then
      %>
      De: <INPUT type="text" style="width: 50px;" name="ano_inicio_adm" maxlength="255" value="<%=(objRS.Fields.Item("ano_inicio_adm").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
      <%
        Else
          Response.Write "<B>(Automático)</B>"
        End If
      %>
      <%
        If objRS.Fields("ano_fim_adm").properties("IsAutoIncrement") = False Then
      %>
      Até: <INPUT type="text" style="width: 50px;" name="ano_fim_adm" maxlength="255" value="<%=(objRS.Fields.Item("ano_fim_adm").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
      <%
        Else
          Response.Write "<B>(Automático)</B>"
        End If
      %>
    </TD>
  </TR>
  <TR class=titulo_campos>
    <TD>
      Partido
      <br>
      <SELECT style="width=350" name="cod_partido" df_verificar="sim" onChange="desabilita_cor(this)" class=campos_formulario>
      <OPTION value=""></OPTION>

<%
  Set objRS2 = Server.CreateObject("ADODB.Recordset")
  objRS2.CursorLocation = 3
  objRS2.CursorType = 3
  objRS2.LockType = 1
  strQ = "SELECT * FROM tb_partido ORDER BY nme_partido ASC"
  objRS2.Open strQ, objCon, , , &H0001
  If Not objRS2.EOF Then
    While Not objRS2.EOF
    If Trim(objRS2.Fields.Item("nme_partido").Value) <> "" Then
      Response.Write "      <OPTION value='" & (objRS2.Fields.Item("id").Value) & "'"
      If Lcase(objRS2.Fields.Item("id").Value) = Lcase(objRS.Fields.Item("cod_partido").Value) then
        Response.Write "selected"
      End If
      Response.Write ">" & (objRS2.Fields.Item("nme_partido").Value) & "</OPTION>"
    End If
      objRS2.MoveNext
    Wend
  End If
  Response.Write("ok")
%>

    </SELECT>

    </TD>
  </TR>
  <TR class=titulo_campos>
    <TD>
      População Urbana - IBGE (2010) (hab)
      <br>
      <%
        If objRS.Fields("qtd_populacao_urbana_2010").properties("IsAutoIncrement") = False Then
      %>
      <INPUT style="width=350" type="text" name="qtd_populacao_urbana_2010" maxlength="255" value="<%=(objRS.Fields.Item("qtd_populacao_urbana_2010").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
      <%
        Else
          Response.Write "<B>(Automático)</B>"
        End If
      %>
    </TD>
  </TR>
  <TR class=titulo_campos>
    <TD>
      Projeção de População (2030)
      <br>
      <%
        If objRS.Fields("qtd_populacao_urbana_2030").properties("IsAutoIncrement") = False Then
      %>
      <INPUT style="width=350" type="text" name="qtd_populacao_urbana_2030" maxlength="255" value="<%=(objRS.Fields.Item("qtd_populacao_urbana_2030").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
      <%
        Else
          Response.Write "<B>(Automático)</B>"
        End If
      %>
    </TD>
  </TR>

  <TR class=titulo_campos>
    <TD>
      Atendido pela sabesp?
      <%
        If objRS.Fields("Observação sobre o Prédio").properties("IsAutoIncrement") = False Then
          value = objRS.Fields.Item("flg_atendido_sabesp").Value
          checked = ""

          If value = True Then
            value = 1
            checked = "checked=""checked"""
          Else
            value = 0
          End If

          Response.Write "<input type=""checkbox"" id=""flg_atendido_sabesp"" name=""flg_atendido_sabesp"" "& checked &"/>"
        Else
          Response.Write "<B>(Automático)</B>"
        End If
      %>
    </TD>
  </TR>

  <TR class=titulo_campos><TD>Concessão<BR>
    <SELECT style="width=350" name="cod_concessao" df_verificar="sim" onChange="desabilita_cor(this)" class=campos_formulario>
        <OPTION value=""></OPTION>

  <%
    Set objRS2 = Server.CreateObject("ADODB.Recordset")
    objRS2.CursorLocation = 3
    objRS2.CursorType = 3
    objRS2.LockType = 1
    strQ = "SELECT * FROM tb_concessao ORDER BY dsc_concessao ASC"
    objRS2.Open strQ, objCon, , , &H0001
    If Not objRS2.EOF Then
      While Not objRS2.EOF
      If Trim(objRS2.Fields.Item("dsc_concessao").Value) <> "" Then
        Response.Write "      <OPTION value='" & (objRS2.Fields.Item("id").Value) & "'"
        If Lcase(objRS2.Fields.Item("id").Value) = Lcase(objRS.Fields.Item("cod_concessao").Value) then
          Response.Write "selected"
        End If
        Response.Write ">" & (objRS2.Fields.Item("dsc_concessao").Value) & "</OPTION>"
      End If
        objRS2.MoveNext
      Wend
    End If
    Response.Write("ok")
  %>

      </SELECT>
    </TD></TR>
    <TR class=titulo_campos><TD>Localização Geográfica (Lat, Long)<br>
<%If objRS.Fields("latitude_longitude").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="latitude_longitude" maxlength="255" value="<%=(objRS.Fields.Item("latitude_longitude").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
<%
Else
  Response.Write "<B>" & (objRS.Fields.Item("latitude_longitude").Value) & "</B>"
End If
%>
  </TD></TR>
  <!-- FIM NOVOS CAMPOS -->

  <TR class=titulo_campos><TD>Observação<br>
<%If objRS.Fields("Observação sobre o Prédio").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="Observação sobre o Prédio" maxlength="255" value="<%=(objRS.Fields.Item("Observação sobre o Prédio").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
<%
Else
  Response.Write "<B>" & (objRS.Fields.Item("Observação sobre o Prédio").Value) & "</B>"
End If
%>
  </TD></TR>

</TABLE>
<input type="submit" name="salvar" value="Enviar" class=botao_enviar>
</form>

<%
    If indice = "" Then
      Response.Write "<BR><B>ATENÇÃO:</B> Crie um campo do tipo <i>AutoIncrement</i> com qualquer nome em sua tabela para evitar erros na alteração dos dados. "
      Response.Write "<a href=""http://www.dataform.com.br/criar_campo_autoincrement.asp"" target=""_blank"">Clique aqui</a> para mais detalhes."
    End If
  End If
End If
%>

</BODY>
</HTML>

<%
Else
  Response.Write "<B>Acesso negado...</B> somente o administrador do site tem acesso a esta página."
  Response.Write "<BR><a href=""" & pagina_login & """ class=""texto_pagina"">Clique aqui</a> para efetuar login no sitema"
End If
%>
