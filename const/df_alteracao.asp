<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/cpf.asp" -->
<%
'*******************************************************************
' Página gerada pelo sistema Dataform 2 - http://www.dataform.com.br
'*******************************************************************
' Altere os valores das variáveis indicadas abaixo se necessário

'String de conexão para o banco de dados do Microsoft Access
strCon = "DBQ=C:\inetpub\wwwroot\original\ARQUIVOS\DADOS\bd_fde.mdb;Driver={Microsoft Access Driver (*.mdb)};"

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
<b>Cadastro de Empresas</b> | Links: <a href="<%=pagina_consulta%>" class="texto_pagina">Página de Consulta</a> | <a href="<%=pagina_inclusao%>" class="texto_pagina">Página de Inclusão<hr size=1 color=gainsboro></a>

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
  If indice <> "" Then strQ = " SELECT * FROM tb_Construtora WHERE " & indice

  objRS.Open strQ, objCon, , , &H0001
  If indice = "" Then objRS.Move Request.Form("recordno") - 1

  If objRS.Fields("cod_construtora").properties("IsAutoIncrement") = False Then
    objRS("cod_construtora") = Trim(Request.Form("cod_construtora"))
  End If
  If objRS.Fields("Construtora").properties("IsAutoIncrement") = False Then
    objRS("Construtora") = Trim(Request.Form("Construtora"))
  End If
  If objRS.Fields("Endereço da Construtora").properties("IsAutoIncrement") = False Then
    objRS("Endereço da Construtora") = Trim(Request.Form("Endereço da Construtora"))
  End If
  If objRS.Fields("cep_empresa").properties("IsAutoIncrement") = False Then
    objRS("cep_empresa") = Trim(Request.Form("cep_empresa"))
  End If
  If objRS.Fields("cod_municipio_empresa").properties("IsAutoIncrement") = False Then
    objRS("cod_municipio_empresa") = Trim(Request.Form("cod_municipio_empresa"))
  End If
  If objRS.Fields("email_empresa").properties("IsAutoIncrement") = False Then
    objRS("email_empresa") = Trim(Request.Form("email_empresa"))
  End If
  If objRS.Fields("site_empresa").properties("IsAutoIncrement") = False Then
    objRS("site_empresa") = Trim(Request.Form("site_empresa"))
  End If
  If objRS.Fields("cnpj_empresa").properties("IsAutoIncrement") = False Then
    objRS("cnpj_empresa") = Trim(Request.Form("cnpj_empresa"))
  End If
  If objRS.Fields("Fone da Construtora").properties("IsAutoIncrement") = False Then
    objRS("Fone da Construtora") = Trim(Request.Form("Fone da Construtora"))
  End If
  If objRS.Fields("Engenheiro responsável").properties("IsAutoIncrement") = False Then
    objRS("Engenheiro responsável") = Trim(Request.Form("Engenheiro responsável"))
  End If
  If objRS.Fields("Número do CREA").properties("IsAutoIncrement") = False Then
    objRS("Número do CREA") = Trim(Request.Form("Número do CREA"))
  End If
  If objRS.Fields("telefone_responsavel").properties("IsAutoIncrement") = False Then
    objRS("telefone_responsavel") = Trim(Request.Form("telefone_responsavel"))
  End If
  If objRS.Fields("email_responsavel").properties("IsAutoIncrement") = False Then
    objRS("email_responsavel") = Trim(Request.Form("email_responsavel"))
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
  If indice <> "" Then strQ = " SELECT * FROM tb_Construtora WHERE " & indice

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
  <TR class=titulo_campos><TD>cod_empresa<br>
<%If objRS.Fields("cod_construtora").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="cod_construtora" maxlength="255" value="<%=(objRS.Fields.Item("cod_construtora").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
<%
Else
  Response.Write "<B>" & (objRS.Fields.Item("cod_construtora").Value) & "</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>Empresa<br>
<%If objRS.Fields("Construtora").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="Construtora" maxlength="255" value="<%=(objRS.Fields.Item("Construtora").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
<%
Else
  Response.Write "<B>" & (objRS.Fields.Item("Construtora").Value) & "</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>Endereço da Empresa<br>
<%If objRS.Fields("Endereço da Construtora").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="Endereço da Construtora" maxlength="255" value="<%=(objRS.Fields.Item("Endereço da Construtora").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
<%
Else
  Response.Write "<B>" & (objRS.Fields.Item("Endereço da Construtora").Value) & "</B>"
End If
%>
  </TD></TR>

  <TR class=titulo_campos><TD>CEP da Empresa<br>
<%If objRS.Fields("cep_empresa").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="cep_empresa" maxlength="255" value="<%=(objRS.Fields.Item("cep_empresa").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
<%
Else
  Response.Write "<B>" & (objRS.Fields.Item("cep_empresa").Value) & "</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>Município da Empresa<br>
<%If objRS.Fields("cod_municipio_empresa").properties("IsAutoIncrement") = False Then%>
    <SELECT style="width=350" name="cod_municipio_empresa" df_verificar="sim" onChange="desabilita_cor(this)" class=campos_formulario>
      <OPTION value=""></OPTION>

<%
  Set objRS2 = Server.CreateObject("ADODB.Recordset")
  objRS2.CursorLocation = 3
  objRS2.CursorType = 3
  objRS2.LockType = 1
  strQ = "SELECT cod_mun, Municipios FROM tb_Municipios ORDER BY Municipios ASC"
  objRS2.Open strQ, objCon, , , &H0001
  If Not objRS2.EOF Then
    While Not objRS2.EOF
    If Trim(objRS2.Fields.Item("Municipios").Value) <> "" Then
      Response.Write "      <OPTION value='" & (objRS2.Fields.Item("cod_mun").Value) & "'"
      If Lcase(objRS2.Fields.Item("cod_mun").Value) = Lcase(objRS.Fields.Item("cod_municipio_empresa").Value) then
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
<%
Else
  Response.Write "<B>" & (objRS.Fields.Item("cod_municipio_empresa").Value) & "</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>Email da Empresa<br>
<%If objRS.Fields("email_empresa").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="email_empresa" maxlength="255" value="<%=(objRS.Fields.Item("email_empresa").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
<%
Else
  Response.Write "<B>" & (objRS.Fields.Item("email_empresa").Value) & "</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>Site da Empresa<br>
<%If objRS.Fields("site_empresa").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="site_empresa" maxlength="255" value="<%=(objRS.Fields.Item("site_empresa").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
<%
Else
  Response.Write "<B>" & (objRS.Fields.Item("site_empresa").Value) & "</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>Telefone da Empresa<br>
<%If objRS.Fields("Fone da Construtora").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="Fone da Construtora" maxlength="255" value="<%=(objRS.Fields.Item("Fone da Construtora").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
<%
Else
  Response.Write "<B>" & (objRS.Fields.Item("Fone da Construtora").Value) & "</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>CNPJ da Empresa<br>
<%If objRS.Fields("cnpj_empresa").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="cnpj_empresa" maxlength="255" value="<%=(objRS.Fields.Item("cnpj_empresa").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
<%
Else
  Response.Write "<B>" & (objRS.Fields.Item("cnpj_empresa").Value) & "</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>Engenheiro responsável<br>
<%If objRS.Fields("Engenheiro responsável").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="Engenheiro responsável" maxlength="255" value="<%=(objRS.Fields.Item("Engenheiro responsável").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
<%
Else
  Response.Write "<B>" & (objRS.Fields.Item("Engenheiro responsável").Value) & "</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>Número do CREA<br>
<%If objRS.Fields("Número do CREA").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="Número do CREA" maxlength="255" value="<%=(objRS.Fields.Item("Número do CREA").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
<%
Else
  Response.Write "<B>" & (objRS.Fields.Item("Número do CREA").Value) & "</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>Telefone responsável<br>
<%If objRS.Fields("telefone_responsavel").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="telefone_responsavel" maxlength="255" value="<%=(objRS.Fields.Item("telefone_responsavel").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
<%
Else
  Response.Write "<B>" & (objRS.Fields.Item("telefone_responsavel").Value) & "</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>Email responsável<br>
<%If objRS.Fields("email_responsavel").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="email_responsavel" maxlength="255" value="<%=(objRS.Fields.Item("email_responsavel").Value)%>" onKeyPress="desabilita_cor(this)" class=campos_formulario>
<%
Else
  Response.Write "<B>" & (objRS.Fields.Item("email_responsavel").Value) & "</B>"
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
