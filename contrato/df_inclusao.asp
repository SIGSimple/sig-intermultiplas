<% @ LANGUAGE="VBSCRIPT" %>
<%
'*******************************************************************
' Página gerada pelo sistema Dataform 2 - http://www.dataform.com.br
'*******************************************************************
' Altere os valores das variáveis indicadas abaixo se necessário

'String de conexão para o banco de dados do Microsoft Access
'strCon = "DBQ=C:\Inetpub\wwwroot\fde\Dados.mdb;Driver={Microsoft Access Driver (*.mdb)};"
strCon = "DBQ=d:\dominios\terkoisolacao.com\www\DB\fde.mdb;Driver={Microsoft Access Driver (*.mdb)};"

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
<TITLE>Incluir Registro</TITLE>
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
If Not IsEmpty(Request.Form) Then
  Set objCon = Server.CreateObject("ADODB.Connection")
  objCon.Open MM_cpf_STRING

  campo_duplicado = false
  campo_msg = ""

  If campo_duplicado = false Then
    Set objRS= Server.CreateObject("ADODB.Recordset")
    objRS.CursorLocation = 3
    objRS.CursorType = 0
    objRS.LockType = 3
    strQ = "SELECT * FROM tb_contrato Where 1 <> 1"
    objRS.Open strQ, objCon, , , &H0001
    objRS.Addnew()
    If objRS.Fields("area_construida").properties("IsAutoIncrement") = False Then
      objRS("area_construida") = Trim(Request.Form("area_construida"))
    End If
    If objRS.Fields("area_gerenciadora").properties("IsAutoIncrement") = False Then
      objRS("area_gerenciadora") = Trim(Request.Form("area_gerenciadora"))
    End If
    If objRS.Fields("cod_contrato").properties("IsAutoIncrement") = False Then
      objRS("cod_contrato") = Trim(Request.Form("cod_contrato"))
    End If
    If objRS.Fields("cod_predio").properties("IsAutoIncrement") = False Then
      objRS("cod_predio") = Trim(Request.Form("cod_predio"))
    End If
    If objRS.Fields("cod_reg").properties("IsAutoIncrement") = False Then
      objRS("cod_reg") = Trim(Request.Form("cod_reg"))
    End If
    If objRS.Fields("Construtora").properties("IsAutoIncrement") = False Then
      objRS("Construtora") = Trim(Request.Form("Construtora"))
    End If
    If objRS.Fields("crit_calculo").properties("IsAutoIncrement") = False Then
      objRS("crit_calculo") = Trim(Request.Form("crit_calculo"))
    End If
    If objRS.Fields("crit_reajuste").properties("IsAutoIncrement") = False Then
      objRS("crit_reajuste") = Trim(Request.Form("crit_reajuste"))
    End If
    If objRS.Fields("desc_interv_gerenciadora").properties("IsAutoIncrement") = False Then
      objRS("desc_interv_gerenciadora") = Trim(Request.Form("desc_interv_gerenciadora"))
    End If
    If objRS.Fields("Desc_intervencao_FDE").properties("IsAutoIncrement") = False Then
      objRS("Desc_intervencao_FDE") = Trim(Request.Form("Desc_intervencao_FDE"))
    End If
    If objRS.Fields("dig_contrato").properties("IsAutoIncrement") = False Then
      objRS("dig_contrato") = Trim(Request.Form("dig_contrato"))
    End If
    If objRS.Fields("dt_abertura").properties("IsAutoIncrement") = False Then
      objRS("dt_abertura") = Trim(Request.Form("dt_abertura"))
    End If
    If objRS.Fields("dt_assinatura").properties("IsAutoIncrement") = False Then
      objRS("dt_assinatura") = Trim(Request.Form("dt_assinatura"))
    End If
    If objRS.Fields("dt_base").properties("IsAutoIncrement") = False Then
      objRS("dt_base") = Trim(Request.Form("dt_base"))
    End If
    If objRS.Fields("dt_CI").properties("IsAutoIncrement") = False Then
      objRS("dt_CI") = Trim(Request.Form("dt_CI"))
    End If
    If objRS.Fields("dt_impressao_ois").properties("IsAutoIncrement") = False Then
      objRS("dt_impressao_ois") = Trim(Request.Form("dt_impressao_ois"))
    End If
    If objRS.Fields("dt_termino").properties("IsAutoIncrement") = False Then
      objRS("dt_termino") = Trim(Request.Form("dt_termino"))
    End If
    If objRS.Fields("dt_TRD").properties("IsAutoIncrement") = False Then
      objRS("dt_TRD") = Trim(Request.Form("dt_TRD"))
    End If
    If objRS.Fields("dt_TRP").properties("IsAutoIncrement") = False Then
      objRS("dt_TRP") = Trim(Request.Form("dt_TRP"))
    End If
    If objRS.Fields("fator_reducao").properties("IsAutoIncrement") = False Then
      objRS("fator_reducao") = Trim(Request.Form("fator_reducao"))
    End If
    If objRS.Fields("fical").properties("IsAutoIncrement") = False Then
      objRS("fical") = Trim(Request.Form("fical"))
    End If
    If objRS.Fields("gerenc_mede").properties("IsAutoIncrement") = False Then
      If Request.Form("gerenc_mede") <> "" Then
        objRS("gerenc_mede") = true
      Else
        objRS("gerenc_mede") = false
      End If
    End If
    If objRS.Fields("Obra_pi").properties("IsAutoIncrement") = False Then
      If Request.Form("Obra_pi") <> "" Then
        objRS("Obra_pi") = true
      Else
        objRS("Obra_pi") = false
      End If
    End If
    If objRS.Fields("orc_FDE").properties("IsAutoIncrement") = False Then
      objRS("orc_FDE") = Trim(Request.Form("orc_FDE"))
    End If
    If objRS.Fields("orgao").properties("IsAutoIncrement") = False Then
      objRS("orgao") = Trim(Request.Form("orgao"))
    End If
    If objRS.Fields("PI").properties("IsAutoIncrement") = False Then
      objRS("PI") = Trim(Request.Form("PI"))
    End If
    If objRS.Fields("prz_contrato").properties("IsAutoIncrement") = False Then
      objRS("prz_contrato") = Trim(Request.Form("prz_contrato"))
    End If
    If objRS.Fields("reducao").properties("IsAutoIncrement") = False Then
      objRS("reducao") = Trim(Request.Form("reducao"))
    End If
    If objRS.Fields("solic_aditamento").properties("IsAutoIncrement") = False Then
      If Request.Form("solic_aditamento") <> "" Then
        objRS("solic_aditamento") = true
      Else
        objRS("solic_aditamento") = false
      End If
    End If
    If objRS.Fields("vl_contrato").properties("IsAutoIncrement") = False Then
      objRS("vl_contrato") = Trim(Request.Form("vl_contrato"))
    End If
    objRS.Update
    objRS.Close
    Set objRS = Nothing
%>

<BR><B>Registro salvo</B><BR>O registro foi cadastrado 
com sucesso.<BR><BR>

<%
  Else
%>

    <BR><B>Atenção!</B><BR><BR>
    O campo <%=campo_msg%> <i>"<%=valor_msg%>"</i> não pode ser salvo, pois já está cadastrado na tabela.
    <BR><A href="javascript:history.go(-1)">Clique aqui</a> para voltar ao cadastro</A><BR><BR>

<%
  End If

  objCon.Close
  Set objCon = Nothing
Else
  Set objCon = Server.CreateObject("ADODB.Connection")
  objCon.Open MM_cpf_STRING

  Set objRS= Server.CreateObject("ADODB.Recordset")
  objRS.CursorLocation = 2
  objRS.CursorType = 0
  objRS.LockType = 3
  strQ = "SELECT * FROM tb_contrato Where 1 <> 1"
  objRS.Open strQ, objCon, , , &H0001
%>

<B>Incluir um novo registro</B><BR>Preencha corretamente 
os dados abaixo:<BR>
<form name="form_incluir" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" onSubmit="return verifica_form(this);">
<TABLE border=0 cellpadding=2 cellspacing=1 class=tabela_formulario>
  <TR class=titulo_campos><TD>area_construida<br>
<%If objRS.Fields("area_construida").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="area_construida" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>area_gerenciadora<br>
<%If objRS.Fields("area_gerenciadora").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="area_gerenciadora" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>cod_contrato<br>
<%If objRS.Fields("cod_contrato").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="cod_contrato" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>cod_predio<br>
<%If objRS.Fields("cod_predio").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="cod_predio" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>cod_reg<br>
<%If objRS.Fields("cod_reg").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="cod_reg" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>Construtora<BR>
    <SELECT style="width=350" name="Construtora" df_verificar="sim" onChange="desabilita_cor(this)" class=campos_formulario>
      <OPTION value=""></OPTION>

<%
  Set objRS2 = Server.CreateObject("ADODB.Recordset")
  objRS2.CursorLocation = 3
  objRS2.CursorType = 3
  objRS2.LockType = 1
  strQ = "SELECT Construtora FROM tb_Construtora ORDER BY Construtora ASC"
  objRS2.Open strQ, objCon, , , &H0001
  If Not objRS2.EOF Then
    While Not objRS2.EOF
	  If Trim(objRS2.Fields.Item("Construtora").Value) <> "" Then
	    Response.Write "      <OPTION value='" & (objRS2.Fields.Item("Construtora").Value) & "'>" & (objRS2.Fields.Item("Construtora").Value) & "</OPTION>"
	  End If
      objRS2.MoveNext
    Wend
  End If
  Response.Write("ok")
%>

    </SELECT>
  </TD></TR>
  <TR class=titulo_campos><TD>crit_calculo<br>
<%If objRS.Fields("crit_calculo").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="crit_calculo" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>crit_reajuste<br>
<%If objRS.Fields("crit_reajuste").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="crit_reajuste" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>desc_interv_gerenciadora<br>
<%If objRS.Fields("desc_interv_gerenciadora").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="desc_interv_gerenciadora" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>Desc_intervencao_FDE<br>
<%If objRS.Fields("Desc_intervencao_FDE").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="Desc_intervencao_FDE" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>dig_contrato<br>
<%If objRS.Fields("dig_contrato").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="dig_contrato" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>dt_abertura<br>
<%If objRS.Fields("dt_abertura").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="dt_abertura" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>dt_assinatura<br>
<%If objRS.Fields("dt_assinatura").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="dt_assinatura" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>dt_base<br>
<%If objRS.Fields("dt_base").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="dt_base" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>dt_CI<br>
<%If objRS.Fields("dt_CI").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="dt_CI" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>dt_impressao_ois<br>
<%If objRS.Fields("dt_impressao_ois").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="dt_impressao_ois" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>dt_termino<br>
<%If objRS.Fields("dt_termino").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="dt_termino" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>dt_TRD<br>
<%If objRS.Fields("dt_TRD").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="dt_TRD" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>dt_TRP<br>
<%If objRS.Fields("dt_TRP").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="dt_TRP" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>fator_reducao<br>
<%If objRS.Fields("fator_reducao").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="fator_reducao" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>fical<br>
<%If objRS.Fields("fical").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="fical" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD><input name="gerenc_mede" type="checkbox">gerenc_mede
  </TD></TR>
  <TR class=titulo_campos><TD><input name="Obra_pi" type="checkbox">Obra_pi
  </TD></TR>
  <TR class=titulo_campos><TD>orc_FDE<br>
<%If objRS.Fields("orc_FDE").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="orc_FDE" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>orgao<br>
<%If objRS.Fields("orgao").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="orgao" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>PI<br>
<%If objRS.Fields("PI").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="PI" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>prz_contrato<br>
<%If objRS.Fields("prz_contrato").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="prz_contrato" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD>reducao<br>
<%If objRS.Fields("reducao").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="reducao" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
  <TR class=titulo_campos><TD><input name="solic_aditamento" type="checkbox">solic_aditamento
  </TD></TR>
  <TR class=titulo_campos><TD>vl_contrato<br>
<%If objRS.Fields("vl_contrato").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="vl_contrato" maxlength="255" value="" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>(Automático)</B>"
End If
%>
  </TD></TR>
</TABLE>
<input type="submit" name="submit" value="Enviar" class=botao_enviar>
</form>

<%
  objCon.Close
  Set objCon = Nothing
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
