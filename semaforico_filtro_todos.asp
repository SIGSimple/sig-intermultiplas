<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_cpf_STRING
Recordset1_cmd.CommandText = "SELECT * FROM c_Semaforico" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Recordset2
Dim Recordset2_cmd
Dim Recordset2_numRows

Set Recordset2_cmd = Server.CreateObject ("ADODB.Command")
Recordset2_cmd.ActiveConnection = MM_cpf_STRING
Recordset2_cmd.CommandText = "SELECT * FROM tb_situacao_pi" 
Recordset2_cmd.Prepared = true

Set Recordset2 = Recordset2_cmd.Execute
Recordset2_numRows = 0
%>
<%
Dim Recordset3__MMColParam
Recordset3__MMColParam = "1"
If (Request.Form("cod_situacao") <> "") Then 
  Recordset3__MMColParam = Request.Form("cod_situacao")
End If
%>
<%
Dim Recordset3
Dim Recordset3_cmd
Dim Recordset3_numRows

Set Recordset3_cmd = Server.CreateObject ("ADODB.Command")
Recordset3_cmd.ActiveConnection = MM_cpf_STRING
Recordset3_cmd.CommandText = "SELECT * FROM tb_situacao_pi WHERE cod_situacao = ?" 
Recordset3_cmd.Prepared = true
Recordset3_cmd.Parameters.Append Recordset3_cmd.CreateParameter("param1", 5, 1, -1, Recordset3__MMColParam) ' adDouble

Set Recordset3 = Recordset3_cmd.Execute
Recordset3_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>RELAT&Oacute;RIO SEMAF&Oacute;RICO</title>
<style type="text/css">
<!--
.style3 {font-family: Arial, Helvetica, sans-serif; font-size: 9px; }
.style5 {font-family: Arial, Helvetica, sans-serif; font-size: 9px; color: #FFFFFF; }
.style8 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 14px;
	font-weight: bold;
	color: #0000CC;
}
.style10 {
	font-size: 16px;
	font-family: Arial, Helvetica, sans-serif;
	color: #990000;
}
.style12 {
	font-size: 12px;
	font-family: Arial, Helvetica, sans-serif;
	color: #990000;
	font-weight: bold;
}
-->
</style>
<script language="JavaScript">
<!--
function FP_preloadImgs() {//v1.0
 var d=document,a=arguments; if(!d.FP_imgs) d.FP_imgs=new Array();
 for(var i=0; i<a.length; i++) { d.FP_imgs[i]=new Image; d.FP_imgs[i].src=a[i]; }
}

function FP_swapImg() {//v1.0
 var doc=document,args=arguments,elm,n; doc.$imgSwaps=new Array(); for(n=2; n<args.length;
 n+=2) { elm=FP_getObjectByID(args[n]); if(elm) { doc.$imgSwaps[doc.$imgSwaps.length]=elm;
 elm.$src=elm.src; elm.src=args[n+1]; } }
}

function FP_getObjectByID(id,o) {//v1.0
 var c,el,els,f,m,n; if(!o)o=document; if(o.getElementById) el=o.getElementById(id);
 else if(o.layers) c=o.layers; else if(o.all) el=o.all[id]; if(el) return el;
 if(o.id==id || o.name==id) return o; if(o.childNodes) c=o.childNodes; if(c)
 for(n=0; n<c.length; n++) { el=FP_getObjectByID(id,c[n]); if(el) return el; }
 f=o.forms; if(f) for(n=0; n<f.length; n++) { els=f[n].elements;
 for(m=0; m<els.length; m++){ el=FP_getObjectByID(id,els[n]); if(el) return el; } }
 return null;
}
// -->
</script>
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
</head>

<body onload="FP_preloadImgs(/*url*/'button389.jpg',/*url*/'button390.jpg')">
<form id="form1" name="form1" method="post" action="">
  <label for="dt_registro"></label>
  <table width="994" border="0">
    <tr bgcolor="#CCCCCC">
      <td colspan="2" class="style10">RELAT&Oacute;RIO DE OBRAS - <strong>SEMAF&Oacute;RICO</strong><strong class="style10"> - TODOS</strong></td>
      <td width="319"><div align="right"><img src="imagens/logo.JPG" width="318" height="41" /></div></td>
    </tr>

    <tr>
      <td width="371">&nbsp;</td>
      <td width="290"><a href="saida_excel.asp">
		<img border="0" id="img1" src="button391.jpg" height="20" width="100" alt="Saída Excel" onmouseover="FP_swapImg(1,0,/*id*/'img1',/*url*/'button389.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img1',/*url*/'button391.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img1',/*url*/'button390.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img1',/*url*/'button389.jpg')" fp-style="fp-btn: Embossed Capsule 5" fp-title="Saída Excel"></a></td>
      <td><div align="center"><a href="inicio.asp"><img src="button141.jpg" width="100" height="20" /></a></div></td>
    </tr>
  </table>
  <script language="JavaScript">document.write("<font color='#GGGGGG' size='1' face='arial'>") 
var mydate=new Date()
var year=mydate.getYear()
if (year<2000)
year += (year < 1900) ? 1900 : 0
var day=mydate.getDay()
var month=mydate.getMonth()
var daym=mydate.getDate()
if (daym<10)
daym="0"+daym
var dayarray=new Array("Domingo","Segunda-feira","Terça-feira","Quarta-feira","Quinta-feira","Sexta-feira","Sábado")
var montharray=new Array(" de Janeiro de "," de Fevereiro de "," de Março de ","de Abril de ","de Maio de ","de Junho de","de Julho de ","de Agosto de ","de Setembro de "," de Outubro de "," de Novembro de "," de Dezembro de ")
document.write("   "+dayarray[day]+", "+daym+" "+montharray[month]+year+" ")
document.write("</b></i></font>")
</script>
</form>
<table width="2846" border="0">
  <tr bgcolor="#999999">
    <td width="60" rowspan="2"><span class="style5">PI-&Iacute;tem</span></td>
    <td width="30" rowspan="2"><span class="style5">C&oacute;digo do Pr&eacute;dio</span></td>
    <td width="30" rowspan="2"><span class="style5">Nome da Unidade </span></td>
    <td width="30" rowspan="2"><span class="style5">Munic&iacute;pio</span></td>
    <td width="100" rowspan="2"><span class="style5">Construtora</span></td>
    <td width="60" rowspan="2"><span class="style5">Fiscal</span></td>
    <td width="30" rowspan="2"><span class="style5">&Oacute;rg&atilde;o</span></td>
    <td width="30" rowspan="2"><span class="style5">Iternven&ccedil;&atilde;o FDE </span></td>
    <td colspan="3"><div align="center"><span class="style5">Valor (x1.000) / % Financeiro</span></div></td>
    <td width="30" rowspan="2"><span class="style5">N&ordm; &Uacute;ltima Medi&ccedil;&atilde;o </span></td>
    <td width="30" rowspan="2"><span class="style5">Data da Abertura </span></td>
    <td width="30" rowspan="2"><span class="style5">Prazo do Contrato </span></td>
    <td width="30" rowspan="2"><span class="style5">Dias Corridos &agrave; partir da OIS </span></td>
    <td width="30" rowspan="2"><span class="style5">Dias Corridos X Prazo Contratual </span></td>
    <td width="30" rowspan="2"><span class="style5">% Avan&ccedil;o F&iacute;sico Anterior</span></td>
    <td width="30" rowspan="2"><span class="style5">% Avan&ccedil;o F&iacute;sico Atual</span></td>
    <td width="30" rowspan="2"><span class="style5">T&eacute;rmino Contratual </span></td>
    <td width="30" rowspan="2"><span class="style5">Previs&atilde;o de T&eacute;rmino </span></td>
    <td width="34" rowspan="2"><span class="style5">Data do TRP </span></td>
    <td width="30" rowspan="2"><p class="style5">Medi&ccedil;&atilde;o Final SIM/N&Atilde;O </p></td>
    <td width="34" rowspan="2"><span class="style5">Data do TRD </span></td>
    <td width="246" rowspan="2"><span class="style5">A&ccedil;&otilde;es</span></td>
  </tr>
  <tr bgcolor="#999999">
    <td width="95">&nbsp;</td>
    <td width="30"><span class="style5">Anterior</span></td>
    <td width="30"><span class="style5">Atual</span></td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) %>
  <tr bgcolor="#CCCCCC">
    <td class="style3"><%=(Recordset1.Fields.Item("PI-item").Value)%></td>
    <td class="style3"><%=(Recordset1.Fields.Item("cod_predio").Value)%></td>
    <td class="style3"><%=(Recordset1.Fields.Item("Nome_Unidade").Value)%></td>
    <td class="style3"><%=(Recordset1.Fields.Item("Município").Value)%></td>
    <td class="style3"><%=(Recordset1.Fields.Item("Construtora").Value)%></td>
    <td class="style3"><%=(Recordset1.Fields.Item("fiscal").Value)%></td>
    <td class="style3"><%=(Recordset1.Fields.Item("Órgão").Value)%></td>
    <td class="style3"><%=(Recordset1.Fields.Item("Descrição da Intervenção FDE").Value)%></td>
    <td class="style3"><div align="right"><%= FormatNumber((Recordset1.Fields.Item("valor_mil").Value), 2, -2, -2, -2) %></div></td>
    <td class="style3"><div align="right"><%= FormatPercent((Recordset1.Fields.Item("medicao_anterior").Value), 2, -2, -2, -2) %></div></td>
    <td class="style3"><div align="right"><%= FormatPercent((Recordset1.Fields.Item("medicaoatual").Value), 2, -2, -2, -2) %></div></td>
    <td class="style3"><%=(Recordset1.Fields.Item("nº da Última Medição").Value)%></td>
    <td class="style3"><%=(Recordset1.Fields.Item("Data da Abertura").Value)%></td>
    <td class="style3"><%=(Recordset1.Fields.Item("Prazo Contratual (dias)").Value)%></td>
    <td class="style3"><div align="right"><%= Round((Recordset1.Fields.Item("Dias Corridos a partir da OIS").Value)) %></div></td>
    <td class="style3"><div align="right"><%= FormatPercent((Recordset1.Fields.Item("Dias Corridos X Prazo Contratual").Value), 2, -2, -2, -2) %></div></td>
    <td class="style3"><%= FormatPercent((Recordset1.Fields.Item("medicao_anterior").Value), 2, -2, -2, -2) %></td>
    <td class="style3"><div align="right"><%=(Recordset1.Fields.Item("Avanço Físico Atual").Value)%></div></td>
    <td class="style3"><%=(Recordset1.Fields.Item("Término Contratual").Value)%></td>
    <td class="style3"><%=(Recordset1.Fields.Item("Previsão de Término Fiscal").Value)%></td>
    <td class="style3"><%=(Recordset1.Fields.Item("Data do TRP").Value)%></td>
    <td class="style3"><%=(Recordset1.Fields.Item("Med_final").Value)%></td>
    <td class="style3"><%=(Recordset1.Fields.Item("Data do TRD").Value)%></td>
    <td class="style3"><%=(Recordset1.Fields.Item("ações").Value)%></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
</table>
<p>&nbsp;</p>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>
<%
Recordset3.Close()
Set Recordset3 = Nothing
%>