<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_cpf_STRING
  MM_editTable = "tb_pi"
  MM_editColumn = "PI"
  MM_recordId = "'" + Request.Form("MM_recordId") + "'"
  MM_editRedirectUrl = "atualiza_pi_fiscal.asp"
  MM_fieldsStr  = "PI|value|cod_situacao|value|Data_da_abertura|value|Data_do_TRP|value|Data_do_TRD|value"
  MM_columnsStr = "PI|',none,''|cod_situacao|none,none,NULL|[Data da Abertura]|',none,NULL|[Data do TRP]|',none,NULL|[Data do TRD]|',none,NULL"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim rs_mun
Dim rs_mun_numRows

Set rs_mun = Server.CreateObject("ADODB.Recordset")
rs_mun.ActiveConnection = MM_cpf_STRING
rs_mun.Source = "SELECT * FROM tb_Municipios ORDER BY Municipios ASC"
rs_mun.CursorType = 0
rs_mun.CursorLocation = 2
rs_mun.LockType = 1
rs_mun.Open()

rs_mun_numRows = 0
%>
<%
Dim rs_fiscal
Dim rs_fiscal_numRows

Set rs_fiscal = Server.CreateObject("ADODB.Recordset")
rs_fiscal.ActiveConnection = MM_cpf_STRING
rs_fiscal.Source = "SELECT * FROM tb_responsavel ORDER BY Responsável ASC"
rs_fiscal.CursorType = 0
rs_fiscal.CursorLocation = 2
rs_fiscal.LockType = 1
rs_fiscal.Open()

rs_fiscal_numRows = 0
%>
<%
Dim rs_construtora
Dim rs_construtora_numRows

Set rs_construtora = Server.CreateObject("ADODB.Recordset")
rs_construtora.ActiveConnection = MM_cpf_STRING
rs_construtora.Source = "SELECT cod_construtora, Construtora FROM tb_Construtora ORDER BY Construtora ASC"
rs_construtora.CursorType = 0
rs_construtora.CursorLocation = 2
rs_construtora.LockType = 1
rs_construtora.Open()

rs_construtora_numRows = 0
%>
<%
Dim rs_atualiza_pi__MMColParam
rs_atualiza_pi__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  rs_atualiza_pi__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim rs_atualiza_pi
Dim rs_atualiza_pi_numRows

Set rs_atualiza_pi = Server.CreateObject("ADODB.Recordset")
rs_atualiza_pi.ActiveConnection = MM_cpf_STRING
rs_atualiza_pi.Source = "SELECT *  FROM tb_pi  WHERE PI = '" + Replace(rs_atualiza_pi__MMColParam, "'", "''") + "'"
rs_atualiza_pi.CursorType = 0
rs_atualiza_pi.CursorLocation = 2
rs_atualiza_pi.LockType = 1
rs_atualiza_pi.Open()

rs_atualiza_pi_numRows = 0
%>
<%
Dim rs_predio_filtro__MMColParam
rs_predio_filtro__MMColParam = "1"
If (Request.QueryString("cod_predio") <> "") Then 
  rs_predio_filtro__MMColParam = Request.QueryString("cod_predio")
End If
%>
<%
Dim rs_predio_filtro
Dim rs_predio_filtro_numRows

Set rs_predio_filtro = Server.CreateObject("ADODB.Recordset")
rs_predio_filtro.ActiveConnection = MM_cpf_STRING
rs_predio_filtro.Source = "SELECT tb_predio.cod_predio, [cod_predio]+' - '+[Nome_Unidade] AS Expr1  FROM tb_predio  WHERE cod_predio = '" + Replace(rs_predio_filtro__MMColParam, "'", "''") + "'  ORDER BY [cod_predio]+' - '+[Nome_Unidade];"
rs_predio_filtro.CursorType = 0
rs_predio_filtro.CursorLocation = 2
rs_predio_filtro.LockType = 1
rs_predio_filtro.Open()

rs_predio_filtro_numRows = 0
%>
<%
Dim rs_situacao
Dim rs_situacao_numRows

Set rs_situacao = Server.CreateObject("ADODB.Recordset")
rs_situacao.ActiveConnection = MM_cpf_STRING
rs_situacao.Source = "SELECT *  FROM tb_situacao_pi  ORDER BY desc_situacao ASC"
rs_situacao.CursorType = 0
rs_situacao.CursorLocation = 2
rs_situacao.LockType = 1
rs_situacao.Open()

rs_situacao_numRows = 0
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_cpf_STRING
Recordset1.Source = "SELECT tb_predio.cod_predio, [cod_predio]+' - '+[Nome_Unidade] AS Expr1  FROM tb_predio  ORDER BY [cod_predio]+' - '+[Nome_Unidade];  "
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim rs_predio
Dim rs_predio_numRows

Set rs_predio = Server.CreateObject("ADODB.Recordset")
rs_predio.ActiveConnection = MM_cpf_STRING
rs_predio.Source = "SELECT tb_predio.cod_predio, [cod_predio]+' - '+[Nome_Unidade] AS Expr1  FROM tb_predio  ORDER BY [cod_predio]+' - '+[Nome_Unidade];  "
rs_predio.CursorType = 0
rs_predio.CursorLocation = 2
rs_predio.LockType = 1
rs_predio.Open()

rs_predio_numRows = 0
%>
<%
Dim Recordset2__MMColParam
Recordset2__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  Recordset2__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim Recordset2
Dim Recordset2_numRows

Set Recordset2 = Server.CreateObject("ADODB.Recordset")
Recordset2.ActiveConnection = MM_cpf_STRING
Recordset2.Source = "SELECT tb_pi.PI, tb_predio.Município  FROM tb_predio INNER JOIN tb_pi ON tb_predio.cod_predio = tb_pi.cod_predio  WHERE PI = '" + Replace(Recordset2__MMColParam, "'", "''") + "'  ORDER BY tb_predio.Município  "
Recordset2.CursorType = 0
Recordset2.CursorLocation = 2
Recordset2.LockType = 1
Recordset2.Open()

Recordset2_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style1 {
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
	font-size: 10;
}
.style6 {font-family: Arial, Helvetica, sans-serif; font-size: 10; }
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

<body onload="FP_preloadImgs(/*url*/'button55.jpg', /*url*/'button56.jpg', /*url*/'button58.jpg', /*url*/'button59.jpg')">
<form method="POST" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline" bgcolor="#666666">
      <td colspan="6" align="right" nowrap bgcolor="#EBEBEB"><div align="left" class="style1">
        
          <div align="left">
            <select name="select" disabled="disabled" style="font-family: Arial Black; font-size: 8pt; color: #000000; font-weight: bold" size="1">
              <%
While (NOT rs_predio.EOF)
%><option value="<%=(rs_predio.Fields.Item("cod_predio").Value)%>" <%If (Not isNull((rs_atualiza_pi.Fields.Item("cod_predio").Value))) Then If (CStr(rs_predio.Fields.Item("cod_predio").Value) = CStr((rs_atualiza_pi.Fields.Item("cod_predio").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rs_predio.Fields.Item("Expr1").Value)%></option>
              <%
  rs_predio.MoveNext()
Wend
If (rs_predio.CursorType > 0) Then
  rs_predio.MoveFirst
Else
  rs_predio.Requery
End If
%>
            </select>
          </div>
      </div></td>
    </tr>
    <tr valign="baseline" bgcolor="#999999">
      <td width="156" align="right" nowrap="nowrap" bgcolor="#EBEBEB"><span class="style1">Município</span></td>
      <td colspan="5" bgcolor="#EBEBEB"><div align="left" class="style6"><%=(Recordset2.Fields.Item("Município").Value)%></div></td>
    </tr>
    <tr valign="baseline" bgcolor="#999999">
      <td colspan="6" align="right" nowrap="nowrap" bgcolor="#EBEBEB" height="14">
		<a href="OS_alteracao_fiscal.asp?pi=<%=(rs_atualiza_pi.Fields.Item("PI").Value)%>" target="_blank"><img src="button57.jpg" alt="Dados da Os" name="img1" width="100" height="20" border="0" id="img1" onmousedown="FP_swapImg(1,0,/*id*/'img1',/*url*/'button56.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img1',/*url*/'button55.jpg')" onmouseover="FP_swapImg(1,0,/*id*/'img1',/*url*/'button55.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img1',/*url*/'button57.jpg')" fp-style="fp-btn: Embossed Capsule 5" fp-title="Dados da Os"></a></td>
    </tr>
    <tr valign="baseline" bgcolor="#999999">
      <td align="right" nowrap bgcolor="#EBEBEB"><span class="style1">PI</span></td>
      <td width="192" bgcolor="#EBEBEB"><input name="PI" type="text" class="style6" value="<%=(rs_atualiza_pi.Fields.Item("PI").Value)%>" size="20" readonly="true" /></td>
      <td colspan="4" bgcolor="#EBEBEB"><div align="left"><strong><span class="style6">Descrição da Intervenção FDE</span></strong>
            <textarea name="Descrio_da_Interveno_FDE" cols="32" readonly="true" class="style6"><%=(rs_atualiza_pi.Fields.Item("Descrição da Intervenção FDE").Value)%></textarea>
      </div></td>
    </tr>
    <tr valign="baseline" bgcolor="#999999">
      <td align="right" nowrap="nowrap" bgcolor="#EBEBEB"><span class="style1">Construtora</span></td>
      <td colspan="5" bgcolor="#EBEBEB"><select name="cod_construtora" disabled="disabled" class="style6">
        <option value="" <%If (Not isNull(rs_atualiza_pi.Fields.Item("cod_construtora").Value)) Then If ("" = CStr(rs_atualiza_pi.Fields.Item("cod_construtora").Value)) Then Response.Write("selected=""selected""") : Response.Write("")%>></option>
<%
While (NOT rs_construtora.EOF)
%><option value="<%=(rs_construtora.Fields.Item("cod_construtora").Value)%>" <%If (Not isNull(rs_atualiza_pi.Fields.Item("cod_construtora").Value)) Then If (CStr(rs_construtora.Fields.Item("cod_construtora").Value) = CStr(rs_atualiza_pi.Fields.Item("cod_construtora").Value)) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rs_construtora.Fields.Item("Construtora").Value)%></option>
          <%
  rs_construtora.MoveNext()
Wend
If (rs_construtora.CursorType > 0) Then
  rs_construtora.MoveFirst
Else
  rs_construtora.Requery
End If
%>
      </select>
      <div align="right"></div></td>
    </tr>
    <tr valign="baseline" bgcolor="#999999">
      <td align="right" nowrap bgcolor="#EBEBEB"><span class="style1">Número do Contrato</span></td>
      <td bgcolor="#EBEBEB"><input name="Nmero_do_Contrato" type="text" class="style6" value="<%=(rs_atualiza_pi.Fields.Item("Número do Contrato").Value)%>" size="32" readonly="true" /></td>
      <td colspan="4" bgcolor="#EBEBEB"><div align="left"><strong><span class="style6">Dígito do Contrato</span></strong>
        <input name="Dgito_do_Contrato" type="text" class="style6" value="<%=(rs_atualiza_pi.Fields.Item("Dígito do Contrato").Value)%>" size="10" readonly="true" />
        <a target="_blank" href="contratos.asp?pi=<%=(rs_atualiza_pi.Fields.Item("PI").Value)%>">
		<img src="button60.jpg" alt="Dados do Contrato" name="img2" width="125" height="20" border="0" id="img2" onmousedown="FP_swapImg(1,0,/*id*/'img2',/*url*/'button59.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img2',/*url*/'button58.jpg')" onmouseover="FP_swapImg(1,0,/*id*/'img2',/*url*/'button58.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img2',/*url*/'button60.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Dados do Contrato"></a></div></td>
    </tr>
    <tr valign="baseline" bgcolor="#999999">
      <td align="right" nowrap bgcolor="#EBEBEB"><span class="style1">Órgão</span></td>
      <td bgcolor="#EBEBEB"><input name="rgo" type="text" class="style6" value="<%=(rs_atualiza_pi.Fields.Item("Órgão").Value)%>" size="32" readonly="true" /></td>
      <td width="113" bgcolor="#EBEBEB"><div align="right"><strong><span class="style6">Fiscal</span></strong></div></td>
      <td width="90" bgcolor="#EBEBEB"><select name="cod_fiscal" disabled="disabled">
        <option value="" <%If (Not isNull(rs_atualiza_pi.Fields.Item("cod_fiscal").Value)) Then If ("" = CStr(rs_atualiza_pi.Fields.Item("cod_fiscal").Value)) Then Response.Write("selected=""selected""") : Response.Write("")%>></option>
<%
While (NOT rs_fiscal.EOF)
%><option value="<%=(rs_fiscal.Fields.Item("cod_fiscal").Value)%>" <%If (Not isNull(rs_atualiza_pi.Fields.Item("cod_fiscal").Value)) Then If (CStr(rs_fiscal.Fields.Item("cod_fiscal").Value) = CStr(rs_atualiza_pi.Fields.Item("cod_fiscal").Value)) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rs_fiscal.Fields.Item("Responsável").Value)%></option>
          <%
  rs_fiscal.MoveNext()
Wend
If (rs_fiscal.CursorType > 0) Then
  rs_fiscal.MoveFirst
Else
  rs_fiscal.Requery
End If
%>
      </select></td>
      <td width="145" bgcolor="#EBEBEB"><div align="right"><span class="style1">Gerenciadora Mede ?</span></div></td>
      <td width="90" bgcolor="#EBEBEB"><input name="Gerenciadora_Mede_" type="checkbox" disabled="true" value=1 <%If (CStr((rs_atualiza_pi.Fields.Item("Gerenciadora Mede ?").Value)) = CStr("True")) Then Response.Write("checked=""checked""") : Response.Write("")%> /></td>
    </tr>
    <tr valign="baseline" bgcolor="#999999">
      <td align="right" nowrap bgcolor="#EBEBEB"><span class="style1">Est&aacute;gio da Obra </span></td>
      <td colspan="5" bgcolor="#EBEBEB"><label for="select"></label>
        <select name="cod_situacao" id="select">
          <option value="" <%If (Not isNull((rs_atualiza_pi.Fields.Item("cod_situacao").Value))) Then If ("" = CStr((rs_atualiza_pi.Fields.Item("cod_situacao").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>></option>
          <%
While (NOT rs_situacao.EOF)
%><option value="<%=(rs_situacao.Fields.Item("cod_situacao").Value)%>" <%If (Not isNull((rs_atualiza_pi.Fields.Item("cod_situacao").Value))) Then If (CStr(rs_situacao.Fields.Item("cod_situacao").Value) = CStr((rs_atualiza_pi.Fields.Item("cod_situacao").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rs_situacao.Fields.Item("desc_situacao").Value)%></option>
          <%
  rs_situacao.MoveNext()
Wend
If (rs_situacao.CursorType > 0) Then
  rs_situacao.MoveFirst
Else
  rs_situacao.Requery
End If
%>
        </select>        <div align="right"></div></td>
    </tr>
    <tr valign="baseline" bgcolor="#999999">
      <td colspan="6" align="right" nowrap bgcolor="#EBEBEB"><div align="right"></div>
      <div align="left"></div>      <div align="right"></div></td>
    </tr>
    <tr valign="baseline" bgcolor="#999999">
      <td align="right" nowrap bgcolor="#EBEBEB"><span class="style1">Data da Abertura</span></td>
      <td bgcolor="#EBEBEB"><div align="left"><span class="style1">
        <input name="Data_da_abertura" type="text" class="style6" id="Data_da_abertura" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" value="<%=(rs_atualiza_pi.Fields.Item("Data da Abertura").Value)%>" size="15" />
      </span></div></td>
      <td bgcolor="#EBEBEB">
        <div align="left"><span class="style1">Data do TRP</span></div></td><td bgcolor="#EBEBEB"><input name="Data_do_TRP" type="text" class="style6" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" value="<%=(rs_atualiza_pi.Fields.Item("Data do TRP").Value)%>" size="15" /></td>
      <td bgcolor="#EBEBEB"><div align="right"><strong><span class="style6">Data do TRD</span></strong></div></td>
      <td bgcolor="#EBEBEB"><input name="Data_do_TRD" type="text" class="style6" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" value="<%=(rs_atualiza_pi.Fields.Item("Data do TRD").Value)%>" size="15" /></td>
    </tr>
    <tr valign="baseline" bgcolor="#999999">
      <td align="right" nowrap bgcolor="#EBEBEB">&nbsp;</td>
      <td bgcolor="#EBEBEB">&nbsp;</td>
      <td bgcolor="#EBEBEB"><div align="right"></div></td>
      <td bgcolor="#EBEBEB">&nbsp;</td>
      <td bgcolor="#EBEBEB"><a href="OS_alteracao.asp?pi=<%=(rs_atualiza_pi.Fields.Item("PI").Value)%>"></a></td>
      <td bgcolor="#EBEBEB">&nbsp;</td>
    </tr>
    <tr valign="baseline" bgcolor="#999999">
      <td align="right" nowrap bgcolor="#EBEBEB">&nbsp;</td>
      <td bgcolor="#EBEBEB"><input type="submit" value="Salvar">      </td>
      <td bgcolor="#EBEBEB"><div align="right"></div></td>
      <td bgcolor="#EBEBEB">&nbsp;</td>
      <td bgcolor="#EBEBEB">&nbsp;</td>
      <td bgcolor="#EBEBEB">&nbsp;</td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= rs_atualiza_pi.Fields.Item("PI").Value %>">
</form>
<p>&nbsp;</p>
</body>
</html>
<%
rs_mun.Close()
Set rs_mun = Nothing
%>
<%
rs_fiscal.Close()
Set rs_fiscal = Nothing
%>
<%
rs_construtora.Close()
Set rs_construtora = Nothing
%>
<%
rs_atualiza_pi.Close()
Set rs_atualiza_pi = Nothing
%>
<%
rs_predio_filtro.Close()
Set rs_predio_filtro = Nothing
%>
<%
rs_situacao.Close()
Set rs_situacao = Nothing
%>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
rs_predio.Close()
Set rs_predio = Nothing
%>
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>