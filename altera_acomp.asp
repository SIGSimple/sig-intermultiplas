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
  MM_editTable = "tb_Acompanhamento"
  MM_editColumn = "cod_acompanhamento"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "altera_acomp.asp"
  MM_fieldsStr  = "Data_do_Registro|value|Registro|value|Responsvel|value|vistoria|value|dt_vistoria|value|flg_pendencia|value|cod_tipo_pendencia|value|dsc_pendencia|value|dta_limite_pendencia|value|cod_situacao_sso|value|cod_situacao_clima_manha|value|cod_situacao_clima_tarde|value|cod_situacao_clima_noite|value|cod_situacao_limpeza_obra|value|cod_situacao_organizacao_obra|value|flg_dia_perdido|value|flg_dia_trabalhado|value|dsc_nota_clima_manha|value|dsc_nota_clima_tarde|value|dsc_nota_clima_noite|value|dsc_nota_limpeza_obra|value|dsc_nota_organizacao_obra|value|dsc_nota_dia_perdido|value|dsc_nota_dia_trabalhado|value"
  MM_columnsStr = "[Data do Registro]|',none,NULL|Registro|',none,''|cod_fiscal|none,none,NULL|vistoria|none,1,0|dt_vistoria|',none,NULL|flg_pendencia|none,1,0|cod_tipo_pendencia|none,none,NULL|dsc_pendencia|',none,''|dta_limite_pendencia|',none,NULL|cod_situacao_sso|none,none,NULL|cod_situacao_clima_manha|none,none,NULL|cod_situacao_clima_tarde|none,none,NULL|cod_situacao_clima_noite|none,none,NULL|cod_situacao_limpeza_obra|none,none,NULL|cod_situacao_organizacao_obra|none,none,NULL|flg_dia_perdido|none,1,0|flg_dia_trabalhado|none,1,0|dsc_nota_clima_manha|',none,''|dsc_nota_clima_tarde|',none,''|dsc_nota_clima_noite|',none,''|dsc_nota_limpeza_obra|',none,''|dsc_nota_organizacao_obra|',none,''|dsc_nota_dia_perdido|',none,''|dsc_nota_dia_trabalhado|',none,''"

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
Dim rs_altera_acomp__MMColParam
rs_altera_acomp__MMColParam = "1"
If (Request.QueryString("cod_acompanhamento") <> "") Then 
  rs_altera_acomp__MMColParam = Request.QueryString("cod_acompanhamento")
End If
%>
<%
Dim rs_altera_acomp
Dim rs_altera_acomp_numRows

Set rs_altera_acomp = Server.CreateObject("ADODB.Recordset")
rs_altera_acomp.ActiveConnection = MM_cpf_STRING
rs_altera_acomp.Source = "SELECT *  FROM c_Acomp_resp  WHERE cod_acompanhamento = " + Replace(rs_altera_acomp__MMColParam, "'", "''") + ""
rs_altera_acomp.CursorType = 0
rs_altera_acomp.CursorLocation = 2
rs_altera_acomp.LockType = 1
rs_altera_acomp.Open()

rs_altera_acomp_numRows = 0
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_cpf_STRING
Recordset1.Source = "SELECT * FROM tb_responsavel ORDER BY Respons�vel ASC"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rs_altera_acomp_numRows = rs_altera_acomp_numRows + Repeat1__numRows
%>
<%
Dim rs_situacao_sso
Dim rs_situacao_sso_numRows

Set rs_situacao_sso = Server.CreateObject("ADODB.Recordset")
rs_situacao_sso.ActiveConnection = MM_cpf_STRING
rs_situacao_sso.Source = "SELECT *  FROM tb_situacao_sso"
rs_situacao_sso.CursorType = 0
rs_situacao_sso.CursorLocation = 2
rs_situacao_sso.LockType = 1
rs_situacao_sso.Open()

rs_situacao_sso_numRows = 0
%>
<%
Dim rs_situacao_clima
Dim rs_situacao_clima_numRows

Set rs_situacao_clima = Server.CreateObject("ADODB.Recordset")
rs_situacao_clima.ActiveConnection = MM_cpf_STRING
rs_situacao_clima.Source = "SELECT *  FROM tb_situacao_clima"
rs_situacao_clima.CursorType = 0
rs_situacao_clima.CursorLocation = 2
rs_situacao_clima.LockType = 1
rs_situacao_clima.Open()

rs_situacao_clima_numRows = 0
%>
<%
Dim rs_situacao_obra
Dim rs_situacao_obra_numRows

Set rs_situacao_obra = Server.CreateObject("ADODB.Recordset")
rs_situacao_obra.ActiveConnection = MM_cpf_STRING
rs_situacao_obra.Source = "SELECT *  FROM tb_situacao_obra"
rs_situacao_obra.CursorType = 0
rs_situacao_obra.CursorLocation = 2
rs_situacao_obra.LockType = 1
rs_situacao_obra.Open()

rs_situacao_obra_numRows = 0
%>
<%
Dim rs_fiscal
Dim rs_fiscal_numRows

Set rs_fiscal = Server.CreateObject("ADODB.Recordset")
rs_fiscal.ActiveConnection = MM_cpf_STRING
rs_fiscal.Source = "SELECT *  FROM tb_responsavel  ORDER BY Respons�vel ASC"
rs_fiscal.CursorType = 0
rs_fiscal.CursorLocation = 2
rs_fiscal.LockType = 1
rs_fiscal.Open()

rs_fiscal_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style20 {	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
	font-size: 18px;
	color: #333333;
}
.style28 {font-family: Arial, Helvetica, sans-serif; font-size: 10; font-weight: bold; }
.style31 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
.style33 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; }
.style34 {font-family: Arial, Helvetica, sans-serif}
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
        alert("Informe uma das opc�es");
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
	  if (form[campo].df_validar == "cpf_cnpj") msg = "informe corretamente o n�mero do CPF ou CNPJ";
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
var mensagem = "Este campo aceita somente n�meros"
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
var mensagem = "informe corretamente o n�mero do CPF"
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
var mensagem = "informe corretamente o n�mero do CNPJ"
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
alert("Data inv�lida!"); 
}
} 
}
//-->
</script>
<script type="text/javascript">
function mostraEsconde(empresa,LO)
{
    var check = document.getElementById(empresa);
    var LO = document.getElementById(LO);
    
    if(check.checked==true)
        LO.style.display = 'block';
    else
        LO.style.display = 'none';

}

</script>

<style type="text/css">
  body {
    font-family: Arial, Helvetica, sans-serif !important;
    font-size: 12px !important;
  }

  textarea {
    width: 98%;
  }
</style>
<link rel="stylesheet" href="//code.jquery.com/ui/1.11.3/themes/smoothness/jquery-ui.css">
<script src="//code.jquery.com/jquery-1.11.2.min.js"></script>
<script type="text/javascript" src="//code.jquery.com/ui/1.11.4/jquery-ui.min.js"></script>
<script type="text/javascript" src="js/datepicker-pt-BR.js"></script>
<script type="text/javascript">
  $(function() {
    $(".datepicker").datepicker($.datepicker.regional["pt-BR"]);
  });
</script>
</head>

<body>

<center><h3>Altera��o de RDO</h3></center>

<form method="POST" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC"><strong class="style5">Num. Autos:</strong></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
        <input name="PI" type="text" value="<%=(rs_altera_acomp.Fields.Item("PI").Value)%>" size="18" readonly="true">
      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC"><strong class="style5">Data do Registro:</strong></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
        <input name="Data_do_Registro" type="text" class="datepicker" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" value="<%=(rs_altera_acomp.Fields.Item("Data do Registro").Value)%>" size="15" /></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC"><strong class="style5">Respons�vel:</strong></td>
      <td b<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
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
  MM_editTable = "tb_Acompanhamento"
  MM_editColumn = "cod_acompanhamento"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "altera_acomp.asp"
  MM_fieldsStr  = "Data_do_Registro|value|Registro|value|Responsvel|value|vistoria|value|dt_vistoria|value|flg_pendencia|value|cod_tipo_pendencia|value|dsc_pendencia|value|dta_limite_pendencia|value|cod_situacao_sso|value|cod_situacao_clima_manha|value|cod_situacao_clima_tarde|value|cod_situacao_clima_noite|value|cod_situacao_limpeza_obra|value|cod_situacao_organizacao_obra|value|flg_dia_perdido|value|flg_dia_trabalhado|value|dsc_nota_clima_manha|value|dsc_nota_clima_tarde|value|dsc_nota_clima_noite|value|dsc_nota_limpeza_obra|value|dsc_nota_organizacao_obra|value|dsc_nota_dia_perdido|value|dsc_nota_dia_trabalhado|value"
  MM_columnsStr = "[Data do Registro]|',none,NULL|Registro|',none,''|cod_fiscal|none,none,NULL|vistoria|none,1,0|dt_vistoria|',none,NULL|flg_pendencia|none,1,0|cod_tipo_pendencia|none,none,NULL|dsc_pendencia|',none,''|dta_limite_pendencia|',none,NULL|cod_situacao_sso|none,none,NULL|cod_situacao_clima_manha|none,none,NULL|cod_situacao_clima_tarde|none,none,NULL|cod_situacao_clima_noite|none,none,NULL|cod_situacao_limpeza_obra|none,none,NULL|cod_situacao_organizacao_obra|none,none,NULL|flg_dia_perdido|none,1,0|flg_dia_trabalhado|none,1,0|dsc_nota_clima_manha|',none,''|dsc_nota_clima_tarde|',none,''|dsc_nota_clima_noite|',none,''|dsc_nota_limpeza_obra|',none,''|dsc_nota_organizacao_obra|',none,''|dsc_nota_dia_perdido|',none,''|dsc_nota_dia_trabalhado|',none,''"

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
Dim rs_altera_acomp__MMColParam
rs_altera_acomp__MMColParam = "1"
If (Request.QueryString("cod_acompanhamento") <> "") Then 
  rs_altera_acomp__MMColParam = Request.QueryString("cod_acompanhamento")
End If
%>
<%
Dim rs_altera_acomp
Dim rs_altera_acomp_numRows

Set rs_altera_acomp = Server.CreateObject("ADODB.Recordset")
rs_altera_acomp.ActiveConnection = MM_cpf_STRING
rs_altera_acomp.Source = "SELECT *  FROM c_Acomp_resp  WHERE cod_acompanhamento = " + Replace(rs_altera_acomp__MMColParam, "'", "''") + ""
rs_altera_acomp.CursorType = 0
rs_altera_acomp.CursorLocation = 2
rs_altera_acomp.LockType = 1
rs_altera_acomp.Open()

rs_altera_acomp_numRows = 0
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_cpf_STRING
Recordset1.Source = "SELECT * FROM tb_responsavel ORDER BY Respons�vel ASC"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rs_altera_acomp_numRows = rs_altera_acomp_numRows + Repeat1__numRows
%>
<%
Dim rs_situacao_sso
Dim rs_situacao_sso_numRows

Set rs_situacao_sso = Server.CreateObject("ADODB.Recordset")
rs_situacao_sso.ActiveConnection = MM_cpf_STRING
rs_situacao_sso.Source = "SELECT *  FROM tb_situacao_sso"
rs_situacao_sso.CursorType = 0
rs_situacao_sso.CursorLocation = 2
rs_situacao_sso.LockType = 1
rs_situacao_sso.Open()

rs_situacao_sso_numRows = 0
%>
<%
Dim rs_situacao_clima
Dim rs_situacao_clima_numRows

Set rs_situacao_clima = Server.CreateObject("ADODB.Recordset")
rs_situacao_clima.ActiveConnection = MM_cpf_STRING
rs_situacao_clima.Source = "SELECT *  FROM tb_situacao_clima"
rs_situacao_clima.CursorType = 0
rs_situacao_clima.CursorLocation = 2
rs_situacao_clima.LockType = 1
rs_situacao_clima.Open()

rs_situacao_clima_numRows = 0
%>
<%
Dim rs_situacao_obra
Dim rs_situacao_obra_numRows

Set rs_situacao_obra = Server.CreateObject("ADODB.Recordset")
rs_situacao_obra.ActiveConnection = MM_cpf_STRING
rs_situacao_obra.Source = "SELECT *  FROM tb_situacao_obra"
rs_situacao_obra.CursorType = 0
rs_situacao_obra.CursorLocation = 2
rs_situacao_obra.LockType = 1
rs_situacao_obra.Open()

rs_situacao_obra_numRows = 0
%>
<%
Dim rs_fiscal
Dim rs_fiscal_numRows

Set rs_fiscal = Server.CreateObject("ADODB.Recordset")
rs_fiscal.ActiveConnection = MM_cpf_STRING
rs_fiscal.Source = "SELECT *  FROM tb_responsavel  ORDER BY Respons�vel ASC"
rs_fiscal.CursorType = 0
rs_fiscal.CursorLocation = 2
rs_fiscal.LockType = 1
rs_fiscal.Open()

rs_fiscal_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style20 {	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
	font-size: 18px;
	color: #333333;
}
.style28 {font-family: Arial, Helvetica, sans-serif; font-size: 10; font-weight: bold; }
.style31 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
.style33 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; }
.style34 {font-family: Arial, Helvetica, sans-serif}
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
        alert("Informe uma das opc�es");
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
	  if (form[campo].df_validar == "cpf_cnpj") msg = "informe corretamente o n�mero do CPF ou CNPJ";
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
var mensagem = "Este campo aceita somente n�meros"
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
var mensagem = "informe corretamente o n�mero do CPF"
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
var mensagem = "informe corretamente o n�mero do CNPJ"
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
alert("Data inv�lida!"); 
}
} 
}
//-->
</script>
<script type="text/javascript">
function mostraEsconde(empresa,LO)
{
    var check = document.getElementById(empresa);
    var LO = document.getElementById(LO);
    
    if(check.checked==true)
        LO.style.display = 'block';
    else
        LO.style.display = 'none';

}

</script>

<style type="text/css">
  body {
    font-family: Arial, Helvetica, sans-serif !important;
    font-size: 12px !important;
  }

  textarea {
    width: 98%;
  }
</style>
<link rel="stylesheet" href="//code.jquery.com/ui/1.11.3/themes/smoothness/jquery-ui.css">
<script src="//code.jquery.com/jquery-1.11.2.min.js"></script>
<script type="text/javascript" src="//code.jquery.com/ui/1.11.4/jquery-ui.min.js"></script>
<script type="text/javascript" src="js/datepicker-pt-BR.js"></script>
<script type="text/javascript">
  $(function() {
    $(".datepicker").datepicker($.datepicker.regional["pt-BR"]);
  });
</script>
</head>

<body>

<center><h3>Altera��o de RDO</h3></center>

<form method="POST" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC"><strong class="style5">Num. Autos:</strong></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
        <input name="PI" type="text" value="<%=(rs_altera_acomp.Fields.Item("PI").Value)%>" size="18" readonly="true">
      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC"><strong class="style5">Data do Registro:</strong></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
        <input name="Data_do_Registro" type="text" class="datepicker" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" value="<%=(rs_altera_acomp.Fields.Item("Data do Registro").Value)%>" size="15" /></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC"><strong class="style5">Respons�vel:</strong></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
        <select name="Responsvel">
          <option value=""></option>
          <%
            While (NOT rs_fiscal.EOF)
          %>
          <option value="<%=(rs_fiscal.Fields.Item("cod_fiscal").Value)%>"><%=(rs_fiscal.Fields.Item("Respons�vel").Value)%></option>
          <%
              If Trim(rs_fiscal.Fields.Item("Respons�vel").Value) <> "" Then
                Response.Write "      <OPTION value='" & (rs_fiscal.Fields.Item("cod_fiscal").Value) & "'"
                If Lcase(rs_fiscal.Fields.Item("cod_fiscal").Value) = Lcase(rs_altera_acomp.Fields.Item("cod_fiscal").Value) then
                  Response.Write "selected"
                End If
                Response.Write ">" & (rs_fiscal.Fields.Item("Respons�vel").Value) & "</OPTION>"
              End If
              rs_fiscal.MoveNext()
              Wend
              If (rs_fiscal.CursorType > 0) Then
                rs_fiscal.MoveFirst
              Else
                rs_fiscal.Requery
              End If
          %>
      </select>
      </td>
    </tr>
    <tr valign="middle">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC"><strong class="style5">Registro:</strong></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC"><textarea name="Registro" cols="32" rows="5" maxlength="255" style="width: 97%;"><%=(rs_altera_acomp.Fields.Item("Registro").Value)%></textarea></td>
    </tr>
    <tr valign="middle">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC"><strong class="style5">MSST:</strong></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
        Tipo Situa��o: <br/>
        <select name="cod_situacao_sso" style="width: 100%;">
          <option value=""></option>
          <%
            While (NOT rs_situacao_sso.EOF)
              If Trim(rs_situacao_sso.Fields.Item("dsc_situacao_sso").Value) <> "" Then
                Response.Write "      <OPTION value='" & (rs_situacao_sso.Fields.Item("id").Value) & "'"
                If Lcase(rs_situacao_sso.Fields.Item("id").Value) = Lcase(rs_altera_acomp.Fields.Item("cod_situacao_sso").Value) then
                  Response.Write "selected"
                End If
                Response.Write ">" & (rs_situacao_sso.Fields.Item("dsc_situacao_sso").Value) & "</OPTION>"
              End If
  