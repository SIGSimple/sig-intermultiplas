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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_cpf_STRING
  MM_editTable = "tb_Acompanhamento"
  MM_editRedirectUrl = "acompanhamento_inclui.asp"
  MM_fieldsStr  = "PI|value|Data_do_Registro|value|Responsvel|value|Previso|value|Registro|value|vistoria|value|n_LO|value|dt_vistoria|value"
  MM_columnsStr = "PI|',none,''|[Data do Registro]|',none,NULL|cod_fiscal|none,none,NULL|Previsão|',none,''|Registro|',none,''|vistoria|none,1,0|n_LO|',none,''|dt_vistoria|',none,NULL"

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
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form2" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_cpf_STRING
  MM_editTable = "tb_pi"
  MM_editColumn = "PI"
  MM_recordId = "'" + Request.Form("MM_recordId") + "'"
  MM_editRedirectUrl = "acompanhamento_inclui.asp"
  MM_fieldsStr  = "Data_da_abertura|value|Data_do_TRP|value|Data_do_TRD|value|cod_situacao|value|Foi_solicitado_Aditamento|value|Valor_do_Aditamento|value|Prazo_do_Aditamento|value"
  MM_columnsStr = "[Data da Abertura]|',none,NULL|[Data do TRP]|',none,NULL|[Data do TRD]|',none,NULL|cod_situacao|none,none,NULL|[Foi solicitado Aditamento?]|none,1,0|[Valor do Aditamento]|',none,''|[Prazo do Aditamento]|none,none,NULL"

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
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
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
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
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
Dim rs_pi__MMColParam
rs_pi__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  rs_pi__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim rs_pi
Dim rs_pi_numRows

Set rs_pi = Server.CreateObject("ADODB.Recordset")
rs_pi.ActiveConnection = MM_cpf_STRING
rs_pi.Source = "SELECT tb_pi.PI, [tb_predio].[cod_predio] & ' - ' & [tb_predio].[Município] AS Expr1, c_Semaforico1.[Avanço Físico Atual],[tb_predio].[cod_predio]  FROM c_Semaforico1 INNER JOIN (tb_predio INNER JOIN tb_pi ON tb_predio.cod_predio = tb_pi.cod_predio) ON c_Semaforico1.[PI-item] = tb_pi.PI  WHERE PI = '" + Replace(rs_pi__MMColParam, "'", "''") + "'  ORDER BY tb_pi.PI;"
rs_pi.CursorType = 0
rs_pi.CursorLocation = 2
rs_pi.LockType = 1
rs_pi.Open()

rs_pi_numRows = 0
%>
<%
Dim rs_fiscal
Dim rs_fiscal_numRows

Set rs_fiscal = Server.CreateObject("ADODB.Recordset")
rs_fiscal.ActiveConnection = MM_cpf_STRING
rs_fiscal.Source = "SELECT *  FROM tb_responsavel  ORDER BY Responsável ASC"
rs_fiscal.CursorType = 0
rs_fiscal.CursorLocation = 2
rs_fiscal.LockType = 1
rs_fiscal.Open()

rs_fiscal_numRows = 0
%>
<%
Dim rs_acomp__MMColParam
rs_acomp__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  rs_acomp__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim rs_acomp
Dim rs_acomp_numRows

Set rs_acomp = Server.CreateObject("ADODB.Recordset")
rs_acomp.ActiveConnection = MM_cpf_STRING
rs_acomp.Source = "SELECT *  FROM tb_Acompanhamento  WHERE PI = '" + Replace(rs_acomp__MMColParam, "'", "''") + "'  "
rs_acomp.CursorType = 0
rs_acomp.CursorLocation = 2
rs_acomp.LockType = 1
rs_acomp.Open()

rs_acomp_numRows = 0
%>
<%
Dim rs_lista_acomp__MMColParam
rs_lista_acomp__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  rs_lista_acomp__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim rs_lista_acomp
Dim rs_lista_acomp_numRows

Set rs_lista_acomp = Server.CreateObject("ADODB.Recordset")
rs_lista_acomp.ActiveConnection = MM_cpf_STRING
rs_lista_acomp.Source = "SELECT *  FROM cLista_acomp  WHERE PI = '" + Replace(rs_lista_acomp__MMColParam, "'", "''") + "'  "
rs_lista_acomp.CursorType = 0
rs_lista_acomp.CursorLocation = 2
rs_lista_acomp.LockType = 1
rs_lista_acomp.Open()

rs_lista_acomp_numRows = 10
%>
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_cpf_STRING
Recordset1.Source = "SELECT tb_pi.PI, tb_pi.[Descrição da Intervenção FDE]  FROM tb_pi  WHERE PI = '" + Replace(Recordset1__MMColParam, "'", "''") + "'  GROUP BY tb_pi.PI, tb_pi.[Descrição da Intervenção FDE];  "
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
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
Recordset2.Source = "SELECT * FROM tb_pi WHERE PI = '" + Replace(Recordset2__MMColParam, "'", "''") + "'"
Recordset2.CursorType = 0
Recordset2.CursorLocation = 2
Recordset2.LockType = 1
Recordset2.Open()

Recordset2_numRows = 0
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
Dim Recordset3__MMColParam
Recordset3__MMColParam = "1"
If (Request.QueryString("PI-item") <> "") Then 
  Recordset3__MMColParam = Request.QueryString("PI-item")
End If
%>
<%
Dim Recordset3
Dim Recordset3_cmd
Dim Recordset3_numRows

Set Recordset3_cmd = Server.CreateObject ("ADODB.Command")
Recordset3_cmd.ActiveConnection = MM_cpf_STRING
Recordset3_cmd.CommandText = "SELECT * FROM c_Semaforico1 WHERE [PI-item] = ?" 
Recordset3_cmd.Prepared = true
Recordset3_cmd.Parameters.Append Recordset3_cmd.CreateParameter("param1", 200, 1, 50, Recordset3__MMColParam) ' adVarChar

Set Recordset3 = Recordset3_cmd.Execute
Recordset3_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rs_acomp_numRows = rs_acomp_numRows + Repeat1__numRows
%>
<%

Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = 1000
Repeat2__index = 0
rs_lista_acomp_numRows = rs_lista_acomp_numRows + Repeat2__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style5 {font-family: Arial, Helvetica, sans-serif; font-size: 12; }
.style6 {font-size: 12}
.style19 {font-family: Arial, Helvetica, sans-serif; font-size: 14px; font-weight: bold; }
.style22 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
.style26 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
.style27 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
}
.style1 {	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
	font-size: 10;
}
.style28 {font-family: Arial, Helvetica, sans-serif; font-size: 10; }
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

</style>
</head>

<body>
<div align="center">
  &nbsp;<table width="823" border="0">
    <tr bgcolor="#CCCCCC">
      <td><span class="style19"><%=(rs_pi.Fields.Item("PI").Value)%></span></td>
      <td colspan="2"><div align="left"><span class="style19"><a href="filtro_exibir_acomp_adm_url.asp?cod_predio=<%=(rs_pi.Fields.Item("cod_predio").Value)%>" target="_blank"><%=(rs_pi.Fields.Item("Expr1").Value)%></a></span></div></td>
      <td><span class="style27">Avan&ccedil;o F&iacute;sico Atual</span>: <%= FormatPercent((rs_pi.Fields.Item("Avanço Físico Atual").Value), 2, -2, -2, -2) %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="147" bgcolor="#CCCCCC" class="style27">Interven&ccedil;&atilde;o FDE </td>
      <td width="326" bgcolor="#CCCCCC"><div align="left" class="style27"><%=(Recordset1.Fields.Item("Descrição da Intervenção FDE").Value)%></div></td>
      <td width="110" bgcolor="#CCCCCC" class="style27">T&eacute;rmino contratual</td>
      <td width="222" bgcolor="#CCCCCC"><span class="style26"><%=(rs_lista_acomp.Fields.Item("termino_contratual").Value)%></span></td>
    </tr>
  </table>
  <table width="823" border="0">
    <tr bgcolor="#FFFFFF">
      <td bgcolor="#FFFFFF" class="style27"><a href="med_constr_inclui_fiscal.asp?pi=<%=(rs_pi.Fields.Item("PI").Value)%>" target="_blank">Medi&ccedil;&otilde;es</a></td>
    </tr>
  </table>
  <form id="form2" name="form2" method="POST" action="<%=MM_editAction%>">
    <table align="center">
      <tr valign="baseline" bgcolor="#999999">
        <td align="right" nowrap="nowrap" bgcolor="#EBEBEB"><span class="style1">Data da Abertura</span></td>
        <td bgcolor="#EBEBEB"><div align="left"><span class="style1">
            <input name="Data_da_abertura" type="text" class="style28" id="Data_da_abertura" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" value="<%=(Recordset2.Fields.Item("Data da Abertura").Value)%>" size="15" />
        </span></div></td>
        <td bgcolor="#EBEBEB"><div align="left"><span class="style1">Data do TRP</span></div></td>
        <td bgcolor="#EBEBEB"><input name="Data_do_TRP" type="text" class="style28" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" value="<%=(Recordset2.Fields.Item("Data do TRP").Value)%>" size="15" /></td>
        <td bgcolor="#EBEBEB"><div align="right"><strong><span class="style28">Data do TRD</span></strong></div></td>
        <td bgcolor="#EBEBEB"><input name="Data_do_TRD" type="text" class="style28" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" value="<%=(Recordset2.Fields.Item("Data do TRD").Value)%>" size="15" /></td>
      </tr>
    </table>
      <table align="center">
        <tr valign="baseline" bgcolor="#999999">
          <td align="right" nowrap="nowrap" bgcolor="#EBEBEB"><span class="style1">Est&aacute;gio da Obra </span></td>
          <td colspan="5" bgcolor="#EBEBEB"><label for="select"></label>
              <select name="cod_situacao" id="select">
                <option value="" <%If (Not isNull((Recordset2.Fields.Item("cod_situacao").Value))) Then If ("" = CStr((Recordset2.Fields.Item("cod_situacao").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>></option>
                <%
While (NOT rs_situacao.EOF)
%><option value="<%=(rs_situacao.Fields.Item("cod_situacao").Value)%>" <%If (Not isNull((Recordset2.Fields.Item("cod_situacao").Value))) Then If (CStr(rs_situacao.Fields.Item("cod_situacao").Value) = CStr((Recordset2.Fields.Item("cod_situacao").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(rs_situacao.Fields.Item("desc_situacao").Value)%></option>
                <%
  rs_situacao.MoveNext()
Wend
If (rs_situacao.CursorType > 0) Then
  rs_situacao.MoveFirst
Else
  rs_situacao.Requery
End If
%>
            </select>
              <div align="right"></div></td>
        </tr>
      </table>
      <p>&nbsp;</p>
      <table align="center">
        <tr valign="baseline">
          <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><span class="style5">Foi solicitado Aditamento?:</span></td>
          <td bgcolor="#CCCCCC"><div align="left">
            <input <%If (CStr((Recordset2.Fields.Item("Foi solicitado Aditamento?").Value)) = CStr("True")) Then Response.Write("checked=""checked""") : Response.Write("")%> type="checkbox" name="Foi_solicitado_Aditamento" value="1">
          </div></td>
        </tr>
        <tr valign="baseline">
          <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><span class="style5">Valor do Aditamento:</span></td>
          <td bgcolor="#CCCCCC"><div align="left">
            <input type="text" name="Valor_do_Aditamento" value="<%=(Recordset2.Fields.Item("Valor do Aditamento").Value)%>" size="20" />
          </div></td>
        </tr>
        <tr valign="baseline">
          <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><span class="style5">Prazo do Aditamento:</span></td>
          <td bgcolor="#CCCCCC"><div align="left">
            <input type="text" name="Prazo_do_Aditamento" value="<%=(Recordset2.Fields.Item("Prazo do Aditamento").Value)%>" size="15" />
          </div></td>
        </tr>
      </table>
      <label for="Submit"></label>
    <input type="submit" name="Submit" value="Salvar" id="Submit" />
    <input type="hidden" name="MM_update" value="form2">
    <input type="hidden" name="MM_recordId" value="<%= Recordset2.Fields.Item("PI").Value %>">
  </form>
</div>
<form method="POST" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC"><span class="style5">PI:</span></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC"><select name="PI">
        <%
While (NOT rs_pi.EOF)
%>
        <option value="<%=(rs_pi.Fields.Item("PI").Value)%>"><%=(rs_pi.Fields.Item("PI").Value)%></option>
        <%
  rs_pi.MoveNext()
Wend
If (rs_pi.CursorType > 0) Then
  rs_pi.MoveFirst
Else
  rs_pi.Requery
End If
%>
      </select>      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC"><span class="style5">Data do Registro:</span></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC"><input name="Data_do_Registro" type="text" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" value="<% =Date %>" size="15">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC"><span class="style5">Responsável:</span></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC"><select name="Responsvel">
        <option value=""></option>
        <%
While (NOT rs_fiscal.EOF)
%><option value="<%=(rs_fiscal.Fields.Item("cod_fiscal").Value)%>"><%=(rs_fiscal.Fields.Item("Responsável").Value)%></option>
          <%
  rs_fiscal.MoveNext()
Wend
If (rs_fiscal.CursorType > 0) Then
  rs_fiscal.MoveFirst
Else
  rs_fiscal.Requery
End If
%>
      </select>      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC"><span class="style5">Previsão:</span></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC"><input name="Previso" type="text" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" value="" size="15">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC"><span class="style5">Registro:</span></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC"><textarea name="Registro" cols="32"></textarea>      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC"><span class="style5">Vistoria:</span></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC"><input <%If (CStr((rs_lista_acomp.Fields.Item("vist").Value)) = CStr("True")) Then Response.Write("checked=""checked""") : Response.Write("")%> type="checkbox" name="vistoria" value="1" id="vistoria" />
      <label for="checkbox"></label></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap="nowrap" bordercolor="#CCCCCC" bgcolor="#CCCCCC"><span class="style29">N&uacute;mero do   LO</span> </td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC"><label>
        <input name="n_LO" type="text" id="n_LO" size="15" />
        </label>
        ex: XX00000</td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap="nowrap" bordercolor="#CCCCCC" bgcolor="#CCCCCC"><span class="style29">Data da Vistoria </span></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC"><input name="dt_vistoria" type="text" id="dt_vistoria" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" size="15" /></td>
    </tr>
    
    <tr valign="baseline">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC"><span class="style6"></span></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC"><input type="submit" value="Salvar">      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1">
</form>
<div align="center">
  <table border="0">
    <tr bgcolor="#999999">
      <td><span class="style26">Data do Registro</span></td>
      <td><span class="style26">Registro</span></td>
      <td><span class="style26">Previs&atilde;o</span></td>
      <td><span class="style26">Respons&aacute;vel</span></td>
      <td><span class="style26">Vistoria</span></td>
      <td width="76"><span class="style26">N&ordm; LO </span></td>
      <td width="94"><span class="style26">Data Vistoria </span></td>
    </tr>
    <% While ((Repeat2__numRows <> 0) AND (NOT rs_lista_acomp.EOF)) %>
      <tr bgcolor="#CCCCCC">
        <td><span class="style22"><%=(rs_lista_acomp.Fields.Item("Data do Registro").Value)%></span></td>
        <td><div align="left"><span class="style22"><%=(rs_lista_acomp.Fields.Item("Registro").Value)%></span></div></td>
        <td><span class="style22"><%=(rs_lista_acomp.Fields.Item("Previsão").Value)%></span></td>
        <td><span class="style22"><%=(rs_lista_acomp.Fields.Item("Responsável").Value)%></span></td>
        <td class="style22"><%=(rs_lista_acomp.Fields.Item("vist").Value)%></td>
        <td class="style22"><%=(rs_lista_acomp.Fields.Item("n_lo").Value)%></td>
        <td class="style22"><%=(rs_lista_acomp.Fields.Item("dt_vistoria").Value)%></td>
      </tr>
      <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rs_lista_acomp.MoveNext()
Wend
%>
  </table>
</div>
</body>
</html>
<%
rs_pi.Close()
Set rs_pi = Nothing
%>
<%
rs_fiscal.Close()
Set rs_fiscal = Nothing
%>
<%
rs_acomp.Close()
Set rs_acomp = Nothing
%>
<%
rs_lista_acomp.Close()
Set rs_lista_acomp = Nothing
%>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>
<%
rs_situacao.Close()
Set rs_situacao = Nothing
%>
<%
Recordset3.Close()
Set Recordset3 = Nothing
%>
