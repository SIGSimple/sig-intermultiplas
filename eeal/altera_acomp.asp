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
  MM_fieldsStr  = "PI|value|Data_do_Registro|value|Previso|value|cod_fiscal|value|Registro|value|n_LO|value|dt_vistoria|value"
  MM_columnsStr = "PI|',none,''|[Data do Registro]|',none,NULL|Previsão|',none,''|cod_fiscal|none,none,NULL|Registro|',none,''|n_LO|',none,''|dt_vistoria|',none,NULL"

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
Recordset1.Source = "SELECT * FROM tb_responsavel ORDER BY Responsável ASC"
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

<body>
<p align="center"><span class="style20">ALTERA&Ccedil;&Atilde;O DO ACOMPANHAMENTO</span></p>
<form method="POST" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style28">PI:</span></td>
      <td bgcolor="#CCCCCC"><input name="PI" type="text" value="<%=(rs_altera_acomp.Fields.Item("PI").Value)%>" size="18" readonly="true">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style28">Data do Registro:</span></td>
      <td bgcolor="#CCCCCC"><input name="Data_do_Registro" type="text" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" value="<%=(rs_altera_acomp.Fields.Item("Data do Registro").Value)%>" size="15">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style28">Previsão:</span></td>
      <td bgcolor="#CCCCCC"><input name="Previso" type="text" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" value="<%=(rs_altera_acomp.Fields.Item("Previsão").Value)%>" size="15">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style28">Respons&aacute;vel:</span></td>
      <td bgcolor="#CCCCCC"><label for="select"></label>
        <select name="cod_fiscal" id="cod_fiscal">
          <option value="" <%If (Not isNull((rs_altera_acomp.Fields.Item("cod_fiscal").Value))) Then If ("" = CStr((rs_altera_acomp.Fields.Item("cod_fiscal").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>></option>
          <%
While (NOT Recordset1.EOF)
%><option value="<%=(Recordset1.Fields.Item("cod_fiscal").Value)%>" <%If (Not isNull((rs_altera_acomp.Fields.Item("cod_fiscal").Value))) Then If (CStr(Recordset1.Fields.Item("cod_fiscal").Value) = CStr((rs_altera_acomp.Fields.Item("cod_fiscal").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(Recordset1.Fields.Item("Responsável").Value)%></option>
          <%
  Recordset1.MoveNext()
Wend
If (Recordset1.CursorType > 0) Then
  Recordset1.MoveFirst
Else
  Recordset1.Requery
End If
%>
        </select></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style28">Registro:</span></td>
      <td bgcolor="#CCCCCC"><textarea name="Registro" cols="40"><%=(rs_altera_acomp.Fields.Item("Registro").Value)%></textarea>      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap="nowrap" bordercolor="#CCCCCC" bgcolor="#CCCCCC"><span class="style29 style34"><strong>N&uacute;mero do   LO</strong></span> </td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC"><label>
        <input name="n_LO" type="text" id="n_LO" value="<%=(rs_altera_acomp.Fields.Item("n_LO").Value)%>" size="15" />
        </label>
        ex: XX00000</td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap="nowrap" bordercolor="#CCCCCC" bgcolor="#CCCCCC"><span class="style29 style34"><strong>Data da Vistoria </strong></span></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC"><input name="dt_vistoria" type="text" id="dt_vistoria" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" value="<%=(rs_altera_acomp.Fields.Item("dt_vistoria").Value)%>" size="15" /></td>
    </tr>
    
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC">&nbsp;</td>
      <td bgcolor="#CCCCCC"><input type="submit" value="Salvar">      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= rs_altera_acomp.Fields.Item("cod_acompanhamento").Value %>">
</form>
<p>&nbsp;</p>

<div align="center">
  <table border="0">
    <tr bgcolor="#CCCCCC">
      <td bgcolor="#999999"><span class="style33">PI</span></td>
      <td bgcolor="#999999"><span class="style33">Data do Registro</span></td>
      <td bgcolor="#999999"><span class="style33">Registro</span></td>
      <td bgcolor="#999999"><span class="style33">Respons&aacute;vel</span></td>
      <td bgcolor="#999999"><span class="style33">Previs&atilde;o</span></td>
      <td width="76" bgcolor="#999999" class="style33"><span class="style26">N&ordm; LO </span></td>
      <td width="94" bgcolor="#999999" class="style33"><span class="style26">Data Vistoria </span></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rs_altera_acomp.EOF)) %>
      <tr bgcolor="#CCCCCC">
        <td><span class="style31"><%=(rs_altera_acomp.Fields.Item("PI").Value)%></span></td>
        <td><span class="style31"><%=(rs_altera_acomp.Fields.Item("Data do Registro").Value)%></span></td>
        <td><span class="style31"><%=(rs_altera_acomp.Fields.Item("Registro").Value)%></span></td>
        <td><span class="style31"><%=(rs_altera_acomp.Fields.Item("Responsável").Value)%></span></td>
        <td><span class="style31"><%=(rs_altera_acomp.Fields.Item("Previsão").Value)%></span></td>
        <td class="style33"><%=(rs_altera_acomp.Fields.Item("n_LO").Value)%></td>
        <td class="style33"><%=(rs_altera_acomp.Fields.Item("dt_vistoria").Value)%></td>
      </tr>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_altera_acomp.MoveNext()
Wend
%>
  </table>
</div>
</body>
</html>
<%
rs_altera_acomp.Close()
Set rs_altera_acomp = Nothing
%>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
