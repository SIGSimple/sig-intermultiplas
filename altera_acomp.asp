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
  MM_fieldsStr  = "Data_do_Registro|value|Registro|value|Responsvel|value|vistoria|value|dt_vistoria|value|flg_pendencia|value|cod_tipo_pendencia|value|dsc_pendencia|value|dta_limite_pendencia|value|cod_situacao_sso|value|dsc_nota_sso|value|cod_situacao_clima_manha|value|cod_situacao_clima_tarde|value|cod_situacao_clima_noite|value|cod_situacao_limpeza_obra|value|cod_situacao_organizacao_obra|value|dsc_nota_clima_manha|value|dsc_nota_clima_tarde|value|dsc_nota_clima_noite|value|dsc_nota_limpeza_obra|value|dsc_nota_organizacao_obra|value|dsc_nota_dia_perdido|value|dsc_nota_dia_trabalhado|value"
  MM_columnsStr = "[Data do Registro]|',none,NULL|Registro|',none,''|cod_fiscal|none,none,NULL|vistoria|none,1,0|dt_vistoria|',none,NULL|flg_pendencia|none,1,0|cod_tipo_pendencia|none,none,NULL|dsc_pendencia|',none,''|dta_limite_pendencia|',none,NULL|cod_situacao_sso|none,none,NULL|dsc_nota_sso|',none,''|cod_situacao_clima_manha|none,none,NULL|cod_situacao_clima_tarde|none,none,NULL|cod_situacao_clima_noite|none,none,NULL|cod_situacao_limpeza_obra|none,none,NULL|cod_situacao_organizacao_obra|none,none,NULL|dsc_nota_clima_manha|',none,''|dsc_nota_clima_tarde|',none,''|dsc_nota_clima_noite|',none,''|dsc_nota_limpeza_obra|',none,''|dsc_nota_organizacao_obra|',none,''|dsc_nota_dia_perdido|',none,''|dsc_nota_dia_trabalhado|',none,''"

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
rs_fiscal.Source = "SELECT *  FROM tb_responsavel  ORDER BY Responsável ASC"
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

<center><h3>Alteração de RDO</h3></center>

<form method="POST" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC"><strong class="style5">Data do Registro:</strong></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
        <input name="Data_do_Registro" type="text" class="datepicker" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" value="<%=(rs_altera_acomp.Fields.Item("Data do Registro").Value)%>" size="15" /></td>
    </tr>
    
    <tr valign="baseline">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC"><strong class="style5">Responsável:</strong></td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
        <select name="Responsvel">
          <option value=""></option>
          <%
            While (NOT rs_fiscal.EOF)
          %>
          <option value="<%=(rs_fiscal.Fields.Item("cod_fiscal").Value)%>"><%=(rs_fiscal.Fields.Item("Responsável").Value)%></option>
          <%
              If Trim(rs_fiscal.Fields.Item("Responsável").Value) <> "" Then
                Response.Write "      <OPTION value='" & (rs_fiscal.Fields.Item("cod_fiscal").Value) & "'"
                If Lcase(rs_fiscal.Fields.Item("cod_fiscal").Value) = Lcase(rs_altera_acomp.Fields.Item("cod_fiscal").Value) then
                  Response.Write "selected"
                End If
                Response.Write ">" & (rs_fiscal.Fields.Item("Responsável").Value) & "</OPTION>"
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
        Tipo Situação: <br/>
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
              rs_situacao_sso.MoveNext()
            Wend
            If (rs_situacao_sso.CursorType > 0) Then
              rs_situacao_sso.MoveFirst
            Else
              rs_situacao_sso.Requery
            End If
          %>
        </select>
        <br/>
        Nota/Observações: <br/>
        <textarea name="dsc_nota_sso" cols="32"><%=(rs_altera_acomp.Fields.Item("dsc_nota_sso").Value)%></textarea>
      </td>
    </tr>

    <tr valign="middle">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC">
        <strong class="style5">Condições da Obra:</strong>
      </td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
        <table>
          <tr>
            <td>
              Clima Manhã:
            </td>
            <td>
              <select name="cod_situacao_clima_manha" style="width: 100%;">
                <option value=""></option>
                <%
                  While (NOT rs_situacao_clima.EOF)
                    If Trim(rs_situacao_clima.Fields.Item("dsc_situacao_clima").Value) <> "" Then
                      Response.Write "      <OPTION value='" & (rs_situacao_clima.Fields.Item("id").Value) & "'"
                      If Lcase(rs_situacao_clima.Fields.Item("id").Value) = Lcase(rs_altera_acomp.Fields.Item("cod_situacao_clima_manha").Value) then
                        Response.Write "selected"
                      End If
                      Response.Write ">" & (rs_situacao_clima.Fields.Item("dsc_situacao_clima").Value) & "</OPTION>"
                    End If
                    rs_situacao_clima.MoveNext()
                  Wend
                  If (rs_situacao_clima.CursorType > 0) Then
                    rs_situacao_clima.MoveFirst
                  Else
                    rs_situacao_clima.Requery
                  End If
                %>
              </select>
            </td>
            <td>
              <input type="text" name="dsc_nota_clima_manha" placeholder="Notas" value="<%=(rs_altera_acomp.Fields.Item("dsc_nota_clima_manha").Value)%>">
            </td>
          </tr>
          <tr>
            <td>
              Clima Tarde:
            </td>
            <td>
              <select name="cod_situacao_clima_tarde" style="width: 100%;">
                <option value=""></option>
                <%
                  While (NOT rs_situacao_clima.EOF)
                    If Trim(rs_situacao_clima.Fields.Item("dsc_situacao_clima").Value) <> "" Then
                      Response.Write "      <OPTION value='" & (rs_situacao_clima.Fields.Item("id").Value) & "'"
                      If Lcase(rs_situacao_clima.Fields.Item("id").Value) = Lcase(rs_altera_acomp.Fields.Item("cod_situacao_clima_tarde").Value) then
                        Response.Write "selected"
                      End If
                      Response.Write ">" & (rs_situacao_clima.Fields.Item("dsc_situacao_clima").Value) & "</OPTION>"
                    End If
                    rs_situacao_clima.MoveNext()
                  Wend
                  If (rs_situacao_clima.CursorType > 0) Then
                    rs_situacao_clima.MoveFirst
                  Else
                    rs_situacao_clima.Requery
                  End If
                %>
              </select>
            </td>
            <td>
              <input type="text" name="dsc_nota_clima_tarde" placeholder="Notas" value="<%=(rs_altera_acomp.Fields.Item("dsc_nota_clima_tarde").Value)%>">
            </td>
          </tr>
          <tr>
            <td>
              Clima Noite:
            </td>
            <td>
              <select name="cod_situacao_clima_noite" style="width: 100%;">
                <option value=""></option>
                <%
                  While (NOT rs_situacao_clima.EOF)
                    If Trim(rs_situacao_clima.Fields.Item("dsc_situacao_clima").Value) <> "" Then
                      Response.Write "      <OPTION value='" & (rs_situacao_clima.Fields.Item("id").Value) & "'"
                      If Lcase(rs_situacao_clima.Fields.Item("id").Value) = Lcase(rs_altera_acomp.Fields.Item("cod_situacao_clima_noite").Value) then
                        Response.Write "selected"
                      End If
                      Response.Write ">" & (rs_situacao_clima.Fields.Item("dsc_situacao_clima").Value) & "</OPTION>"
                    End If
                    rs_situacao_clima.MoveNext()
                  Wend
                  If (rs_situacao_clima.CursorType > 0) Then
                    rs_situacao_clima.MoveFirst
                  Else
                    rs_situacao_clima.Requery
                  End If
                %>
              </select>
            </td>
            <td>
              <input type="text" name="dsc_nota_clima_noite" placeholder="Notas" value="<%=(rs_altera_acomp.Fields.Item("dsc_nota_clima_noite").Value)%>">
            </td>
          </tr>
          <tr>
            <td>
              Limpeza da Obra:
            </td>
            <td>
              <select name="cod_situacao_limpeza_obra" style="width: 100%;">
                <option value=""></option>
                <%
                  While (NOT rs_situacao_obra.EOF)
                    If Trim(rs_situacao_obra.Fields.Item("dsc_situacao_obra").Value) <> "" Then
                      Response.Write "      <OPTION value='" & (rs_situacao_obra.Fields.Item("id").Value) & "'"
                      If Lcase(rs_situacao_obra.Fields.Item("id").Value) = Lcase(rs_altera_acomp.Fields.Item("cod_situacao_limpeza_obra").Value) then
                        Response.Write "selected"
                      End If
                      Response.Write ">" & (rs_situacao_obra.Fields.Item("dsc_situacao_obra").Value) & "</OPTION>"
                    End If
                    rs_situacao_obra.MoveNext()
                  Wend
                  If (rs_situacao_obra.CursorType > 0) Then
                    rs_situacao_obra.MoveFirst
                  Else
                    rs_situacao_obra.Requery
                  End If
                %>
              </select>
            </td>
            <td>
              <input type="text" name="dsc_nota_limpeza_obra" placeholder="Notas" value="<%=(rs_altera_acomp.Fields.Item("dsc_nota_limpeza_obra").Value)%>">
            </td>
          </tr>
          <tr>
            <td>
              Organização da Obra:
            </td>
            <td>
              <select name="cod_situacao_organizacao_obra" style="width: 100%;">
                <option value=""></option>
                <%
                  While (NOT rs_situacao_obra.EOF)
                    If Trim(rs_situacao_obra.Fields.Item("dsc_situacao_obra").Value) <> "" Then
                      Response.Write "      <OPTION value='" & (rs_situacao_obra.Fields.Item("id").Value) & "'"
                      If Lcase(rs_situacao_obra.Fields.Item("id").Value) = Lcase(rs_altera_acomp.Fields.Item("cod_situacao_organizacao_obra").Value) then
                        Response.Write "selected"
                      End If
                      Response.Write ">" & (rs_situacao_obra.Fields.Item("dsc_situacao_obra").Value) & "</OPTION>"
                    End If
                    rs_situacao_obra.MoveNext()
                  Wend
                  If (rs_situacao_obra.CursorType > 0) Then
                    rs_situacao_obra.MoveFirst
                  Else
                    rs_situacao_obra.Requery
                  End If
                %>
              </select>
            </td>
            <td>
              <input type="text" name="dsc_nota_organizacao_obra" placeholder="Notas" value="<%=(rs_altera_acomp.Fields.Item("dsc_nota_organizacao_obra").Value)%>">
            </td>
          </tr>
        </table>
      </td>
    </tr>

    <tr valign="baseline">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC">
        <strong class="style5">Foi Realizada a Vistoria?</strong>
      </td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
        <label>
          <%
            value = rs_altera_acomp.Fields.Item("vistoria").Value
            checked = ""

            If value = True Then
              value = 1
              checked = "checked=""checked"""
            Else
              value = 0
            End If

            Response.Write "<input type=""checkbox"" onclick=""mostraEsconde('campo_vistoria','dta_vistoria');"" id=""campo_vistoria"" name=""vistoria"" "& checked &"/>"
          %>
        </label>
        <div id="dta_vistoria" style="<% If Not rs_altera_acomp.Fields.Item("vistoria").Value Then Response.Write "display:none;" End If %>">
          Data da Vistoria:
          <br/>
          <input name="dt_vistoria" type="text" id="dt_vistoria" class="datepicker" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" size="15" value="<%= rs_altera_acomp.Fields.Item("dt_vistoria").Value %>" />
        </div>
      </td>
    </tr>

    <tr valign="baseline">
      <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC">
        <strong class="style5">É Pendência?</strong>
      </td>
      <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
        <label>
          <%
            value = rs_altera_acomp.Fields.Item("flg_pendencia").Value
            checked = ""

            If value = True Then
              value = 1
              checked = "checked=""checked"""
            Else
              value = 0
            End If

            Response.Write "<input type=""checkbox"" onclick=""mostraEsconde('campo_pendencia','tipo_pendencia'); mostraEsconde('campo_pendencia','desc_pendencia'); mostraEsconde('campo_pendencia','dt_limite_pendencia');"" id=""campo_pendencia"" name=""flg_pendencia"" "& checked &"/>"
          %>
        </label>
        <div id="tipo_pendencia" style="<% If Not rs_altera_acomp.Fields.Item("flg_pendencia").Value Then Response.Write "display:none;" End If %>">
          Tipo de Pendência:
          <br/>
          <select name="cod_tipo_pendencia" style="width: 100%;">
            <option value=""></option>
            <%
              Set objCon = Server.CreateObject("ADODB.Connection")
                  objCon.Open MM_cpf_STRING

              strQ = "SELECT * FROM tb_tipo_pendencia "

              Set rs_combo = Server.CreateObject("ADODB.Recordset")
                rs_combo.CursorLocation = 3
                rs_combo.CursorType = 3
                rs_combo.LockType = 1
                rs_combo.Open strQ, objCon, , , &H0001

              While (NOT rs_combo.EOF)
                If Trim(rs_combo.Fields.Item("dsc_tipo_pendencia").Value) <> "" Then
                  Response.Write "      <OPTION value='" & (rs_combo.Fields.Item("id").Value) & "'"
                  If Lcase(rs_combo.Fields.Item("id").Value) = Lcase(rs_altera_acomp.Fields.Item("cod_tipo_pendencia").Value) then
                    Response.Write "selected"
                  End If
                  Response.Write ">" & (rs_combo.Fields.Item("dsc_tipo_pendencia").Value) & "</OPTION>"
                End If
                rs_combo.MoveNext()
              Wend
              If (rs_combo.CursorType > 0) Then
                rs_combo.MoveFirst
              Else
                rs_combo.Requery
              End If
            %>
          </select>
        </div>
        <div id="desc_pendencia" style="<% If Not rs_altera_acomp.Fields.Item("flg_pendencia").Value Then Response.Write "display:none;" End If %>">
          Descrição:
          <br/>
          <textarea name="dsc_pendencia" cols="32" maxlength="255"><%=(rs_altera_acomp.Fields.Item("dsc_pendencia").Value)%></textarea>
        </div>
        <div id="dt_limite_pendencia" style="<% If Not rs_altera_acomp.Fields.Item("flg_pendencia").Value Then Response.Write "display:none;" End If %>">
          Data Limíte da Pendência:
          <br/>
          <input name="dta_limite_pendencia" type="text" id="dta_limite_pendencia" value="<%=(rs_altera_acomp.Fields.Item("dta_limite_pendencia").Value)%>" class="datepicker" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" size="15" />
        </div>
      </td>
    </tr>
    
    <tr valign="baseline">
      <td colspan="2">
        <input type="submit" value="Salvar" style="float: right;">
      </td>
    </tr>
  </table>

  <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= rs_altera_acomp.Fields.Item("cod_acompanhamento").Value %>">
</form>