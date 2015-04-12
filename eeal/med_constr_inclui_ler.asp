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
  MM_editTable = "tb_Medicao_Construtora"
  MM_editRedirectUrl = "med_constr_inclui.asp"
  MM_fieldsStr  = "PI|value|Nmero_da_Medio|value|Valor_da_Medio|value|Porcentagem_de_Avano|value|_Medio_Final_|value|Valor_do_Contrato|value|Data_de_Envio_para_FDE|value|informacao_placa|value"
  MM_columnsStr = "PI|',none,''|n_medicao|none,none,NULL|vlr_medicao|',none,''|[Porcentagem de Avanço]|',none,''|[É Medição Final ?]|none,1,0|[Valor do Contrato]|',none,''|[Data de Envio para FDE]|',none,NULL|[Informação da Placa de Obra]|',none,''"

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
Dim rs_med_constr__MMColParam
rs_med_constr__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  rs_med_constr__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim rs_med_constr
Dim rs_med_constr_numRows

Set rs_med_constr = Server.CreateObject("ADODB.Recordset")
rs_med_constr.ActiveConnection = MM_cpf_STRING
rs_med_constr.Source = "SELECT *  FROM cLista_pi_med_contr  WHERE PI = '" + Replace(rs_med_constr__MMColParam, "'", "''") + "'  ORDER BY [N_Medicao] DESC"
rs_med_constr.CursorType = 0
rs_med_constr.CursorLocation = 2
rs_med_constr.LockType = 1
rs_med_constr.Open()

rs_med_constr_numRows = 0
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
rs_pi.Source = "SELECT tb_pi.PI, [tb_predio].[cod_predio] & ' - ' & [Nome_Unidade] AS Expr1  FROM tb_predio INNER JOIN tb_pi ON tb_predio.cod_predio = tb_pi.cod_predio  WHERE PI = '" + Replace(rs_pi__MMColParam, "'", "''") + "'  ORDER BY tb_pi.PI;  "
rs_pi.CursorType = 0
rs_pi.CursorLocation = 2
rs_pi.LockType = 1
rs_pi.Open()

rs_pi_numRows = 0
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
Recordset1.Source = "SELECT * FROM tb_pi WHERE PI = '" + Replace(Recordset1__MMColParam, "'", "''") + "'"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim rs_soma_med__MMColParam
rs_soma_med__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  rs_soma_med__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim rs_soma_med
Dim rs_soma_med_numRows

Set rs_soma_med = Server.CreateObject("ADODB.Recordset")
rs_soma_med.ActiveConnection = MM_cpf_STRING
rs_soma_med.Source = "SELECT tb_pi.PI, Sum(tb_Medicao_Construtora.vlr_medicao) AS SomaDevlr_medicao  FROM tb_Medicao_Construtora RIGHT JOIN tb_pi ON tb_Medicao_Construtora.PI = tb_pi.PI  WHERE tb_pi.PI = '" + Replace(rs_soma_med__MMColParam, "'", "''") + "'  GROUP BY tb_pi.PI;    "
rs_soma_med.CursorType = 0
rs_soma_med.CursorLocation = 2
rs_soma_med.LockType = 1
rs_soma_med.Open()

rs_soma_med_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rs_med_constr_numRows = rs_med_constr_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rs_med_constr_total
Dim rs_med_constr_first
Dim rs_med_constr_last

' set the record count
rs_med_constr_total = rs_med_constr.RecordCount

' set the number of rows displayed on this page
If (rs_med_constr_numRows < 0) Then
  rs_med_constr_numRows = rs_med_constr_total
Elseif (rs_med_constr_numRows = 0) Then
  rs_med_constr_numRows = 1
End If

' set the first and last displayed record
rs_med_constr_first = 1
rs_med_constr_last  = rs_med_constr_first + rs_med_constr_numRows - 1

' if we have the correct record count, check the other stats
If (rs_med_constr_total <> -1) Then
  If (rs_med_constr_first > rs_med_constr_total) Then
    rs_med_constr_first = rs_med_constr_total
  End If
  If (rs_med_constr_last > rs_med_constr_total) Then
    rs_med_constr_last = rs_med_constr_total
  End If
  If (rs_med_constr_numRows > rs_med_constr_total) Then
    rs_med_constr_numRows = rs_med_constr_total
  End If
End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = rs_med_constr
MM_rsCount   = rs_med_constr_total
MM_size      = rs_med_constr_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rs_med_constr_first = MM_offset + 1
rs_med_constr_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rs_med_constr_first > MM_rsCount) Then
    rs_med_constr_first = MM_rsCount
  End If
  If (rs_med_constr_last > MM_rsCount) Then
    rs_med_constr_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style3 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
.style9 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; }
.style17 {font-family: Arial, Helvetica, sans-serif; font-size: 12; font-weight: bold; }
.style27 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-weight: bold;
}
.style19 {font-family: Arial, Helvetica, sans-serif; font-size: 12; font-weight: bold; color: #000099; }
.style37 {font-size: 10}
.style44 {font-family: Arial, Helvetica, sans-serif; font-size: 16px; font-weight: bold; color: #000066; }
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
<div align="center">
  <table width="812" border="0">
    <tr bgcolor="#CCCCCC">
      <td width="139"><span class="style17"><%=(rs_pi.Fields.Item("PI").Value)%></span></td>
      <td width="663"><div align="left"><span class="style19"><%=(rs_pi.Fields.Item("Expr1").Value)%></span></div></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td><span class="style27">Interven&ccedil;&atilde;o FDE </span></td>
      <td><div align="left"><span class="style17"><%=(Recordset1.Fields.Item("Descrição da Intervenção FDE").Value)%></span></div></td>
    </tr>
  </table>
</div>
<form method="POST" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style9">PI:</span></td>
      <td bgcolor="#CCCCCC"><select name="PI">
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
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style9">Número da Medição:</span></td>
      <td bgcolor="#CCCCCC"><input name="Nmero_da_Medio" type="text" value="<% =0 %>" size="5" /></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style9">Valor da Medição:</span></td>
      <td bgcolor="#CCCCCC"><input name="Valor_da_Medio" type="text" value="<% =0 %>" size="18" /></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style9">Porcentagem de Avanço:</span></td>
      <td bgcolor="#CCCCCC"><input name="Porcentagem_de_Avano" type="text" value="<% =0 %>" size="12" /></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style9">É Medição Final ?:</span></td>
      <td bgcolor="#CCCCCC"><input <%If (CStr((rs_med_constr.Fields.Item("É MediçãoFinal ?").Value)) = CStr("True")) Then Response.Write("checked=""checked""") : Response.Write("")%> type="checkbox" name="_Medio_Final_" value=1 /></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style9">Valor do Contrato:</span></td>
      <td bgcolor="#CCCCCC"><input name="Valor_do_Contrato" type="text" value="<% =0 %>" size="18" /></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style9">Data de Envio para FDE:</span></td>
      <td bgcolor="#CCCCCC"><input name="Data_de_Envio_para_FDE" type="text" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" value="" size="15" /></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style9">Informa&ccedil;&atilde;o da Placa de Obra:</span></td>
      <td bgcolor="#CCCCCC"><textarea name="informacao_placa" cols="15" id="informacao_placa"><% ="" %>
      </textarea></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style37"></span></td>
      <td bgcolor="#CCCCCC">&nbsp;</td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1">
</form>


<div align="center">
  <table border="0">
    <tr bgcolor="#999999">
      <td><div align="center"><span class="style9">N&uacute;mero da Medi&ccedil;&atilde;o</span></div></td>
      <td><div align="center"><span class="style9">Valor da Medi&ccedil;&atilde;o</span></div></td>
      <td><div align="center"><span class="style9">Porcentagem de Avan&ccedil;o</span></div></td>
      <td><div align="center"><span class="style9">&Eacute; Medi&ccedil;&atilde;o Final ?</span></div></td>
      <td><div align="center"><span class="style9">Valor do Contrato</span></div></td>
      <td><span class="style9">Data de Envio para FDE</span></td>
      <td><span class="style9">Informa&ccedil;&atilde;o da Placa de Obra</span></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rs_med_constr.EOF)) %>
      <tr bgcolor="#CCCCCC">
        <td class="style3"><span class="style3"><%=(rs_med_constr.Fields.Item("n_medicao").Value)%></span></td>
        <td class="style3"><div align="right"><span class="style3"><%= FormatNumber((rs_med_constr.Fields.Item("vlrmedicao").Value), 2, -2, -2, -2) %></span></div></td>
        <td class="style3"><div align="center" class="style3"><%= FormatPercent((rs_med_constr.Fields.Item("Porcentagemde Avanço").Value), 2, -2, -2, -2) %></div></td>
        <td class="style3"><%=(rs_med_constr.Fields.Item("É MediçãoFinal ?").Value)%></td>
        <td class="style3"><div align="right"><%= FormatNumber((rs_med_constr.Fields.Item("Valordo Contrato").Value), 2, -2, -2, -2) %></div></td>
        <td class="style3"><span class="style3"><%=(rs_med_constr.Fields.Item("Data de Envio para FDE").Value)%></span></td>
        <td class="style3"><%=(rs_med_constr.Fields.Item("Informação da Placa de Obra").Value)%></td>
      </tr>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_med_constr.MoveNext()
Wend
%>
  </table>

  <table border="0">
    <tr>
      <td><% If MM_offset <> 0 Then %>
            <a href="<%=MM_moveFirst%>"><img src="First.gif" border=0></a>
            <% End If ' end MM_offset <> 0 %>
      </td>
      <td><% If MM_offset <> 0 Then %>
            <a href="<%=MM_movePrev%>"><img src="Previous.gif" border=0></a>
            <% End If ' end MM_offset <> 0 %>
      </td>
      <td><% If Not MM_atTotal Then %>
            <a href="<%=MM_moveNext%>"><img src="Next.gif" border=0></a>
            <% End If ' end Not MM_atTotal %>
      </td>
      <td><% If Not MM_atTotal Then %>
            <a href="<%=MM_moveLast%>"><img src="Last.gif" border=0></a>
            <% End If ' end Not MM_atTotal %>
      </td>
    </tr>
  </table>
</div>
  
<div align="center">
  <table width="278" border="0">
    <tr bgcolor="#CCCCCC">
      <td width="152"><span class="style44">Acumulado desta PI </span></td>
      <td width="116"><span class="style44"><%= FormatNumber((rs_soma_med.Fields.Item("SomaDeVlr_Medicao").Value), 2, -2, -2, -2) %></span></td>
    </tr>
  </table>
</div>
  <p>&nbsp;</p>
</body>
</html>
<%
rs_med_constr.Close()
Set rs_med_constr = Nothing
%>
<%
rs_pi.Close()
Set rs_pi = Nothing
%>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
rs_soma_med.Close()
Set rs_soma_med = Nothing
%>
