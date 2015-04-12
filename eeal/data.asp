<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.Form("data") <> "") Then 
  Recordset1__MMColParam = Request.Form("data")
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_cpf_STRING
Recordset1.Source = "SELECT *  FROM cDataRel  WHERE data = " + Replace(Recordset1__MMColParam, "'", "''") + " or data LIKE '%" + Replace(Recordset1__MMColParam, "'", "''") + "%' and cod_predio is not null"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim rs_datas
Dim rs_datas_numRows

Set rs_datas = Server.CreateObject("ADODB.Recordset")
rs_datas.ActiveConnection = MM_cpf_STRING
rs_datas.Source = "SELECT *  FROM cdatas"
rs_datas.CursorType = 0
rs_datas.CursorLocation = 2
rs_datas.LockType = 1
rs_datas.Open()

rs_datas_numRows = 0
%>
<%
Dim Recordset2__MMColParam
Recordset2__MMColParam = "1"
If (Request.Form("data") <> "") Then 
  Recordset2__MMColParam = Request.Form("data")
End If
%>
<%
Dim Recordset2
Dim Recordset2_numRows

Set Recordset2 = Server.CreateObject("ADODB.Recordset")
Recordset2.ActiveConnection = MM_cpf_STRING
Recordset2.Source = "SELECT count(pi) as contar  FROM cDataRel  WHERE data = " + Replace(Recordset2__MMColParam, "'", "''") + " or data LIKE '%" + Replace(Recordset2__MMColParam, "'", "''") + "%' and cod_predio is not null"
Recordset2.CursorType = 0
Recordset2.CursorLocation = 2
Recordset2.LockType = 1
Recordset2.Open()

Recordset2_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10000000000
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
Recordset2_numRows = Recordset2_numRows + Repeat2__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:(null)0="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" >
<link rel="File-List" href="data_arquivos/filelist.xml">

<title>Untitled Document</title>
<style type="text/css">
<!--
.style3 {font-family: Arial, Helvetica, sans-serif; font-size: 9px; }
.style5 {font-family: Arial, Helvetica, sans-serif; font-size: 9px; color: #FFFFFF; }
.style6 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 14px;
}
.style7 {
	color: #FFFFFF;
	font-weight: bold;
}
.style13 {color: #000000}
.style14 {font-size: 10; font-weight: bold;}
-->
</style>
<script type="text/JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<!--[if !mso]>
<style>
v\:*         { behavior: url(#default#VML) }
o\:*         { behavior: url(#default#VML) }
.shape       { behavior: url(#default#VML) }
</style>
<![endif]--><!--[if gte mso 9]>
<xml><o:shapedefaults v:ext="edit" spidmax="1027"/>
</xml><![endif]-->
</head>

<body onload="MM_preloadImages('imagens/Linha Lateral.jpg')">
<table width="917" border="0">
  <tr>
    <td width="232"><form id="form1" name="form1" method="post" action="">
      <span class="style6">Selecione a Data de Envio FDE </span>
      <select name="data" id="data" style="font-size: 10px; font-family: Arial" size="1">
        <%
While (NOT rs_datas.EOF)
%><option value="<%=(rs_datas.Fields.Item("dt_envio").Value)%>"><%=(rs_datas.Fields.Item("dt_envio").Value)%></option>
        <%
  rs_datas.MoveNext()
Wend
If (rs_datas.CursorType > 0) Then
  rs_datas.MoveFirst
Else
  rs_datas.Requery
End If
%>
      </select>
      <label for="Submit"></label>
      <input type="submit" name="Submit" value="Ok" id="Submit" />
    </form>    </td>
    <td width="88"><a href=# onclick="Javascript:print()"><img src=imagens/imprimir.gif width="43" height="28" /></a></td>
    <td width="73"><div align="center"><a href="qtde_pi.asp" target="_blank" class="style6"></a></div></td>
    <td width="82">
	<p align="right"><!--[if gte vml 1]><v:shapetype id="_x0000_t202"
 coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
 <v:stroke joinstyle="miter"/>
 <v:path gradientshapeok="t" o:connecttype="rect"/>
</v:shapetype><v:shape id="_x0000_s1028" type="#_x0000_t202" alt="" style='position:absolute;
 left:446.25pt;top:18pt;width:74.25pt;height:19.5pt;z-index:1;float:right'
 filled="f" stroked="f">
 <v:textbox>
<table cellspacing="0" cellpadding="0" width="100%" height="100%">
	<tr>
		<td align="center">
		<p align="right"><i><b><font face="Arial" size="2"></font></b></i></td>
	</tr>
</table>
 </v:textbox>
</v:shape><![endif]--><![if !vml]><span style='mso-ignore:vglayout;position:
absolute;z-index:1;left:595px;top:24px;width:103px;height:30px'><img width=103
height=30 src="data_arquivos/image001.gif" align=right v:shapes="_x0000_s1028"></span><![endif]></p>
	<p align="center"></td>
    <td width="475">
	<img border="0" src="imagens/logo_Arcadis.jpg"></td>
  </tr>
  <tr>
    <td colspan="6"><!--[if gte vml 1]><v:line
 id="_x0000_s1025" alt="" style='position:absolute;left:0;text-align:left;
 top:0;z-index:1' from="12pt,45pt" to="693.75pt,45pt" strokecolor="#dccba3"
 strokeweight="3.75pt"/><![endif]--><![if !vml]><span style='mso-ignore:vglayout;
position:absolute;z-index:1;left:13px;top:57px;width:915px;height:6px'><img
width=915 height=6 src="data_arquivos/image002.gif" v:shapes="_x0000_s1025"></span><![endif]>&nbsp;<a href="#" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image4','','imagens/Linha Lateral.jpg',1)"></a></td>
  </tr>
</table>
<table border="0">
  <tr bgcolor="#999999">
    <td width="50"><div align="center"><span class="style5">cod_predio</span></div></td>
    <td width="200"><div align="center"><span class="style5">Nome_Unidade</span></div></td>
    <td width="60"><div align="center"><span class="style5">PI</span></div></td>
    <td width="60"><div align="center"><span class="style5">n&ordm; da medicao</span></div></td>
    <td width="60"><div align="center"><span class="style5">Data de Envio FDE </span></div></td>
    <td width="60"><div align="center"><span class="style5">Valor da Medi&ccedil;&atilde;o</span></div></td>
    <td width="60"><div align="center"><span class="style5">Porcentagem de Avan&ccedil;o</span></div></td>
    <td width="60"><div align="center"><span class="style5">Data Prevista para o T&eacute;rmino</span></div></td>
    <td width="60"><div align="center"><span class="style5">&Eacute; Medi&ccedil;&atilde;o Final ?</span></div></td>
    <td width="60"><div align="center"><span class="style5">Valor do Contrato</span></div></td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) %>
    <tr bgcolor="#CCCCCC">
      <td><span class="style3"><%=(Recordset1.Fields.Item("cod_predio").Value)%></span></td>
      <td><span class="style3"><%=(Recordset1.Fields.Item("Nome_Unidade").Value)%></span></td>
      <td><span class="style3"><%=(Recordset1.Fields.Item("PI").Value)%></span></td>
      <td><div align="center"><span class="style3"><%=(Recordset1.Fields.Item("n_medicao").Value)%></span></div></td>
      <td><div align="center"><span class="style3"><%=(Recordset1.Fields.Item("data").Value)%></span></div></td>
      <td><div align="right"><span class="style3"><%= FormatNumber((Recordset1.Fields.Item("vlr_medicao").Value), 2, -2, -2, -2) %></span></div></td>
      <td><div align="center" class="style3"><%= FormatPercent((Recordset1.Fields.Item("Porcentagem_Avanço").Value), 2, -2, -2, -2) %></div></td>
      <td><div align="center"><span class="style3"><%=(Recordset1.Fields.Item("Data Prevista para o Término").Value)%></span></div></td>
      <td><div align="center" class="style3"><%=(Recordset1.Fields.Item("É MediçãoFinal ?").Value)%></div></td>
      <td><div align="right"><span class="style3"><%= FormatNumber((Recordset1.Fields.Item("Valor do Contrato").Value), 2, -2, -2, -2) %></span></div></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
</table>
<h5 class="style6"><%=(Recordset2.Fields.Item("contar").Value)%> Registro(s) encontrado(s) </h5>
<p><!--[if gte vml 1]><v:line
 id="_x0000_s1027" alt="" style='position:absolute;left:0;text-align:left;
 top:0;z-index:2' from="693pt,15pt" to="693pt,840.75pt" strokecolor="#dccba3"
 strokeweight="3.75pt"/><![endif]--><![if !vml]><span style='mso-ignore:vglayout;
position:absolute;z-index:2;left:921px;top:17px;width:6px;height:1107px'><img
width=6 height=1107 src="data_arquivos/image003.gif" v:shapes="_x0000_s1027"></span>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
rs_datas.Close()
Set rs_datas = Nothing
%>
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>