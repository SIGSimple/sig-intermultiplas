<%@LANGUAGE="VBSCRIPT"%>
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="1,2"
MM_authFailedURL="erro1.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (false Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Nova pagina 1</title>
<script language="JavaScript">
<!--
function FP_swapImg() {//v1.0
 var doc=document,args=arguments,elm,n; doc.$imgSwaps=new Array(); for(n=2; n<args.length;
 n+=2) { elm=FP_getObjectByID(args[n]); if(elm) { doc.$imgSwaps[doc.$imgSwaps.length]=elm;
 elm.$src=elm.src; elm.src=args[n+1]; } }
}

function FP_preloadImgs() {//v1.0
 var d=document,a=arguments; if(!d.FP_imgs) d.FP_imgs=new Array();
 for(var i=0; i<a.length; i++) { d.FP_imgs[i]=new Image; d.FP_imgs[i].src=a[i]; }
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
<style type="text/css">
<!--
.style3 {font-size: 16pt}
-->
</style>
<!--mstheme--><link rel="stylesheet" href="spri1011.css">
<meta name="Microsoft Theme" content="spring 1011">
</head>

<body onLoad="FP_preloadImgs(/*url*/'button114.jpg', /*url*/'button115.jpg', /*url*/'button116.jpg', /*url*/'button117.jpg', /*url*/'button6E.jpg', /*url*/'button6F.jpg', /*url*/'button118.jpg', /*url*/'button119.jpg', /*url*/'button174.jpg', /*url*/'button175.jpg', /*url*/'button180.jpg', /*url*/'button181.jpg', /*url*/'button205.jpg', /*url*/'button206.jpg', /*url*/'button1A3.jpg', /*url*/'button1B3.jpg')">

<div align="center">
  <div align="center"><font face="Arial"><strong><span class="style3">
	<font size="5">&Aacute;rea do Fiscal</font></span></strong></font></div>
	<table x:str border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; width: 638px" id="table1">
    <colgroup>
      <col width="64" style="width: 48pt"><col width="64" style="width: 48pt">
      <col width="64" style="width: 48pt">
    </colgroup>
    <tr height="18" style="height: 13.5pt">
      <td height="18" colspan="3" style="height: 13.5pt; width: 237px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px"></td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" colspan="3" style="height: 13.5pt; width: 237px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px">		</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
        <a href="const/df_consulta.asp">
		<img border="0" id="img1" src="button120.jpg" height="26" width="200" alt="Consulta de Construtoras" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Consulta de Construtoras" onMouseOver="FP_swapImg(1,0,/*id*/'img1',/*url*/'button116.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img1',/*url*/'button120.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img1',/*url*/'button117.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img1',/*url*/'button116.jpg')"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">
		<a href="rel_pi_filtro_do_fiscal.asp">
		<img border="0" id="img9" src="button176.jpg" height="26" width="200" alt="Busca por Fiscal" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Busca por Fiscal" onMouseOver="FP_swapImg(1,0,/*id*/'img9',/*url*/'button174.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img9',/*url*/'button176.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img9',/*url*/'button175.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img9',/*url*/'button174.jpg')"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;
		</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
		<a href="predios/df_consulta.asp">
		<img border="0" id="img5" src="button121.jpg" height="26" width="200" alt="Cadastro de Pr�dios" onMouseOver="FP_swapImg(1,0,/*id*/'img5',/*url*/'button114.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img5',/*url*/'button121.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img5',/*url*/'button115.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img5',/*url*/'button114.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Cadastro de Pr�dios"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">
		<a href="busca_pi_unidade_Fiscal.asp">
		<img border="0" id="img11" src="button207.jpg" height="26" width="200" alt="Busca por Nome" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Busca por Nome" onMouseOver="FP_swapImg(1,0,/*id*/'img11',/*url*/'button205.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img11',/*url*/'button207.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img11',/*url*/'button206.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img11',/*url*/'button205.jpg')"></a></td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
        <a href="filtro_acomp.asp"><img border="0" id="img7" src="button6D.jpg" height="26" width="200" alt="Acompanhamento" onMouseOver="FP_swapImg(1,0,/*id*/'img7',/*url*/'button6E.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img7',/*url*/'button6D.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img7',/*url*/'button6F.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img7',/*url*/'button6E.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Acompanhamento"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">
		<a href="filtro_cod_predio_fiscal.asp">
		<img border="0" id="img10" src="button182.jpg" height="26" width="200" alt="Busca por C�digo" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Busca por C�digo" onMouseOver="FP_swapImg(1,0,/*id*/'img10',/*url*/'button180.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img10',/*url*/'button182.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img10',/*url*/'button181.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img10',/*url*/'button180.jpg')"></a></td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
        </td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">
        <a href="busca_pi_cod_Fiscal.asp">
		<img border="0" id="img12" src="button211.jpg" height="26" width="200" alt="Busca por n� de PI" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Busca por n� de PI" onMouseOver="FP_swapImg(1,0,/*id*/'img12',/*url*/'button1A3.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img12',/*url*/'button211.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img12',/*url*/'button1B3.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img12',/*url*/'button1A3.jpg')"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="17" style="height: 12.75pt">
      <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="17" style="height: 12.75pt">
      <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">
		<a href="rel_medicao.asp">
        <img border="0" id="img8" src="button122.jpg" height="26" width="200" alt="Relat�rios" onMouseOver="FP_swapImg(1,0,/*id*/'img8',/*url*/'button118.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img8',/*url*/'button122.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img8',/*url*/'button119.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img8',/*url*/'button118.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Relat�rios"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
  </table>
</div>
</body>

</html>
