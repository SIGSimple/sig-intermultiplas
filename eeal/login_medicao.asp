<%@LANGUAGE="VBSCRIPT"%>
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="1,3"
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

<body onLoad="FP_preloadImgs(/*url*/'button4D.jpg', /*url*/'button4E.jpg', /*url*/'button104.jpg', /*url*/'button105.jpg', /*url*/'button108.jpg', /*url*/'button109.jpg', /*url*/'button110.jpg', /*url*/'button5A.jpg', /*url*/'button5C.jpg', /*url*/'button5D.jpg', /*url*/'buttonB1.jpg', /*url*/'buttonC1.jpg', /*url*/'button177.jpg', /*url*/'button178.jpg', /*url*/'button208.jpg', /*url*/'button209.jpg')">

<div align="center">
  <div align="center"><font face="Arial"><strong><span class="style3">
	<font size="5">&Aacute;rea de  Medi&ccedil;&otilde;es </font> </span></strong></font></div>
	<table x:str border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; width: 638px" id="table1">
    <colgroup>
      <col width="64" style="width: 48pt"><col width="64" style="width: 48pt">
      <col width="64" style="width: 48pt">
    </colgroup>
    <tr height="18" style="height: 13.5pt">
      <td height="18" colspan="3" style="height: 13.5pt; width: 237px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px"></td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; width: 237px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px">		</td>
	    <td style="width: 26px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px">&nbsp;</td>
	    <td style="width: 185px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
        <a href="const/df_consulta.asp">
		<img src="button4C.jpg" alt="Consulta de Construtoras" name="img1" width="200" height="26" border="0" id="img1" onMouseDown="FP_swapImg(1,0,/*id*/'img1',/*url*/'button4E.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img1',/*url*/'button4D.jpg')" onMouseOver="FP_swapImg(1,0,/*id*/'img1',/*url*/'button4D.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img1',/*url*/'button4C.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Consulta de Construtoras"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">
		<a href="filtro_cod_predio.asp">
		<img border="0" id="img11" src="button179.jpg" height="26" width="200" alt="Busca por Código" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Busca por Código" onMouseOver="FP_swapImg(1,0,/*id*/'img11',/*url*/'button177.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img11',/*url*/'button179.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img11',/*url*/'button178.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img11',/*url*/'button177.jpg')"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
        <a href="predios/df_consulta.asp">
		<img border="0" id="img5" src="button4F.jpg" height="26" width="200" alt="Cadastro de Prédios" onMouseOver="FP_swapImg(1,0,/*id*/'img5',/*url*/'button104.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img5',/*url*/'button4F.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img5',/*url*/'button105.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img5',/*url*/'button104.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Cadastro de Prédios"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">
		<a href="rel_pi_filtro_de_medicoes.asp">
		<img border="0" id="img10" src="buttonA1.jpg" height="26" width="200" alt="Busca por Fiscal" onMouseOver="FP_swapImg(1,0,/*id*/'img10',/*url*/'buttonB1.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img10',/*url*/'buttonA1.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img10',/*url*/'buttonC1.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img10',/*url*/'buttonB1.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Busca por Fiscal"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
		<a href="filtro_med_constr.asp"><img src="button112.jpg" alt="Medição Construtoras" name="img9" width="200" height="26" border="0" id="img9" onMouseDown="FP_swapImg(1,0,/*id*/'img9',/*url*/'button109.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img9',/*url*/'button108.jpg')" onMouseOver="FP_swapImg(1,0,/*id*/'img9',/*url*/'button108.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img9',/*url*/'button112.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Medição Construtoras"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">
		<a href="busca_pi_unidade_FM.asp">
		<img border="0" id="img12" src="button210.jpg" height="26" width="200" alt="Busca por Nome" onMouseOver="FP_swapImg(1,0,/*id*/'img12',/*url*/'button208.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img12',/*url*/'button210.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img12',/*url*/'button209.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img12',/*url*/'button208.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Busca por Nome"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
		<a href="filtro_acomp.asp"><img border="0" id="img7" src="button113.jpg" height="26" width="200" alt="Acompanhamento" onMouseOver="FP_swapImg(1,0,/*id*/'img7',/*url*/'button110.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img7',/*url*/'button113.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img7',/*url*/'button5A.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img7',/*url*/'button110.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Acompanhamento"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">
		<a href="busca_pi_cod_FM.asp">
		<img border="0" src="button161.jpg" width="200" height="26"></a>&nbsp;		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
        </td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">
          <a href="rel_medicao.asp"><img src="button5B.jpg" alt="Relatórios" name="img8" width="200" height="26" border="0" id="img8" onMouseDown="FP_swapImg(1,0,/*id*/'img8',/*url*/'button5D.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img8',/*url*/'button5C.jpg')" onMouseOver="FP_swapImg(1,0,/*id*/'img8',/*url*/'button5C.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img8',/*url*/'button5B.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Relatórios"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="17" style="height: 12.75pt">
      <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="17" style="height: 12.75pt">
      <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
  </table>
</div>
</body>

</html>
