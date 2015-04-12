<%@LANGUAGE="VBSCRIPT"%>
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="1,4,3,2"
MM_authFailedURL="erro.asp"
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

<body onLoad="FP_preloadImgs(/*url*/'button135.jpg', /*url*/'button136.jpg', /*url*/'button189.jpg', /*url*/'button190.jpg', /*url*/'button192.jpg', /*url*/'button193.jpg', /*url*/'button195.jpg', /*url*/'button196.jpg', /*url*/'button1A2.jpg', /*url*/'button1B2.jpg')">

<div align="center">
  <div align="center"><span class="style3"><strong><font face="Arial" size="5">
	Relatórios</font></strong></span><font face="Arial"><strong><span class="style3"><font size="5"> </font> </span></strong></font></div>
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
        <a href="data.asp">
		<img border="0" id="img8" src="button137.jpg" height="26" width="200" alt="Medições Construtora por Data" onMouseOver="FP_swapImg(1,0,/*id*/'img8',/*url*/'button135.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img8',/*url*/'button137.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img8',/*url*/'button136.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img8',/*url*/'button135.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Medições Construtora por Data"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
      <a target="_top" href="rel_pi.asp">
		<img border="0" id="img9" src="button191.jpg" height="26" width="200" alt="Relação dos PIS do sistema" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Relação dos PIS do sistema" onmouseover="FP_swapImg(1,0,/*id*/'img9',/*url*/'button189.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img9',/*url*/'button191.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img9',/*url*/'button190.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img9',/*url*/'button189.jpg')"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
	  <a href="PI_situacao.asp">
		<img border="0" id="img10" src="button194.jpg" height="26" width="200" alt="PI por estágio de obra" onmouseover="FP_swapImg(1,0,/*id*/'img10',/*url*/'button192.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img10',/*url*/'button194.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img10',/*url*/'button193.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img10',/*url*/'button192.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="PI por estágio de obra"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
      <a href="rel_plan_contr_med.asp">
		<img border="0" id="img11" src="button197.jpg" height="26" width="200" alt="Relatório. Plan. Cont. Medições" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Relatório. Plan. Cont. Medições" onmouseover="FP_swapImg(1,0,/*id*/'img11',/*url*/'button195.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img11',/*url*/'button197.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img11',/*url*/'button196.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img11',/*url*/'button195.jpg')"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;
        </td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="17" style="height: 12.75pt">
      <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="17" style="height: 12.75pt">
      <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		
		<a href="rel_ultimas_ocor.asp">
		<img border="0" id="img12" src="button198.jpg" height="26" width="200" alt="Relatório Últimas Ocorrências" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Relatório Últimas Ocorrências" onmouseover="FP_swapImg(1,0,/*id*/'img12',/*url*/'button1A2.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img12',/*url*/'button198.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img12',/*url*/'button1B2.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img12',/*url*/'button1A2.jpg')"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
  </table>
</div>
</body>

</html>
