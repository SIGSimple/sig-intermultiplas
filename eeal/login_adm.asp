<%@LANGUAGE="VBSCRIPT"%>
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="1"
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
</script><style type="text/css">
<!--
.style3 {font-size: 16pt}
-->
</style><!--mstheme--><link rel="stylesheet" href="spri1011.css">
<meta name="Microsoft Theme" content="spring 1011">
</head>

<body onLoad="FP_preloadImgs(/*url*/'button2C1.jpg', /*url*/'button2D1.jpg', /*url*/'button2F1.jpg', /*url*/'button83.jpg', /*url*/'button84.jpg', /*url*/'button85.jpg', /*url*/'button88.jpg', /*url*/'button89.jpg', /*url*/'button3B.jpg', /*url*/'button3C.jpg', /*url*/'button3E.jpg', /*url*/'button3F.jpg', /*url*/'button90.jpg', /*url*/'button91.jpg', /*url*/'button92.jpg', /*url*/'button93.jpg', /*url*/'button94.jpg', /*url*/'button95.jpg', /*url*/'button126.jpg', /*url*/'button127.jpg', /*url*/'button200.jpg', /*url*/'button201.jpg', /*url*/'button162.jpg', /*url*/'button163.jpg', /*url*/'button183.jpg', /*url*/'button184.jpg', /*url*/'button186.jpg', /*url*/'button187.jpg')">

<div align="center">
  <div align="center"><font face="Arial"><strong><span class="style3">
	<font size="5">&Aacute;rea do Administrador </font></span></strong></font></div>
	<table x:str border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; width: 757px" id="table1">
    <colgroup>
      <col width="64" style="width: 48pt">
      <col width="64" style="width: 48pt">
      <col width="64" style="width: 48pt"><col width="64" style="width: 48pt">
    	<col width="64" style="width: 48pt">
    </colgroup>
    <tr height="18" style="height: 13.5pt">
      <td height="18" colspan="5" style="height: 13.5pt; width: 221px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px"></td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; width: 221px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px">		</td>
	    <td height="18" style="height: 13.5pt; width: 17px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px">		</td>
	    <td height="18" style="height: 13.5pt; width: 222px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px">		</td>
	    <td style="width: 220px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px">&nbsp;</td>
	    <td style="width: 332px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="221">
        <a href="const/df_consulta.asp">
		<img border="0" id="img15" src="button128.jpg" height="26" width="200" alt="Cadastro de Construtoras" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Cadastro de Construtoras" onMouseOver="FP_swapImg(1,0,/*id*/'img15',/*url*/'button126.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img15',/*url*/'button128.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img15',/*url*/'button127.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img15',/*url*/'button126.jpg')"></a></td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="17">		</td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="222">
	      <a href="predios/df_consulta.asp">
			<img border="0" id="img1" src="button2E1.jpg" height="26" width="200" alt="Cadastro de Pr�dios" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Cadastro de Pr�dios" onMouseOver="FP_swapImg(1,0,/*id*/'img1',/*url*/'button2F1.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img1',/*url*/'button2E1.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img1',/*url*/'button83.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img1',/*url*/'button2F1.jpg')"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="220">
	      <p align="center">
			<a href="cidades/df_consulta.asp">
			<img border="0" id="img14" src="button97.jpg" height="26" width="200" alt="Cadastro de Munic�pios" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Cadastro de Munic�pios" onMouseOver="FP_swapImg(1,0,/*id*/'img14',/*url*/'button90.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img14',/*url*/'button97.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img14',/*url*/'button91.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img14',/*url*/'button90.jpg')"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="332">
	      <a href="depto/df_consulta.asp">
			<img border="0" id="img12" src="button98.jpg" height="26" width="200" alt="Cadastro de Departamentos" onMouseOver="FP_swapImg(1,0,/*id*/'img12',/*url*/'button92.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img12',/*url*/'button98.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img12',/*url*/'button93.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img12',/*url*/'button92.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Cadastro de Departamentos"></a></td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="221">		</td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="17">		</td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="222">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="220">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="332">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="221">
        <a href="cad_usuario.asp">
		<img border="0" id="img13" src="button2B1.jpg" height="26" width="200" alt="Cadastro de usu�rios" onMouseOver="FP_swapImg(1,0,/*id*/'img13',/*url*/'button2C1.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img13',/*url*/'button2B1.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img13',/*url*/'button2D1.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img13',/*url*/'button2C1.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Cadastro de usu�rios"></a></td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="17">		</td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="222">
	      <a href="cadastro_pi.asp">
			<img border="0" id="img5" src="button100.jpg" height="26" width="200" alt="Cadastro de novo PI" onMouseOver="FP_swapImg(1,0,/*id*/'img5',/*url*/'button84.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img5',/*url*/'button100.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img5',/*url*/'button85.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img5',/*url*/'button84.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Cadastro de novo PI"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="220">
		<p align="center"></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="332">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="221">		</td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="17">		</td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="222">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="220">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="332">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="221">		
		<a href="cad_resp.asp">
		<img border="0" id="img16" src="button202.jpg" height="26" width="200" alt="Cadastro de Fiscais" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Cadastro de Fiscais" onMouseOver="FP_swapImg(1,0,/*id*/'img16',/*url*/'button200.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img16',/*url*/'button202.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img16',/*url*/'button201.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img16',/*url*/'button200.jpg')"></a></td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="17">		</td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="222">
	      <a href="filtro_adm.asp"><img src="button102.jpg" alt="Acompanhamento" name="img9" width="200" height="26" border="0" id="img9" onMouseDown="FP_swapImg(1,0,/*id*/'img9',/*url*/'button89.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img9',/*url*/'button88.jpg')" onMouseOver="FP_swapImg(1,0,/*id*/'img9',/*url*/'button88.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img9',/*url*/'button102.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Acompanhamento"></a></td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="220">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="332">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="221">		</td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="17">		</td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="222">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="220">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="332">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="221">		
		<a href="cad_mensagem.asp">
		<img border="0" id="img17" src="button164.jpg" height="26" width="200" alt="Alterar texto da p�gina Inicial" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Alterar texto da p�gina Inicial" onMouseOver="FP_swapImg(1,0,/*id*/'img17',/*url*/'button162.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img17',/*url*/'button164.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img17',/*url*/'button163.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img17',/*url*/'button162.jpg')"></a></td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="17">		</td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="222"><a href="filtro_acomp.asp">
		</a></a>
          <a href="filtro_med_constr.asp"><img border="0" id="img7" src="button3A.jpg" height="26" width="200" alt="Medi��o Construtoras" onMouseOver="FP_swapImg(1,0,/*id*/'img7',/*url*/'button3B.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img7',/*url*/'button3A.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img7',/*url*/'button3C.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img7',/*url*/'button3B.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Medi��o Construtoras"></a></td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="220">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="332">&nbsp;		</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="221">		</td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="17">		</td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="222">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="220">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="332">&nbsp;		</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="221">		
		<a href="cad_diretoria.asp">
		<img border="0" id="img18" src="button185.jpg" height="26" width="200" alt="Cadastro de Dir de Ensino" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Cadastro de Dir de Ensino" onMouseOver="FP_swapImg(1,0,/*id*/'img18',/*url*/'button183.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img18',/*url*/'button185.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img18',/*url*/'button184.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img18',/*url*/'button183.jpg')"></a></td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="17">		</td>
	    <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="222">
          <a href="msg_construcao.asp"><img src="button3D.jpg" alt="Medi��o Gerenciadora" name="img11" width="200" height="26" border="0" id="img11" onMouseDown="FP_swapImg(1,0,/*id*/'img11',/*url*/'button3F.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img11',/*url*/'button3E.jpg')" onMouseOver="FP_swapImg(1,0,/*id*/'img11',/*url*/'button3E.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img11',/*url*/'button3D.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Medi��o Gerenciadora"></a></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="220">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="332">&nbsp;		</td>
    </tr>
    <tr height="17" style="height: 12.75pt">
      <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="221">		</td>
	    <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="17">		</td>
	    <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="222">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="220">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="332">&nbsp;</td>
    </tr>
    <tr height="17" style="height: 12.75pt">
      <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="221">		
		<a href="cad_situacao.asp">
		<img border="0" id="img19" src="button188.jpg" height="26" width="200" alt="Cadastro de Situa��o" onMouseOver="FP_swapImg(1,0,/*id*/'img19',/*url*/'button186.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img19',/*url*/'button188.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img19',/*url*/'button187.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img19',/*url*/'button186.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Cadastro de Situa��o"></a></td>
	    <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="17">		</td>
	    <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="222">
          </td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="220">&nbsp;        </td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="332">
          <a href="rel_adm.asp"><img src="button103.jpg" alt="Relat�rios" name="img8" width="200" height="26" border="0" id="img8" onMouseDown="FP_swapImg(1,0,/*id*/'img8',/*url*/'button95.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img8',/*url*/'button94.jpg')" onMouseOver="FP_swapImg(1,0,/*id*/'img8',/*url*/'button94.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img8',/*url*/'button103.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Relat�rios"></a></td>
    </tr>
    <tr height="17" style="height: 12.75pt">
      <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="221">		</td>
	    <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="17">		</td>
	    <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="222">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="220">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="332">&nbsp;</td>
    </tr>
    <tr height="17" style="height: 12.75pt">
      <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="221">		</td>
	    <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="17">		</td>
	    <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="222">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="220">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="332">&nbsp;</td>
    </tr>
    <tr height="17" style="height: 12.75pt">
      <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="221">		</td>
	    <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="17">		</td>
	    <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="222">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="220">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="332">&nbsp;</td>
    </tr>
  </table>
</div>
</body>

</html>
