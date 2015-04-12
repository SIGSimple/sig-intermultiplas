<%@LANGUAGE="VBSCRIPT"%>
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="1,4"
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
function mmLoadMenus() {
  if (window.mm_menu_0819160203_0) return;
    window.mm_menu_0819160203_0 = new Menu("root",90,18,"",12,"#000000","#FFFFFF","#CCCCCC","#000084","left","middle",3,0,1000,-5,7,true,false,true,0,true,true);
  mm_menu_0819160203_0.addMenuItem("Por&nbsp;PI","location='semaforico.asp'");
  mm_menu_0819160203_0.addMenuItem("Por&nbsp;Prédio","location='semaforico_predio.asp'");
   mm_menu_0819160203_0.hideOnMouseOut=true;
   mm_menu_0819160203_0.bgColor='#555555';
   mm_menu_0819160203_0.menuBorder=1;
   mm_menu_0819160203_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0819160203_0.menuBorderBgColor='#777777';

                      window.mm_menu_0819161142_0 = new Menu("root",264,18,"",12,"#000000","#FFFFFF","#CCCCCC","#CCCCCC","right","middle",3,0,1000,-5,7,true,false,true,0,true,true);
  mm_menu_0819161142_0.addMenuItem("Semafórico&nbsp;-&nbsp;TODOS","window.open('semaforico_filtro_todos.asp', '_parent');");
  mm_menu_0819161142_0.addMenuItem("Por&nbsp;PI&nbsp;-&nbsp;Filtro&nbsp;por&nbsp;Estágio&nbsp;de&nbsp;Obra","window.open('semaforico_filtro_data.asp', '_parent');");
  mm_menu_0819161142_0.addMenuItem("Por&nbsp;Prédio&nbsp;-&nbsp;Filtro&nbsp;por&nbsp;Estágio&nbsp;de&nbsp;Obra","window.open('semaforico_filtro_data_predio.asp', '_parent');");
   mm_menu_0819161142_0.fontWeight="bold";
   mm_menu_0819161142_0.hideOnMouseOut=true;
   mm_menu_0819161142_0.bgColor='#555555';
   mm_menu_0819161142_0.menuBorder=1;
   mm_menu_0819161142_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0819161142_0.menuBorderBgColor='#777777';

    window.mm_menu_1012095609_0 = new Menu("root",205,18,"",12,"#000000","#FFFFFF","#CCCCCC","#000000","left","middle",3,0,1000,-5,7,true,false,true,0,true,true);
  mm_menu_1012095609_0.addMenuItem("SEMAFÓRICO&nbsp;-&nbsp;por&nbsp;PI","window.open('semaforico_filtro_data.asp', '_blank');");
  mm_menu_1012095609_0.addMenuItem("SEMAFÓRICO&nbsp;-&nbsp;por&nbsp;Prédio","window.open('semaforico_filtro_data_predio.asp', '_blank');");
  mm_menu_1012095609_0.addMenuItem("SEMAFÓRICO&nbsp;-&nbsp;Todos","window.open('semaforico_filtro_todos.asp', '_blank');");
   mm_menu_1012095609_0.hideOnMouseOut=true;
   mm_menu_1012095609_0.bgColor='#555555';
   mm_menu_1012095609_0.menuBorder=1;
   mm_menu_1012095609_0.menuLiteBgColor='#FFFFFF';
   mm_menu_1012095609_0.menuBorderBgColor='#777777';

mm_menu_1012095609_0.writeMenus();
} // mmLoadMenus()

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
<!--mstheme--><link rel="stylesheet" href="imagens/spri1011.css">
<meta name="Microsoft Theme" content="spring 1011">
<script language="JavaScript" src="mm_menu.js"></script>
</head>

<body onLoad="FP_preloadImgs(/*url*/'buttonE1.jpg', /*url*/'buttonF1.jpg', /*url*/'button135.jpg', /*url*/'button136.jpg', /*url*/'button144.jpg', /*url*/'button145.jpg', /*url*/'button147.jpg', /*url*/'button148.jpg', /*url*/'button247.jpg', /*url*/'button248.jpg', /*url*/'button150.jpg', /*url*/'button151.jpg', /*url*/'button153.jpg', /*url*/'button154.jpg', /*url*/'button165.jpg', /*url*/'button166.jpg', /*url*/'button212.jpg', /*url*/'button213.jpg', /*url*/'button03.jpg', /*url*/'button04.jpg')">
<script language="JavaScript1.2">mmLoadMenus();</script>
<div align="center">
  <div align="center"><span class="style3"><strong><font face="Arial" size="5"> Relatórios</font></strong></span><font face="Arial"><strong><span class="style3"><font size="5"> </font> </span></strong></font></div>
  <table x:str border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; width: 638px" id="table1">
    <colgroup>
      <col width="64" style="width: 48pt">
      <col width="64" style="width: 48pt">
      <col width="64" style="width: 48pt">
    </colgroup>
    <tr height="18" style="height: 13.5pt">
      <td height="18" colspan="3" style="height: 13.5pt; width: 237px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px"></td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; width: 237px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px"></td>
      <td style="width: 26px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px">&nbsp;</td>
      <td style="width: 185px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237"><a href="data.asp"> <img border="0" id="img8" src="button137.jpg" height="26" width="200" alt="Medições Construtora por Data" onMouseOver="FP_swapImg(1,0,/*id*/'img8',/*url*/'button135.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img8',/*url*/'button137.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img8',/*url*/'button136.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img8',/*url*/'button135.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Medições Construtora por Data"></a></td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">
		<a href="rel_plan_contr_med.asp">
		<img border="0" id="img13" src="button155.jpg" height="26" width="200" alt="Plan. Cont. Medições" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Plan. Cont. Medições" onMouseOver="FP_swapImg(1,0,/*id*/'img13',/*url*/'button153.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img13',/*url*/'button155.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img13',/*url*/'button154.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img13',/*url*/'button153.jpg')"></a></td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237"></td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
		<img src="buttonD1.jpg" alt="Semafórico" name="img7" width="200" height="26" border="0" id="img7" onMouseDown="FP_swapImg(1,0,/*id*/'img7',/*url*/'buttonF1.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img7',/*url*/'buttonE1.jpg')" onMouseOver="FP_swapImg(1,0,/*id*/'img7',/*url*/'buttonE1.jpg');MM_showMenu(window.mm_menu_1012095609_0,197,23,null,'img7')" onMouseOut="FP_swapImg(0,0,/*id*/'img7',/*url*/'buttonD1.jpg');MM_startTimeout();" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Semafórico"></td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">
		<a href="rel_ultimas_ocor.asp">
		<img border="0" id="img14" src="button167.jpg" height="26" width="200" alt="Últimas Ocorrências" onMouseOver="FP_swapImg(1,0,/*id*/'img14',/*url*/'button165.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img14',/*url*/'button167.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img14',/*url*/'button166.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img14',/*url*/'button165.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Últimas Ocorrências"></a></td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237"></td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237"><a href="PI_situacao.asp" target="_blank"> <img src="button146.jpg" alt="PIs por Estágio da Obra" name="img9" width="200" height="26" border="0" id="img9" onMouseDown="FP_swapImg(1,0,/*id*/'img9',/*url*/'button145.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img9',/*url*/'button144.jpg')" onMouseOver="FP_swapImg(1,0,/*id*/'img9',/*url*/'button144.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img9',/*url*/'button146.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="PIs por Estágio da Obra"></a></td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">
		<a href="data-.asp">
		<img border="0" id="img15" src="button02.jpg" height="26" width="200" alt="Medições por PI" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Medições por PI" onmouseover="FP_swapImg(1,0,/*id*/'img15',/*url*/'button03.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img15',/*url*/'button02.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img15',/*url*/'button04.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img15',/*url*/'button03.jpg')"></a></td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237"></td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
		<a target="_parent" href="rel_pi.asp">
		<img src="button149.jpg" alt="Pis do Sistema" name="img10" width="200" height="26" border="0" id="img10" onMouseDown="FP_swapImg(1,0,/*id*/'img10',/*url*/'button148.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img10',/*url*/'button147.jpg')" onMouseOver="FP_swapImg(1,0,/*id*/'img10',/*url*/'button147.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img10',/*url*/'button149.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Pis do Sistema"></a></td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237"></td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
		<a href="busca_pi_unidade.asp">
		<img border="0" id="img11" src="button249.jpg" height="26" width="200" alt="Busca por Nome" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Busca por Nome" onMouseOver="FP_swapImg(1,0,/*id*/'img11',/*url*/'button247.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img11',/*url*/'button249.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img11',/*url*/'button248.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img11',/*url*/'button247.jpg')"></a></td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="17" style="height: 12.75pt">
      <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237"></td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr>
      <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
		<a href="busca_pi_cod.asp">
		<img border="0" id="img12" src="button152.jpg" height="26" width="200" alt="Busca por nº de PI" onMouseOver="FP_swapImg(1,0,/*id*/'img12',/*url*/'button150.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img12',/*url*/'button152.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img12',/*url*/'button151.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img12',/*url*/'button150.jpg')" fp-style="fp-btn: Embossed Capsule 5; fp-proportional: 0" fp-title="Busca por nº de PI"></a></td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    
    <tr height="17" style="height: 12.75pt">
      <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
		<a href="filtro_adm.asp"></a></td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: general; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
  </table>
</div>
</body>

</html>
