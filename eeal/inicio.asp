<%@LANGUAGE="VBSCRIPT"%>
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "default.asp"
  ' redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
  If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
    MM_newQS = "?"
    For Each Item In Request.QueryString
      If (Item <> "MM_Logoutnow") Then
        If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
        MM_newQS = MM_newQS & Item & "=" & Server.URLencode(Request.QueryString(Item))
      End If
    Next
    if (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
  End If
  Response.Redirect(MM_logoutRedirectPage)
End If
%>
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="1,2,3,4"
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
<!--#include file="Connections/cpf.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_cpf_STRING
Recordset1_cmd.CommandText = "SELECT * FROM tb_mensagem" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>CONTROLE FDE</title>
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
.style2 {
	font-family: Arial, Helvetica, sans-serif;
	color: #990033;
}
-->
</style>
</head>

<body onLoad="FP_preloadImgs(/*url*/'button8D.jpg', /*url*/'button8E.jpg', /*url*/'button9A.jpg', /*url*/'button9B.jpg', /*url*/'button7.jpg', /*url*/'button8.jpg', /*url*/'button10.jpg', /*url*/'button11.jpg', /*url*/'button129.jpg', /*url*/'button130.jpg')" style="text-align: center">

<table border="0" cellpadding="0" cellspacing="0" width="799" height="543">
	<!-- MSTableType="layout" -->
	<tr>
		<td valign="top">
		<!-- MSCellType="DecArea" -->
		<table cellpadding="0" cellspacing="0" width="935" height="98">
			<!-- MSCellFormattingTableID="1" -->
			<tr>
				<td width="925">
				<table cellpadding="0" cellspacing="0" border="0" width="925" height="88">
					<tr>
						<td valign="top" height="88" width="925">
						<!-- MSCellFormattingType="content" -->
						<table border="1" width="100%" id="table1" style="border-width: 0px">
	<tr>
		<td style="border-style: none; border-width: medium" width="8">&nbsp;</td>
		<td width="901" style="border-style: none; border-width: medium">
		<p style="margin-top: 0; margin-bottom: 0"><a href="<%= MM_Logout %>"></a></p>
		<table width="898" border="1" style="margin-bottom: 0">
          <tr>
            <td width="95"><img border="0" src="imagens/novo/fde3.jpg" width="111" height="35"></td>
            <td width="319"><span style="margin-top: 0; margin-bottom: 0"><img src="imagens/logo_Arcadis.JPG" width="318" height="41"></span></td>
            <td width="356"><span style="margin-top: 0; margin-bottom: 0"><span class="style2"><%=(Recordset1.Fields.Item("mensagem").Value)%></span></span></td>
            <td width="100"><span style="margin-top: 0; margin-bottom: 0"><font face="Bauhaus 93" size="4"><a href="<%= MM_Logout %>"><img src="button12.jpg" alt="Logout" name="img6" width="100" height="20" border="0" align="right" id="img6" onMouseDown="FP_swapImg(1,0,/*id*/'img6',/*url*/'button11.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img6',/*url*/'button10.jpg')" onMouseOver="FP_swapImg(1,0,/*id*/'img6',/*url*/'button10.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img6',/*url*/'button12.jpg')" fp-style="fp-btn: Braided Row 5" fp-title="Logout"></a></font></span></td>
          </tr>
        </table>
		</td>
		</tr>
	</table>
	<p align="left">
						<a target="I1" href="login_adm.asp">
						<span style="text-decoration: none">
						<img border="0" id="img5" src="button9.jpg" height="28" width="120" alt="Administrador" onMouseOver="FP_swapImg(1,0,/*id*/'img5',/*url*/'button7.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img5',/*url*/'button9.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img5',/*url*/'button8.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img5',/*url*/'button7.jpg')" fp-style="fp-btn: Embossed Tab 5; fp-proportional: 0" fp-title="Administrador"></span></a><a target="I1" href="login_fiscal.asp"><img border="0" id="img1" src="button8C.jpg" height="28" width="120" alt="Fiscal de Obras" fp-style="fp-btn: Embossed Tab 5; fp-proportional: 0" fp-title="Fiscal de Obras" onMouseOver="FP_swapImg(1,0,/*id*/'img1',/*url*/'button8D.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img1',/*url*/'button8C.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img1',/*url*/'button8E.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img1',/*url*/'button8D.jpg')"></a><a target="I1" href="login_medicao.asp"><img border="0" id="img2" src="button99.jpg" height="28" width="120" alt="Medições" onMouseOver="FP_swapImg(1,0,/*id*/'img2',/*url*/'button9A.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img2',/*url*/'button99.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img2',/*url*/'button9B.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img2',/*url*/'button9A.jpg')" fp-style="fp-btn: Embossed Tab 5; fp-proportional: 0" fp-title="Medições"></a><a target="I1" href="login_fde.asp"><img border="0" id="img7" src="button131.jpg" height="28" width="120" alt="FDE" fp-style="fp-btn: Embossed Tab 5; fp-proportional: 0" fp-title="FDE" onMouseOver="FP_swapImg(1,0,/*id*/'img7',/*url*/'button129.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img7',/*url*/'button131.jpg')" onMouseDown="FP_swapImg(1,0,/*id*/'img7',/*url*/'button130.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img7',/*url*/'button129.jpg')"></a></td>
					</tr>
				</table>
				</td>
				<td height="88" width="10"><font face="Bauhaus 93" size="6">
				<img alt="" width="10" height="88" src="MsoPnl_sh_b_14.jpg"></font></td>
			</tr>
			<tr>
				<td colspan="2" height="10"><font face="Bauhaus 93" size="6">
				<img alt="" width="935" height="10" src="MsoPnl_sh_r_13.jpg"></font></td>
			</tr>
		</table>
		<p>
		<iframe name="I1" src="inicial.htm" width="944" height="675" border="0" frameborder="0" style="border-style: outset; border-width: 3px; background-color: #E2E2E2">
		Seu navegador não oferece suporte para quadros embutidos ou está configurado para não exibi-los.
		</iframe>
		</td>
	</tr>
	<tr>
		<td valign="top" width="799">
		<!-- MSCellType="NavBody" -->
		&nbsp;</td>
	</tr>
</table>
<p align="center">&nbsp;</p>

</body>

</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>