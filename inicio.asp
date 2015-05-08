<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
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
MM_authorizedUsers="1,2,3,4,5,6,7,8"
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

If Request.Form("profile_id") <> "" Then
	If Request.Form("profile_id") = 8 Then
		Response.Redirect("inicio-daee.asp")
	Else
		If Request.Form("profile_id") = 9 Then
			Response.Redirect("atendimento-prefeito.asp")
		Else
			Session("MM_UserAuthorization") = Request.Form("profile_id")
		End If
	End If
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
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>DAEE</title>
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
<script src="//code.jquery.com/jquery-1.11.2.min.js"></script>
<script type="text/javascript">
	$(document).ready(function(){
		$("button#btn_area").on("click", function(e){
			e.preventDefault();
			data = $(this).data();
			redirectPage = "";
			switch(data.codNivel) {
				case 1:
					redirectPage = "login_adm.asp"
					break;
				default:
					redirectPage = "login_fiscal.asp"
					break;
			}
			window.frames['I1'].location.href = redirectPage
		});

		$("select#profile_id").on("change", function(e){
			e.preventDefault();
			$("form#form_change_profile").submit();
		});
	});
</script>
</head>
<body onLoad="FP_preloadImgs(/*url*/'button8D.jpg', /*url*/'button8E.jpg', /*url*/'button9A.jpg', /*url*/'button9B.jpg', /*url*/'button7.jpg', /*url*/'button8.jpg', /*url*/'button10.jpg', /*url*/'button11.jpg', /*url*/'button129.jpg', /*url*/'button130.jpg')" style="text-align: center">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" height="543">
		<tr>
			<td valign="top">
				<table cellpadding="0" cellspacing="0" width="100%" height="98">
					<tr>
						<td width="925">
							<table cellpadding="0" cellspacing="0" border="0" width="100%" height="88">
								<tr>
									<td valign="top" height="88" width="100%">
										<table border="1" width="100%" id="table1" style="border-width: 0px">
											<tr>
												<td style="border-style: none; border-width: medium" width="8">&nbsp;</td>
												<td width="100%" style="border-style: none; border-width: medium">
													<p style="margin-top: 0; margin-bottom: 0">
														<a href="<%= MM_Logout %>"></a>
													</p>
													<table width="100%" border="1" style="margin-bottom: 0">
	          											<tr>
	            											<td width="95" style="text-align:center">
	            												<img border="0" src="consorcio.jpg" style="max-width:50px;">
	        												</td>
	            											<td width="319" style="text-align:center">
	            												<img src="cabecalho.png" style="max-width:320px;"/>
	        												</td>
	        												<td width="356" align="center">
	        													<span style="margin-top: 0; margin-bottom: 0; text-align: center;">
	        														<span class="style2"><%=(Recordset1.Fields.Item("mensagem").Value)%></span>
	    														</span>
															</td>
	            											<td width="100">
	            												<span style="margin-top: 0; margin-bottom: 0">
	            													<font face="Bauhaus 93" size="4">
	            														<a href="<%= MM_Logout %>">
	            															<img src="button12.jpg" alt="Logout" name="img6" width="100" height="20" border="0" align="right" id="img6" onMouseDown="FP_swapImg(1,0,/*id*/'img6',/*url*/'button11.jpg')" onMouseUp="FP_swapImg(0,0,/*id*/'img6',/*url*/'button10.jpg')" onMouseOver="FP_swapImg(1,0,/*id*/'img6',/*url*/'button10.jpg')" onMouseOut="FP_swapImg(0,0,/*id*/'img6',/*url*/'button12.jpg')" fp-style="fp-btn: Braided Row 5" fp-title="Logout">
	        															</a>
	    															</font>
																</span>
															</td>
														</tr>
	        										</table>
												</td>
											</tr>
										</table>
										<button type="button" id="btn_area" data-cod-nivel="<%=(Session("MM_UserAuthorization"))%>">
											<%
												Select Case Session("MM_UserAuthorization")
													Case 1
														Response.Write("ADMINISTRADOR")
													Case 2
														Response.Write("PROJETOS")
													Case 3
														Response.Write("INFORMAÇÕES GERENCIAIS")
													Case 4
														Response.Write("ENG. PLANEJ. OBRAS")
													Case 5
														Response.Write("ENG. OBRAS / FISCAL")
													Case 6
														Response.Write("MEDIÇÕES")
													Case 7
														Response.Write("MEIO AMBIENTE")
												End Select
											%>
										</button>
										<%
											If Session("MM_UserAuthorization_Admin") Then
										%>
										<span>
											<form id="form_change_profile" class="form" role="form" style="display: inline;" method="POST" action="?cp=yes">
												<strong>Ver Como:</strong>
												<select id="profile_id" name="profile_id">
													<option value="1" <% If Session("MM_UserAuthorization") = 1 Then Response.Write "selected='selected'" End If %>>ADMINISTRADOR</option>
													<option value="2" <% If Session("MM_UserAuthorization") = 2 Then Response.Write "selected='selected'" End If %>>PROJETOS</option>
													<option value="3" <% If Session("MM_UserAuthorization") = 3 Then Response.Write "selected='selected'" End If %>>INFORMAÇÕES GERENCIAIS</option>
													<option value="4" <% If Session("MM_UserAuthorization") = 4 Then Response.Write "selected='selected'" End If %>>ENG. PLANEJ. OBRAS</option>
													<option value="5" <% If Session("MM_UserAuthorization") = 5 Then Response.Write "selected='selected'" End If %>>ENG. OBRAS / FISCAL</option>
													<option value="6" <% If Session("MM_UserAuthorization") = 6 Then Response.Write "selected='selected'" End If %>>MEDIÇÕES</option>
													<option value="7" <% If Session("MM_UserAuthorization") = 7 Then Response.Write "selected='selected'" End If %>>MEIO AMBIENTE</option>
													<option value="8" <% If Session("MM_UserAuthorization") = 8 Then Response.Write "selected='selected'" End If %>>DAEE</option>
													<option value="9" <% If Session("MM_UserAuthorization") = 9 Then Response.Write "selected='selected'" End If %>>SUPERINTENDENTE</option>
												</select>
											</form>
										</span>
										<%
											End If
										%>
									</td>
								</tr>
							</table>
						</td>
						<td height="88" width="10">
							<font face="Bauhaus 93" size="6">
								<img alt="" width="10" height="88" src="MsoPnl_sh_b_14.jpg">
							</font>
						</td>
					</tr>
				</table>
				<%
					If Session("MM_UserAuthorization") = 1 Then
				%>
				<iframe name="I1" src="login_adm.asp" onload="this.width=screen.width-40;this.height=screen.height-230;" border="0" frameborder="0" style="border-style: outset; border-width: 3px; background-color: #E2E2E2">
					Seu navegador não oferece suporte para quadros embutidos ou está configurado para não exibi-los.
				</iframe>
				<%
					Else
				%>
				<iframe name="I1" src="login_fiscal.asp" onload="this.width=screen.width-40;this.height=screen.height-230;" border="0" frameborder="0" style="border-style: outset; border-width: 3px; background-color: #E2E2E2">
					Seu navegador não oferece suporte para quadros embutidos ou está configurado para não exibi-los.
				</iframe>
				<%
					End If
				%>
			</td>
		</tr>
		<tr>
			<td valign="top" width="799">&nbsp;</td>
		</tr>
	</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>