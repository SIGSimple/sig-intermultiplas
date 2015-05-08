<%@LANGUAGE="VBSCRIPT"%>
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="2,3,4,5,6,7"
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

<body onLoad="FP_preloadImgs(/*url*/'button114.jpg', /*url*/'button115.jpg', /*url*/'button116.jpg', /*url*/'button117.jpg', /*url*/'button6E.jpg', /*url*/'button6F.jpg', /*url*/'button118.jpg', /*url*/'button119.jpg', /*url*/'button174.jpg', /*url*/'button175.jpg', /*url*/'button180.jpg', /*url*/'button181.jpg', /*url*/'button205.jpg', /*url*/'button206.jpg', /*url*/'button1A3.jpg', /*url*/'button1B3.jpg')">

<div align="center">
  <div align="center">
    <font face="Arial">
      <strong>
        <span class="style3">
          <%
            If Session("MM_UserAuthorization") = 2 Then
          %>
  	      <font size="5">&Aacute;REA DE PROJETOS</font>
          <%
            ElseIf Session("MM_UserAuthorization") = 3 Then
          %>
          <font size="5">&Aacute;REA DE INFORMAÇÕES GERENCIAIS</font>
          <%
            ElseIf Session("MM_UserAuthorization") = 4 Then
          %>
          <font size="5">&Aacute;REA DE ENG. PLANEJ. OBRAS</font>
          <%
            ElseIf Session("MM_UserAuthorization") = 5 Then
          %>
          <font size="5">&Aacute;REA DE ENG. OBRAS / FISCAL</font>
          <%
            ElseIf Session("MM_UserAuthorization") = 6 Then
          %>
          <font size="5">&Aacute;REA DE MEDIÇÕES</font>
          <%
            ElseIf Session("MM_UserAuthorization") = 7 Then
          %>
          <font size="5">&Aacute;REA DE MEIO AMBIENTE</font>
          <%
            End If
          %>
       </span>
     </strong>
   </font>
  </div>
	<table x:str border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; width: 638px" id="table1">
    <colgroup>
      <col width="64" style="width: 48pt"><col width="64" style="width: 48pt">
      <col width="64" style="width: 48pt">
    </colgroup>
    <tr height="18" style="height: 13.5pt">
      <td height="18" colspan="3" style="height: 13.5pt; width: 237px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px"></td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" colspan="3" style="height: 13.5pt; width: 237px; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px">		</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
        <%
          If Session("MM_UserAuthorization") = 1 Or Session("MM_UserAuthorization") = 3 Or Session("MM_UserAuthorization") = 4 Then
        %>
        <a href="cadastro_pi.asp">Cadastro de Empreendimento</a>
        <%
          End If
        %>  
      </td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">
		  <!-- <a href="rel_pi_filtro_do_fiscal.asp">Busca por Fiscal</a> -->
    </td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">
		&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237"></td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">
		  <!-- <a href="busca_pi_unidade_Fiscal.asp">Busca por Nome</a> -->
    </td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">
        <%
          If Session("MM_UserAuthorization") = 6 Then
        %>
        <a href="cad_contrato.asp">Cadastro de Contratos</a>
        <%
          Else
        %>
        &nbsp;
        <%
          End If
        %>
      </td>
      <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">
      <!-- <a href="filtro_cod_predio_fiscal.asp">Busca por Num. Autos</a> -->
    </td>
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
        <a href="filtro_acomp.asp">Acompanhamento de Obra</a>
      </td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="18" style="height: 13.5pt">
      <td height="18" style="height: 13.5pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">
        </td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">
        <!-- <a href="busca_pi_cod_Fiscal.asp">Busca por Bacia</a> -->
      </td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="17" style="height: 12.75pt">
      <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">&nbsp;</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
    <tr height="17" style="height: 12.75pt">
      <td height="17" style="height: 12.75pt; color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="237">		</td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="26">
		  <a href="reports.asp">Relatórios</a>
    </td>
	    <td style="color: windowtext; font-size: 10.0pt; font-weight: 400; font-style: normal; text-decoration: none; font-family: Arial; text-align: center; vertical-align: bottom; white-space: nowrap; border: medium none; padding: 0px" width="185">&nbsp;</td>
    </tr>
  </table>
</div>
</body>

</html>
