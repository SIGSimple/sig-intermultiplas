<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername=CStr(Request.Form("usuario"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization="nivel"
  MM_redirectLoginSuccess="inicio.asp"
  MM_redirectLoginFailed="erro.asp"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_cpf_STRING
  MM_rsUser.Source = "SELECT nome, senha"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM login WHERE nome='" & Replace(MM_valUsername,"'","''") &"' AND senha='" & Replace(Request.Form("senha"),"'","''") & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And true Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>CONSÓRCIO INTERMULTIPLAS</title>
<style type="text/css">
<!--
.style1 {
	font-family: Arial, Helvetica, sans-serif;
	color: #993300;
	font-weight: bold;
}
.style3 {
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
	color: #990000;
}
-->
</style>
<!--mstheme--><link rel="stylesheet" href="netw1011-28591.css">
<meta name="Microsoft Theme" content="network 1011">
<style type="text/css">
<!--
body {
	background-image: url(imagens/novo/mapa.png);
	background-repeat: no-repeat;
	background-position:inherit;
}
-->
</style></head>

<body>
<form ACTION="<%=MM_LoginAction%>" id="form1" name="form1" method="POST">
  <table border="1" width="100%" id="table1" style="border-width: 0px">
	<tr>
		<td style="border-style: none; border-width: medium" width="190">
			<img border="0" src="imagens/novo/fde.jpg" width="120" height="49"></td>
		<td style="border-style: none; border-width: medium" width="273">&nbsp;
		</td>
		<td style="border-style: none; border-width: medium" width="273">&nbsp;
		</td>
		<td style="border-style: none; border-width: medium" width="272">
		<img border="0" src="imagens/logo_Arcadis.jpg" width="674" height="52"></td>
		<td style="border-style: none; border-width: medium">&nbsp;</td>
	</tr>
	<tr>
		<td style="border-style: none; border-width: medium" width="1008" colspan="4">
		<p align="center"><font size="6">&nbsp;</font></p>
	  </td>
	  <td style="border-style: none; border-width: medium">&nbsp;
	  </td>
	</tr>
  </table>
  <div align="center">
    &nbsp;<p>&nbsp;</p>
    <p>&nbsp;</p>
    <p>&nbsp;</p>
    <p>&nbsp;</p>
    <p>&nbsp;</p>
    <table width="596" border="1" align="right" style="border-width: 0px">
      <tr>
        <td width="394" style="border-style: none; border-width: medium"><label>
          <div align="right"><span class="UserGenericHeader">Login</span><font face="Arial, Helvetica, sans-serif">
			</font> </div>
          </label></td>
        <td width="120" style="border-style: none; border-width: medium">
		  <div align="right"><font face="Arial, Helvetica, sans-serif">
	        <input name="usuario" type="text" id="usuario" size="20" />
	    </font></div></td>
        <td width="60" rowspan="3" style="border-style: none; border-width: medium"><img src="imagens/cadeado.jpeg" width="50" height="50" /></td>
      </tr>
      <tr>
        <td style="border-style: none; border-width: medium"><div align="right"><span class="UserGenericHeader">Senha</span></div></td>
        <td style="border-style: none; border-width: medium"><div align="right"><font face="Arial, Helvetica, sans-serif">
          <input name="senha" type="password" id="senha" size="20" />
        </font></div></td>
      </tr>
      <tr>
        <td style="border-style: none; border-width: medium">&nbsp;</td>
        <td style="border-style: none; border-width: medium"><font color="#FF0000">
          <input type="submit" name="Submit" value="Acessar" style="color: #FF0000; font-family: Arial; font-weight: bold" />
        </font></td>
      </tr>
    </table>
     
<p><label></label>
  </div>
  <p align="center">&nbsp;</p>
</form>
</body>
</html>
