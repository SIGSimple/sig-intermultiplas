<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim rs_login
Dim rs_login_numRows

Set rs_login = Server.CreateObject("ADODB.Recordset")
rs_login.ActiveConnection = MM_cpf_STRING
rs_login.Source = "SELECT * FROM login"
rs_login.CursorType = 0
rs_login.CursorLocation = 2
rs_login.LockType = 1
rs_login.Open()

rs_login_numRows = 0
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_cpf_STRING
Recordset1.Source = "SELECT tb_responsavel.Responsável  FROM tb_responsavel  GROUP BY tb_responsavel.Responsável  ORDER BY tb_responsavel.Responsável;  "
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername=CStr(Request.Form("usuario"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization=""
  MM_redirectLoginSuccess="filtro_acomp_.asp"
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
<title>Untitled Document</title>
<style type="text/css">
<!--
.style3 {font-family: Arial, Helvetica, sans-serif; font-size: 14px; }
-->
</style>
</head>

<body>
<form id="form1" name="form1" method="post" action="filtro_acomp_.asp?nome=<%=(rs_login.Fields.Item("nome").Value)%>&amp;senha=<%=(rs_login.Fields.Item("senha").Value)%>">
  <label><span class="style3">Selecione o Fiscal para a pesquisa
  </span></label>
  <label for="select"></label>
  <select name="nome" id="nome">
    <%
While (NOT Recordset1.EOF)
%>
    <option value="<%=(Recordset1.Fields.Item("Responsável").Value)%>"><%=(Recordset1.Fields.Item("Responsável").Value)%></option>
    <%
  Recordset1.MoveNext()
Wend
If (Recordset1.CursorType > 0) Then
  Recordset1.MoveFirst
Else
  Recordset1.Requery
End If
%>
  </select>
  <label><span class="style3">  </span></label>
  <span class="style3">
  <label>
  <input type="submit" name="Submit" value="buscar" />
  </label>
  </span><span class="style3">  </span>
</form>
</body>
</html>
<%
rs_login.Close()
Set rs_login = Nothing
%>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
