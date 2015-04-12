<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_cpf_STRING
  MM_editTable = "tb_pi"
  MM_editColumn = "PI"
  MM_recordId = "'" + Request.Form("MM_recordId") + "'"
  MM_editRedirectUrl = "msg_exclusao.asp"

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
  
End If
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_cpf_STRING
Recordset1.Source = "SELECT tb_PI.Pi, tb_responsavel.Respons�vel, tb_predio.Munic�pio, tb_Construtora.Construtora, tb_tipo_empreendimento.desc_tipo AS tipo_empreendimento, [tb_depto].[sigla]+' - '+[tb_depto].[desc_depto] AS programa, tb_PI.nome_empreendimento FROM tb_depto RIGHT JOIN (tb_tipo_empreendimento RIGHT JOIN (tb_responsavel RIGHT JOIN (tb_predio RIGHT JOIN (tb_Construtora RIGHT JOIN tb_PI ON tb_Construtora.cod_construtora = tb_PI.cod_construtora) ON tb_predio.cod_predio = tb_PI.cod_predio) ON tb_responsavel.cod_fiscal = tb_PI.cod_fiscal) ON tb_tipo_empreendimento.id = tb_PI.cod_tipo_empreendimento) ON tb_depto.cod_depto = tb_PI.cod_programa WHERE (((tb_PI.[PI])='"& Replace(Recordset1__MMColParam, "'", "''") &"')) ORDER BY tb_PI.PI;    "
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style5 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; }
.style7 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; color: #FFFFFF; }
-->
</style>
</head>

<body>
<form id="form1" name="form1" method="POST" action="<%=MM_editAction%>">
  <div align="center">
    <input type="hidden" name="MM_delete" value="form1">
    <input type="hidden" name="MM_recordId" value="<%=(Request.QueryString("PI"))%>">
  </div>
  <label>
  <div align="center">
    <input type="submit" name="Submit" value="Excluir" />
  </div>
  </label>
</form>
<div align="center">
  <table border="1">
    <tr bgcolor="#666666">
      <td><span class="style7">Autos</span></td>
      <td><span class="style7">Nome do Empreendimento</span></td>
      <td><span class="style7">Tipo</span></td>
      <td><span class="style7">Programa</span></td>
      <td><span class="style7">Munic&iacute;pio</span></td>
      <td><span class="style7">Eng. Obras Cons�rcio</span></td>
      <td><span class="style7">Empresa</span></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) %>
      <tr bgcolor="#CCCCCC">
        <td><span class="style5"><%=(Recordset1.Fields.Item("PI").Value)%></span></td>
        <td><span class="style5"><%=(Recordset1.Fields.Item("nome_empreendimento").Value)%></span></td>
        <td><span class="style5"><%=(Recordset1.Fields.Item("tipo_empreendimento").Value)%></span></td>
        <td><span class="style5"><%=(Recordset1.Fields.Item("programa").Value)%></span></td>
        <td><div align="left"><span class="style5"><%=(Recordset1.Fields.Item("Munic�pio").Value)%></span></div></td>
        <td><div align="left"><span class="style5"><%=(Recordset1.Fields.Item("Respons�vel").Value)%></span></div></td>
        <td><div align="left"><span class="style5"><%=(Recordset1.Fields.Item("Construtora").Value)%></span></div></td>
      </tr>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
  </table>
</div>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
