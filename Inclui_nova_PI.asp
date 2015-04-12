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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_cpf_STRING
  MM_editTable = "tb_PI"
  MM_editRedirectUrl = "msg_inclusao.asp"
  MM_fieldsStr  = "PI|value|cod_construtora|value|cod_fiscal|value|cod_mun|value|cod_predio|value"
  MM_columnsStr = "PI|none,none,NULL|cod_construtora|none,none,NULL|cod_fiscal|none,none,NULL|cod_mun|none,none,NULL|cod_predio|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

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
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
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
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_cpf_STRING
Recordset1.Source = "SELECT * FROM tb_PI"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim rs_municipio
Dim rs_municipio_numRows

Set rs_municipio = Server.CreateObject("ADODB.Recordset")
rs_municipio.ActiveConnection = MM_cpf_STRING
rs_municipio.Source = "SELECT * FROM tb_Municipios ORDER BY Municipios ASC"
rs_municipio.CursorType = 0
rs_municipio.CursorLocation = 2
rs_municipio.LockType = 1
rs_municipio.Open()

rs_municipio_numRows = 0
%>
<%
Dim rs_fiscal
Dim rs_fiscal_numRows

Set rs_fiscal = Server.CreateObject("ADODB.Recordset")
rs_fiscal.ActiveConnection = MM_cpf_STRING
rs_fiscal.Source = "SELECT * FROM tb_responsavel ORDER BY Responsável ASC"
rs_fiscal.CursorType = 0
rs_fiscal.CursorLocation = 2
rs_fiscal.LockType = 1
rs_fiscal.Open()

rs_fiscal_numRows = 0
%>
<%
Dim rs_construtora
Dim rs_construtora_numRows

Set rs_construtora = Server.CreateObject("ADODB.Recordset")
rs_construtora.ActiveConnection = MM_cpf_STRING
rs_construtora.Source = "SELECT cod_construtora, Construtora FROM tb_Construtora ORDER BY Construtora ASC"
rs_construtora.CursorType = 0
rs_construtora.CursorLocation = 2
rs_construtora.LockType = 1
rs_construtora.Open()

rs_construtora_numRows = 0
%>
<%
Dim rs_predio
Dim rs_predio_numRows

Set rs_predio = Server.CreateObject("ADODB.Recordset")
rs_predio.ActiveConnection = MM_cpf_STRING
rs_predio.Source = "SELECT tb_predio.cod_predio, [tb_predio].[cod_predio] & ' - ' & [tb_predio].[Nome_Unidade] AS Expr1  FROM tb_PI INNER JOIN tb_predio ON tb_PI.cod_predio = tb_predio.cod_predio  GROUP BY tb_predio.cod_predio, [tb_predio].[cod_predio] & ' - ' & [tb_predio].[Nome_Unidade]  ORDER BY [tb_predio].[cod_predio] & ' - ' & [tb_predio].[Nome_Unidade];"
rs_predio.CursorType = 0
rs_predio.CursorLocation = 2
rs_predio.LockType = 1
rs_predio.Open()

rs_predio_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
</head>

<body>
<form method="post" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC">PI:</td>
      <td bgcolor="#CCCCCC"><input type="text" name="PI" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC">Construtora</td>
      <td bgcolor="#CCCCCC"><select name="cod_construtora">
        <%
While (NOT rs_construtora.EOF)
%><option value="<%=(rs_construtora.Fields.Item("cod_construtora").Value)%>"><%=(rs_construtora.Fields.Item("Construtora").Value)%></option>
          <%
  rs_construtora.MoveNext()
Wend
If (rs_construtora.CursorType > 0) Then
  rs_construtora.MoveFirst
Else
  rs_construtora.Requery
End If
%>
      </select>
      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC">Cod_fiscal:</td>
      <td bgcolor="#CCCCCC"><select name="cod_fiscal">
        <%
While (NOT rs_fiscal.EOF)
%><option value="<%=(rs_fiscal.Fields.Item("cod_fiscal").Value)%>"><%=(rs_fiscal.Fields.Item("Responsável").Value)%></option>
          <%
  rs_fiscal.MoveNext()
Wend
If (rs_fiscal.CursorType > 0) Then
  rs_fiscal.MoveFirst
Else
  rs_fiscal.Requery
End If
%>
      </select>
      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC">Cod_mun:</td>
      <td bgcolor="#CCCCCC"><select name="cod_mun">
        <%
While (NOT rs_municipio.EOF)
%><option value="<%=(rs_municipio.Fields.Item("cod_mun").Value)%>"><%=(rs_municipio.Fields.Item("Municipios").Value)%></option>
          <%
  rs_municipio.MoveNext()
Wend
If (rs_municipio.CursorType > 0) Then
  rs_municipio.MoveFirst
Else
  rs_municipio.Requery
End If
%>
      </select>
      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC">Cod_predio:</td>
      <td bgcolor="#CCCCCC"><select name="cod_predio">
        <%
While (NOT rs_predio.EOF)
%>
        <option value="<%=(rs_predio.Fields.Item("cod_predio").Value)%>"><%=(rs_predio.Fields.Item("Expr1").Value)%></option>
        <%
  rs_predio.MoveNext()
Wend
If (rs_predio.CursorType > 0) Then
  rs_predio.MoveFirst
Else
  rs_predio.Requery
End If
%>
      </select>
      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC">&nbsp;</td>
      <td bgcolor="#CCCCCC"><input type="submit" value="Salvar">
      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1">
</form>
<p>&nbsp;</p>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
rs_municipio.Close()
Set rs_municipio = Nothing
%>
<%
rs_fiscal.Close()
Set rs_fiscal = Nothing
%>
<%
rs_construtora.Close()
Set rs_construtora = Nothing
%>
<%
rs_predio.Close()
Set rs_predio = Nothing
%>
