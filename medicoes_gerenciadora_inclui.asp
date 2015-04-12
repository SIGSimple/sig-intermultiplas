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
  MM_editTable = "tb_Medicao_Gerenciadora"
  MM_editRedirectUrl = "medicoes_gerenciadora_inclui.asp"
  MM_fieldsStr  = "campo_de_trabalho|value|cod_medicao_gerenc|value|Cdigo_do_Prdio|value|Data_de_incio|value|Data_de_Trmino|value|Data_de_Trmino_Contratual|value|Durao_da_OS|value|Gerenciadora_Mede_|value|Justificativa|value|Nome_da_Unidade|value|Nmero_da_OS|value|PI|value"
  MM_columnsStr = "[campo de trabalho]|none,none,NULL|cod_medicao_gerenc|none,none,NULL|[Código do Prédio]|none,none,NULL|[Data de início]|',none,''|[Data de Término]|',none,''|[Data de Término Contratual]|none,none,NULL|[Duração da OS]|none,none,NULL|[Gerenciadora Mede ?]|none,1,0|Justificativa|',none,''|[Nome da Unidade]|',none,''|[Número da OS]|',none,''|PI|none,none,NULL"

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
Dim medicoes__MMColParam
medicoes__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  medicoes__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim medicoes
Dim medicoes_numRows

Set medicoes = Server.CreateObject("ADODB.Recordset")
medicoes.ActiveConnection = MM_cpf_STRING
medicoes.Source = "SELECT * FROM tb_Medicao_Gerenciadora WHERE PI = " + Replace(medicoes__MMColParam, "'", "''") + ""
medicoes.CursorType = 0
medicoes.CursorLocation = 2
medicoes.LockType = 1
medicoes.Open()

medicoes_numRows = 0
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
      <td nowrap align="right">Campo de trabalho:</td>
      <td><input type="text" name="campo_de_trabalho" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Cod_medicao_gerenc:</td>
      <td><input type="text" name="cod_medicao_gerenc" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Código do Prédio:</td>
      <td><input type="text" name="Cdigo_do_Prdio" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Data de início:</td>
      <td><input type="text" name="Data_de_incio" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Data de Término:</td>
      <td><input type="text" name="Data_de_Trmino" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Data de Término Contratual:</td>
      <td><input type="text" name="Data_de_Trmino_Contratual" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Duração da OS:</td>
      <td><input type="text" name="Durao_da_OS" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Gerenciadora Mede ?:</td>
      <td><input type="checkbox" name="Gerenciadora_Mede_" value=1 >
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Justificativa:</td>
      <td><input type="text" name="Justificativa" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Nome da Unidade:</td>
      <td><input type="text" name="Nome_da_Unidade" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Número da OS:</td>
      <td><input type="text" name="Nmero_da_OS" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">PI:</td>
      <td><input type="text" name="PI" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">&nbsp;</td>
      <td><input type="submit" value="Insert record">
      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1">
</form>
<p>&nbsp;</p>
</body>
</html>
<%
medicoes.Close()
Set medicoes = Nothing
%>
