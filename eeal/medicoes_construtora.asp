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
  MM_editTable = "tb_Medicao_Construtora"
  MM_editRedirectUrl = "medicoes_construtora.asp"
  MM_fieldsStr  = "PI|value|Data_da_Medio_na_Obra|value|Data_de_Chegada_no_Escritrio|value|Data_de_Envio_para_FDE|value|Data_de_Execuo|value|Data_Prevista_para_o_Trmino|value|_Medio_Final_|value|Informao_da_Placa_de_Obra|value|Nmero_da_Medio|value|Porcentagem_de_Avano|value|Valor_da_Medio|value|Valor_do_Contrato|value"
  MM_columnsStr = "PI|none,none,NULL|[Data da Medição na Obra]|',none,''|[Data de Chegada no Escritório]|',none,''|[Data de Envio para FDE]|',none,''|[Data de Execução]|',none,''|[Data Prevista para o Término]|',none,''|[É Medição Final ?]|none,1,0|[Informação da Placa de Obra]|',none,''|[Número da Medição]|none,none,NULL|[Porcentagem de Avanço]|',none,''|[Valor da Medição]|',none,''|[Valor do Contrato]|',none,''"

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
Dim rs_lista_med_constr
Dim rs_lista_med_constr_numRows

Set rs_lista_med_constr = Server.CreateObject("ADODB.Recordset")
rs_lista_med_constr.ActiveConnection = MM_cpf_STRING
rs_lista_med_constr.Source = "SELECT tb_Medicao_Construtora.*  FROM tb_Medicao_Construtora  ORDER BY tb_Medicao_Construtora.PI, tb_Medicao_Construtora.[Número da Medição] DESC;  "
rs_lista_med_constr.CursorType = 0
rs_lista_med_constr.CursorLocation = 2
rs_lista_med_constr.LockType = 1
rs_lista_med_constr.Open()

rs_lista_med_constr_numRows = 0
%>
<%
Dim rs_lista_pi
Dim rs_lista_pi_numRows

Set rs_lista_pi = Server.CreateObject("ADODB.Recordset")
rs_lista_pi.ActiveConnection = MM_cpf_STRING
rs_lista_pi.Source = "SELECT tb_pi.Código, tb_pi.PI  FROM tb_pi  ORDER BY tb_pi.PI;  "
rs_lista_pi.CursorType = 0
rs_lista_pi.CursorLocation = 2
rs_lista_pi.LockType = 1
rs_lista_pi.Open()

rs_lista_pi_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 15
Repeat1__index = 0
rs_lista_med_constr_numRows = rs_lista_med_constr_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>medi&ccedil;&otilde;es Construtoras</title>
<style type="text/css">
<!--
.style5 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
.style11 {font-family: Arial, Helvetica, sans-serif; font-size: 11px; font-weight: bold; }
.style36 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
	font-weight: bold;
}
-->
</style>
</head>

<body>
  <p align="center" class="style36">CADASTRO DE MEDI&Ccedil;&Otilde;ES CONSTRUTORAS </p>
  <form method="POST" action="<%=MM_editAction%>" name="form1">
    <table align="center">
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC"><span class="style5">PI:</span></td>
        <td bgcolor="#CCCCCC"><select name="PI">
            <%
While (NOT rs_lista_pi.EOF)
%>
            <option value="<%=(rs_lista_pi.Fields.Item("PI").Value)%>"><%=(rs_lista_pi.Fields.Item("PI").Value)%></option>
            <%
  rs_lista_pi.MoveNext()
Wend
If (rs_lista_pi.CursorType > 0) Then
  rs_lista_pi.MoveFirst
Else
  rs_lista_pi.Requery
End If
%>
          </select>        </td>
        <td bgcolor="#CCCCCC"><span class="style5">É Medição Final ?:</span></td>
        <td bgcolor="#CCCCCC"><input <%If (CStr((rs_lista_med_constr.Fields.Item("É Medição Final ?").Value)) = CStr("True")) Then Response.Write("checked=""checked""") : Response.Write("")%> type="checkbox" name="_Medio_Final_" value=1 /></td>
      </tr>
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC"><span class="style5">Data da Medição na Obra:</span></td>
        <td bgcolor="#CCCCCC"><input type="text" name="Data_da_Medio_na_Obra" value="" size="15">        </td>
        <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><span class="style5">Informa&ccedil;&atilde;o da Placa de Obra:</span></td>
        <td bgcolor="#CCCCCC"><input type="text" name="Informao_da_Placa_de_Obra" value="" size="32" />        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC"><span class="style5">Data de Chegada no Escritório:</span></td>
        <td bgcolor="#CCCCCC"><input type="text" name="Data_de_Chegada_no_Escritrio" value="" size="15">        </td>
        <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><span class="style5">N&uacute;mero da Medi&ccedil;&atilde;o:</span></td>
        <td bgcolor="#CCCCCC"><input type="text" name="Nmero_da_Medio" value="" size="5" />        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC"><span class="style5">Data de Envio para FDE:</span></td>
        <td bgcolor="#CCCCCC"><input type="text" name="Data_de_Envio_para_FDE" value="" size="15">        </td>
        <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><span class="style5">Porcentagem de Avan&ccedil;o:</span></td>
        <td bgcolor="#CCCCCC"><input type="text" name="Porcentagem_de_Avano" value="" size="10" />        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC"><span class="style5">Data de Execução:</span></td>
        <td bgcolor="#CCCCCC"><input type="text" name="Data_de_Execuo" value="" size="15">        </td>
        <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><span class="style5">Valor da Medi&ccedil;&atilde;o:</span></td>
        <td bgcolor="#CCCCCC"><input type="text" name="Valor_da_Medio" value="" size="15" />        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC"><span class="style5">Data Prevista para o Término:</span></td>
        <td bgcolor="#CCCCCC"><input type="text" name="Data_Prevista_para_o_Trmino" value="" size="15">        </td>
        <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><span class="style5">Valor do Contrato:</span></td>
        <td bgcolor="#CCCCCC"><input type="text" name="Valor_do_Contrato" value="" size="15" />        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC">&nbsp;</td>
        <td bgcolor="#CCCCCC">&nbsp;</td>
        <td bgcolor="#CCCCCC">&nbsp;</td>
        <td bgcolor="#CCCCCC">&nbsp;</td>
      </tr>
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC">&nbsp;</td>
        <td bgcolor="#CCCCCC"><input type="submit" value="Salvar">        </td>
        <td bgcolor="#CCCCCC">&nbsp;</td>
        <td bgcolor="#CCCCCC">&nbsp;</td>
      </tr>
    </table>
    <input type="hidden" name="MM_insert" value="form1">
</form>
  <table border="0">
  <tr bgcolor="#999999">
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><span class="style11">c&oacute;digo</span></td>
    <td><span class="style11">PI</span></td>
    <td><span class="style11">N&uacute;mero da Medi&ccedil;&atilde;o</span></td>
    <td><span class="style11">Data da Medi&ccedil;&atilde;o na Obra</span></td>
    <td><span class="style11">Data de Execu&ccedil;&atilde;o</span></td>
    <td><span class="style11">Valor da Medi&ccedil;&atilde;o</span></td>
    <td><span class="style11">Porcentagem de Avan&ccedil;o</span></td>
    <td><span class="style11">Data Prevista para o T&eacute;rmino</span></td>
    <td><span class="style11">&Eacute; Medi&ccedil;&atilde;o Final ?</span></td>
    <td><span class="style11">Valor do Contrato</span></td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT rs_lista_med_constr.EOF)) %>
    <tr bgcolor="#CCCCCC">
      <td><a href="exc_med_constr.asp?cod_med_constr=<%=(rs_lista_med_constr.Fields.Item("cod_med_constr").Value)%>"><img src="imagens/delete.gif" width="16" height="15" border="0" /></a></td>
      <td><a href="altera_med_constr.asp?cod_med_constr=<%=(rs_lista_med_constr.Fields.Item("cod_med_constr").Value)%>"><img src="depto/imagens/edit.gif" width="16" height="15" border="0" /></a></td>
      <td><span class="style5"><%=(rs_lista_med_constr.Fields.Item("cod_med_constr").Value)%></span></td>
      <td><span class="style5"><%=(rs_lista_med_constr.Fields.Item("PI").Value)%></span></td>
      <td><div align="center"><span class="style5"><%=(rs_lista_med_constr.Fields.Item("Número da Medição").Value)%></span></div></td>
      <td><div align="right"><span class="style5"><%=(rs_lista_med_constr.Fields.Item("Data da Medição na Obra").Value)%></span></div></td>
      <td><div align="right"><span class="style5"><%=(rs_lista_med_constr.Fields.Item("Data de Execução").Value)%></span></div></td>
      <td><div align="right"><span class="style5"><%= FormatNumber((rs_lista_med_constr.Fields.Item("Valor da Medição").Value), 2, -2, -2, -2) %></span></div></td>
      <td><div align="right"><span class="style5"><%= FormatPercent((rs_lista_med_constr.Fields.Item("Porcentagem de Avanço").Value), 2, -2, -2, -2) %></span></div></td>
      <td><div align="right"><span class="style5"><%=(rs_lista_med_constr.Fields.Item("Data Prevista para o Término").Value)%></span></div></td>
      <td><div align="center"><span class="style5"><%=(rs_lista_med_constr.Fields.Item("É Medição Final ?").Value)%></span></div></td>
      <td><div align="right"><span class="style5"><%= FormatNumber((rs_lista_med_constr.Fields.Item("Valor do Contrato").Value), 2, -2, -2, -2) %></span></div></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_lista_med_constr.MoveNext()
Wend
%>
</table>
</body>
</html>
<%
rs_lista_med_constr.Close()
Set rs_lista_med_constr = Nothing
%>
<%
rs_lista_pi.Close()
Set rs_lista_pi = Nothing
%>
