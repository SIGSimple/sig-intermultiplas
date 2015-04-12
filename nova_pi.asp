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

If (CStr(Request("MM_insert")) = "form2") Then

  MM_editConnection = MM_cpf_STRING
  MM_editTable = "tb_contrato"
  MM_editRedirectUrl = ""
  MM_fieldsStr  = "cod_predio|value|area_construida|value|area_gerenciadora|value|cod_Construtora|value|cod_reg|value|crit_calculo|value|crit_reajuste|value|desc_interv_gerenciadora|value|Desc_intervencao_FDE|value|vl_contrato|value|dig_contrato|value|dt_abertura|value|dt_assinatura|value|dt_base|value|dt_CI|value|dt_impressao_ois|value|dt_termino|value|dt_TRD|value|dt_TRP|value|fator_reducao|value|fical|value|gerenc_mede|value|Obra_pi|value|orc_FDE|value|orgao|value|PI|value|prz_contrato|value|reducao|value|solic_aditamento|value"
  MM_columnsStr = "cod_predio|',none,''|area_construida|none,none,NULL|area_gerenciadora|none,none,NULL|cod_Construtora|',none,''|cod_reg|none,none,NULL|crit_calculo|',none,''|crit_reajuste|',none,''|desc_interv_gerenciadora|',none,''|Desc_intervencao_FDE|',none,''|vl_contrato|none,none,NULL|dig_contrato|none,none,NULL|dt_abertura|',none,NULL|dt_assinatura|',none,NULL|dt_base|',none,NULL|dt_CI|',none,NULL|dt_impressao_ois|',none,NULL|dt_termino|',none,NULL|dt_TRD|',none,NULL|dt_TRP|',none,NULL|fator_reducao|none,none,NULL|fical|',none,''|gerenc_mede|none,1,0|Obra_pi|none,1,0|orc_FDE|none,none,NULL|orgao|',none,''|PI|none,none,NULL|prz_contrato|none,none,NULL|reducao|none,none,NULL|solic_aditamento|none,1,0"

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
Recordset1.Source = "SELECT cod_predio, Nome_Unidade FROM tb_Predios ORDER BY Nome_Unidade ASC"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim Recordset2
Dim Recordset2_numRows

Set Recordset2 = Server.CreateObject("ADODB.Recordset")
Recordset2.ActiveConnection = MM_cpf_STRING
Recordset2.Source = "SELECT cod_construtora, Construtora FROM tb_Construtora ORDER BY Construtora ASC"
Recordset2.CursorType = 0
Recordset2.CursorLocation = 2
Recordset2.LockType = 1
Recordset2.Open()

Recordset2_numRows = 0
%>
<%
Dim Recordset3
Dim Recordset3_numRows

Set Recordset3 = Server.CreateObject("ADODB.Recordset")
Recordset3.ActiveConnection = MM_cpf_STRING
Recordset3.Source = "SELECT * FROM tb_responsavel ORDER BY Responsável ASC"
Recordset3.CursorType = 0
Recordset3.CursorLocation = 2
Recordset3.LockType = 1
Recordset3.Open()

Recordset3_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style11 {font-family: Arial, Helvetica, sans-serif; font-size: 9; }
.style12 {font-size: 9}
-->
</style>
</head>

<body>
<form id="form1" name="form1" method="post" action="">
</form>


<form method="post" action="<%=MM_editAction%>" name="form2">
  <table align="center">
    <tr>
      <td nowrap align="right" bgcolor="#336699" colspan="2">
  <p align="center"><font face="Bauhaus 93" size="5" color="#FFFFFF">CADASTRO DE 
	PIs</font></p>
  	</td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right" bgcolor="#EEEEEE">&nbsp;</td>
      <td bgcolor="#EEEEEE">&nbsp;</td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right" bgcolor="#EEEEEE"><span class="style11">
		<font size="2">Predio</font><font size="2">:</font></span></td>
      <td bgcolor="#EEEEEE"><select name="cod_predio">
        <%
While (NOT Recordset1.EOF)
%>
        <option value="<%=(Recordset1.Fields.Item("cod_predio").Value)%>"><%=(Recordset1.Fields.Item("Nome_Unidade").Value)%></option>
<%
  Recordset1.MoveNext()
Wend
If (Recordset1.CursorType > 0) Then
  Recordset1.MoveFirst
Else
  Recordset1.Requery
End If
%>
      </select>      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right" bgcolor="#EEEEEE"><span class="style11">
		<font size="2">PI:</font></span></td>
      <td bgcolor="#EEEEEE"><input type="text" name="PI" value="" size="32" />      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right" bgcolor="#EEEEEE"><span class="style11">
		<font size="2">Cod_Construtora</font><font size="2">:</font></span></td>
      <td bgcolor="#EEEEEE"><select name="cod_Construtora">
        <%
While (NOT Recordset2.EOF)
%>
        <option value="<%=(Recordset2.Fields.Item("cod_construtora").Value)%>"><%=(Recordset2.Fields.Item("Construtora").Value)%></option>
<%
  Recordset2.MoveNext()
Wend
If (Recordset2.CursorType > 0) Then
  Recordset2.MoveFirst
Else
  Recordset2.Requery
End If
%>
      </select>      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right" bgcolor="#EEEEEE"><span class="style11">
		<font size="2">Fiscal:</font></span></td>
      <td bgcolor="#EEEEEE"><select name="fical">
        <%
While (NOT Recordset3.EOF)
%><option value="<%=(Recordset3.Fields.Item("cod_resp").Value)%>"><%=(Recordset3.Fields.Item("Responsável").Value)%></option>
          <%
  Recordset3.MoveNext()
Wend
If (Recordset3.CursorType > 0) Then
  Recordset3.MoveFirst
Else
  Recordset3.Requery
End If
%>
      </select>      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right" bgcolor="#EEEEEE"><span class="style11">
		<font size="2">Gerenc_mede</font><font size="2">:</font></span></td>
      <td bgcolor="#EEEEEE"><input type="checkbox" name="gerenc_mede" value=1 >      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right" bgcolor="#EEEEEE"><span class="style11">
		<font size="2">Obra_pi</font><font size="2">:</font></span></td>
      <td bgcolor="#EEEEEE"><input type="checkbox" name="Obra_pi" value=1 >      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right" bgcolor="#EEEEEE"><span class="style11">
		<font size="2">Orc_FDE</font><font size="2">:</font></span></td>
      <td bgcolor="#EEEEEE"><input type="text" name="orc_FDE" value="" size="32">      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right" bgcolor="#EEEEEE"><span class="style11">
		<font size="2">Orgao</font><font size="2">:</font></span></td>
      <td bgcolor="#EEEEEE"><input type="text" name="orgao" value="" size="32">      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right" bgcolor="#EEEEEE"><span class="style11">
		<font size="2">Prz_contrato</font><font size="2">:</font></span></td>
      <td bgcolor="#EEEEEE"><input type="text" name="prz_contrato" value="" size="32">      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right" bgcolor="#EEEEEE"><span class="style11">
		<font size="2">Reducao</font><font size="2">:</font></span></td>
      <td bgcolor="#EEEEEE"><input type="text" name="reducao" value="" size="32">      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right" bgcolor="#EEEEEE"><span class="style11">
		<font size="2">Solic_aditamento</font><font size="2">:</font></span></td>
      <td bgcolor="#EEEEEE"><input type="checkbox" name="solic_aditamento" value=1 >      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right" bgcolor="#EEEEEE"><span class="style12"></span></td>
      <td bgcolor="#EEEEEE"><input type="submit" value="Salvar">      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form2">
</form>
<p>&nbsp;</p>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>
<%
Recordset3.Close()
Set Recordset3 = Nothing
%>