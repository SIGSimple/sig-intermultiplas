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
  MM_editRedirectUrl = "medicoes_gerenciadora.asp"
  MM_fieldsStr  = "PI|value|Data_de_incio|value|Data_de_Trmino|value|Data_de_Trmino_Contratual|value|Durao_da_OS|value|Gerenciadora_Mede_|value|Justificativa|value|Nmero_da_OS|value"
  MM_columnsStr = "PI|none,none,NULL|[Data de início]|',none,''|[Data de Término]|',none,''|[Data de Término Contratual]|',none,''|[Duração da OS]|none,none,NULL|[Gerenciadora Mede ?]|none,1,0|Justificativa|',none,''|[Número da OS]|',none,''"

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
Dim rs_lista_med_ger
Dim rs_lista_med_ger_numRows

Set rs_lista_med_ger = Server.CreateObject("ADODB.Recordset")
rs_lista_med_ger.ActiveConnection = MM_cpf_STRING
rs_lista_med_ger.Source = "SELECT * FROM tb_Medicao_Gerenciadora ORDER BY PI ASC"
rs_lista_med_ger.CursorType = 0
rs_lista_med_ger.CursorLocation = 2
rs_lista_med_ger.LockType = 1
rs_lista_med_ger.Open()

rs_lista_med_ger_numRows = 0
%>
<%
Dim rs_inclui_med_ger
Dim rs_inclui_med_ger_numRows

Set rs_inclui_med_ger = Server.CreateObject("ADODB.Recordset")
rs_inclui_med_ger.ActiveConnection = MM_cpf_STRING
rs_inclui_med_ger.Source = "SELECT * FROM tb_Medicao_Gerenciadora"
rs_inclui_med_ger.CursorType = 0
rs_inclui_med_ger.CursorLocation = 2
rs_inclui_med_ger.LockType = 1
rs_inclui_med_ger.Open()

rs_inclui_med_ger_numRows = 0
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

Repeat1__numRows = 10
Repeat1__index = 0
rs_lista_med_ger_numRows = rs_lista_med_ger_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rs_lista_med_ger_total
Dim rs_lista_med_ger_first
Dim rs_lista_med_ger_last

' set the record count
rs_lista_med_ger_total = rs_lista_med_ger.RecordCount

' set the number of rows displayed on this page
If (rs_lista_med_ger_numRows < 0) Then
  rs_lista_med_ger_numRows = rs_lista_med_ger_total
Elseif (rs_lista_med_ger_numRows = 0) Then
  rs_lista_med_ger_numRows = 1
End If

' set the first and last displayed record
rs_lista_med_ger_first = 1
rs_lista_med_ger_last  = rs_lista_med_ger_first + rs_lista_med_ger_numRows - 1

' if we have the correct record count, check the other stats
If (rs_lista_med_ger_total <> -1) Then
  If (rs_lista_med_ger_first > rs_lista_med_ger_total) Then
    rs_lista_med_ger_first = rs_lista_med_ger_total
  End If
  If (rs_lista_med_ger_last > rs_lista_med_ger_total) Then
    rs_lista_med_ger_last = rs_lista_med_ger_total
  End If
  If (rs_lista_med_ger_numRows > rs_lista_med_ger_total) Then
    rs_lista_med_ger_numRows = rs_lista_med_ger_total
  End If
End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = rs_lista_med_ger
MM_rsCount   = rs_lista_med_ger_total
MM_size      = rs_lista_med_ger_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rs_lista_med_ger_first = MM_offset + 1
rs_lista_med_ger_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rs_lista_med_ger_first > MM_rsCount) Then
    rs_lista_med_ger_first = MM_rsCount
  End If
  If (rs_lista_med_ger_last > MM_rsCount) Then
    rs_lista_med_ger_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style18 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
.style30 {font-family: Arial, Helvetica, sans-serif; font-size: 11px; font-weight: bold; }
.style31 {font-family: Arial, Helvetica, sans-serif}
.style36 {	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
	font-weight: bold;
}
-->
</style>
</head>

<body>
<p align="center"><span class="style36">CADASTRO DE MEDI&Ccedil;&Otilde;ES GERECIADORAS</span></p>
<form method="POST" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style18">PI:</span></td>
      <td bgcolor="#CCCCCC"><label for="select"></label>
        <select name="PI" id="PI">
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
        </select>
</td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style18">Data de início:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="Data_de_incio" value="" size="15">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style18">Data de Término:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="Data_de_Trmino" value="" size="15">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style18">Data de Término Contratual:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="Data_de_Trmino_Contratual" value="" size="15">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style18">Duração da OS:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="Durao_da_OS" value="" size="18">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style18">Gerenciadora Mede ?:</span></td>
      <td bgcolor="#CCCCCC"><input <%If (CStr((rs_lista_med_ger.Fields.Item("Gerenciadora Mede ?").Value)) = CStr("True")) Then Response.Write("checked=""checked""") : Response.Write("")%> type="checkbox" name="Gerenciadora_Mede_" value=1 >      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style18">Justificativa:</span></td>
      <td bgcolor="#CCCCCC"><textarea name="Justificativa" cols="32"></textarea>      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style18">Número da OS:</span></td>
      <td bgcolor="#CCCCCC"><input type="text" name="Nmero_da_OS" value="" size="20">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC"><span class="style31"></span></td>
      <td bgcolor="#CCCCCC"><input type="submit" value="Salvar">      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1">
</form>
<div align="center">
  <table border="0">
    <tr bgcolor="#999999">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><span class="style30">PI</span></td>
      <td><span class="style30">C&oacute;digo do Pr&eacute;dio</span></td>
      <td><span class="style30">Nome da Unidade</span></td>
      <td><span class="style30">Gerenciadora Mede ?</span></td>
      <td><span class="style30">N&uacute;mero da OS</span></td>
      <td><span class="style30">Data de in&iacute;cio</span></td>
      <td><span class="style30">Dura&ccedil;&atilde;o da OS</span></td>
      <td><span class="style30">Data de T&eacute;rmino</span></td>
      <td><span class="style30">Data de T&eacute;rmino Contratual</span></td>
      <td><span class="style30">Justificativa</span></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rs_lista_med_ger.EOF)) %>
      <tr bgcolor="#CCCCCC">
        <td><a href="delete_med_ger.asp?pi=<%=(rs_lista_med_ger.Fields.Item("PI").Value)%>"><img src="depto/imagens/delete.gif" width="16" height="15" border="0" /></a></td>
        <td><img src="depto/imagens/edit.gif" width="16" height="15" /></td>
        <td><span class="style18"><%=(rs_lista_med_ger.Fields.Item("PI").Value)%></span></td>
        <td><span class="style18"><%=(rs_lista_med_ger.Fields.Item("Código do Prédio").Value)%></span></td>
        <td><span class="style18"><%=(rs_lista_med_ger.Fields.Item("Nome da Unidade").Value)%></span></td>
        <td><span class="style18"><%=(rs_lista_med_ger.Fields.Item("Gerenciadora Mede ?").Value)%></span></td>
        <td><span class="style18"><%=(rs_lista_med_ger.Fields.Item("Número da OS").Value)%></span></td>
        <td><span class="style18"><%=(rs_lista_med_ger.Fields.Item("Data de início").Value)%></span></td>
        <td><span class="style18"><%=(rs_lista_med_ger.Fields.Item("Duração da OS").Value)%></span></td>
        <td><span class="style18"><%=(rs_lista_med_ger.Fields.Item("Data de Término").Value)%></span></td>
        <td><span class="style18"><%=(rs_lista_med_ger.Fields.Item("Data de Término Contratual").Value)%></span></td>
        <td><span class="style18"><%=(rs_lista_med_ger.Fields.Item("Justificativa").Value)%></span></td>
      </tr>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_lista_med_ger.MoveNext()
Wend
%>
  </table>

    <table border="0" width="50%" align="center">
    <tr>
      <td width="23%" align="center"><% If MM_offset <> 0 Then %>
          <a href="<%=MM_moveFirst%>"><img src="First.gif" border=0></a>
          <% End If ' end MM_offset <> 0 %>
      </td>
      <td width="31%" align="center"><% If MM_offset <> 0 Then %>
          <a href="<%=MM_movePrev%>"><img src="Previous.gif" border=0></a>
          <% End If ' end MM_offset <> 0 %>
      </td>
      <td width="23%" align="center"><% If Not MM_atTotal Then %>
          <a href="<%=MM_moveNext%>"><img src="Next.gif" border=0></a>
          <% End If ' end Not MM_atTotal %>
      </td>
      <td width="23%" align="center"><% If Not MM_atTotal Then %>
          <a href="<%=MM_moveLast%>"><img src="Last.gif" border=0></a>
          <% End If ' end Not MM_atTotal %>
      </td>
    </tr>
  </table>
</div>
</body>
</html>
<%
rs_lista_med_ger.Close()
Set rs_lista_med_ger = Nothing
%>
<%
rs_inclui_med_ger.Close()
Set rs_inclui_med_ger = Nothing
%>
<%
rs_lista_pi.Close()
Set rs_lista_pi = Nothing
%>
