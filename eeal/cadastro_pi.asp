<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="1,3"
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
  MM_editTable = "tb_pi"
  MM_editRedirectUrl = "cadastro_pi.asp"
  MM_fieldsStr  = "PI|value|cod_predio|value|cod_construtora|value|cod_fiscal|value"
  MM_columnsStr = "PI|',none,''|cod_predio|',none,''|cod_construtora|none,none,NULL|cod_fiscal|none,none,NULL"

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
Dim rs_predio
Dim rs_predio_numRows

Set rs_predio = Server.CreateObject("ADODB.Recordset")
rs_predio.ActiveConnection = MM_cpf_STRING
rs_predio.Source = "SELECT tb_predio.cod_predio, [tb_predio].[cod_predio] & ' - ' & [tb_predio].[Nome_Unidade] AS Expr1, tb_predio.Município  FROM tb_predio LEFT JOIN tb_PI ON tb_predio.cod_predio = tb_PI.cod_predio  GROUP BY tb_predio.cod_predio, [tb_predio].[cod_predio] & ' - ' & [tb_predio].[Nome_Unidade], tb_predio.Município  ORDER BY [tb_predio].[cod_predio] & ' - ' & [tb_predio].[Nome_Unidade];  "
rs_predio.CursorType = 0
rs_predio.CursorLocation = 2
rs_predio.LockType = 1
rs_predio.Open()

rs_predio_numRows = 0
%>
<%
Dim rs_municipio__MMColParam
rs_municipio__MMColParam = "1"
If (Request.QueryString("cod_predio") <> "") Then 
  rs_municipio__MMColParam = Request.QueryString("cod_predio")
End If
%>
<%
Dim rs_municipio
Dim rs_municipio_numRows

Set rs_municipio = Server.CreateObject("ADODB.Recordset")
rs_municipio.ActiveConnection = MM_cpf_STRING
rs_municipio.Source = "SELECT tb_Municipios.cod_mun, tb_predio.cod_predio, tb_Municipios.Municipios  FROM tb_predio INNER JOIN tb_Municipios ON tb_predio.cod_mun = tb_Municipios.cod_mun  WHERE cod_predio = '" + Replace(rs_municipio__MMColParam, "'", "''") + "'  GROUP BY tb_Municipios.cod_mun, tb_predio.cod_predio, tb_Municipios.Municipios;    "
rs_municipio.CursorType = 0
rs_municipio.CursorLocation = 2
rs_municipio.LockType = 1
rs_municipio.Open()

rs_municipio_numRows = 0
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
Dim rs_inclui_pi
Dim rs_inclui_pi_numRows

Set rs_inclui_pi = Server.CreateObject("ADODB.Recordset")
rs_inclui_pi.ActiveConnection = MM_cpf_STRING
rs_inclui_pi.Source = "SELECT * FROM tb_PI"
rs_inclui_pi.CursorType = 0
rs_inclui_pi.CursorLocation = 2
rs_inclui_pi.LockType = 1
rs_inclui_pi.Open()

rs_inclui_pi_numRows = 0
%>
<%
Dim rs_lista_pi
Dim rs_lista_pi_numRows

Set rs_lista_pi = Server.CreateObject("ADODB.Recordset")
rs_lista_pi.ActiveConnection = MM_cpf_STRING
rs_lista_pi.Source = "SELECT tb_PI.PI, tb_predio.cod_predio+' - '+tb_predio.Nome_Unidade AS unidade, tb_responsavel.Responsável, tb_predio.Município, tb_Construtora.Construtora  FROM tb_responsavel INNER JOIN (tb_predio INNER JOIN (tb_Construtora INNER JOIN tb_PI ON tb_Construtora.cod_construtora = tb_PI.cod_construtora) ON tb_predio.cod_predio = tb_PI.cod_predio) ON tb_responsavel.cod_fiscal = tb_PI.cod_fiscal  ORDER BY 2;"
rs_lista_pi.CursorType = 0
rs_lista_pi.CursorLocation = 2
rs_lista_pi.LockType = 1
rs_lista_pi.Open()

rs_lista_pi_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rs_lista_pi_numRows = rs_lista_pi_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rs_lista_pi_total
Dim rs_lista_pi_first
Dim rs_lista_pi_last

' set the record count
rs_lista_pi_total = rs_lista_pi.RecordCount

' set the number of rows displayed on this page
If (rs_lista_pi_numRows < 0) Then
  rs_lista_pi_numRows = rs_lista_pi_total
Elseif (rs_lista_pi_numRows = 0) Then
  rs_lista_pi_numRows = 1
End If

' set the first and last displayed record
rs_lista_pi_first = 1
rs_lista_pi_last  = rs_lista_pi_first + rs_lista_pi_numRows - 1

' if we have the correct record count, check the other stats
If (rs_lista_pi_total <> -1) Then
  If (rs_lista_pi_first > rs_lista_pi_total) Then
    rs_lista_pi_first = rs_lista_pi_total
  End If
  If (rs_lista_pi_last > rs_lista_pi_total) Then
    rs_lista_pi_last = rs_lista_pi_total
  End If
  If (rs_lista_pi_numRows > rs_lista_pi_total) Then
    rs_lista_pi_numRows = rs_lista_pi_total
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

Set MM_rs    = rs_lista_pi
MM_rsCount   = rs_lista_pi_total
MM_size      = rs_lista_pi_numRows
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
rs_lista_pi_first = MM_offset + 1
rs_lista_pi_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rs_lista_pi_first > MM_rsCount) Then
    rs_lista_pi_first = MM_rsCount
  End If
  If (rs_lista_pi_last > MM_rsCount) Then
    rs_lista_pi_last = MM_rsCount
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
<title>Cadastro de PIs</title>
<style type="text/css">
<!--
.style7 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; color: #FFFFFF; }
.style9 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
.style10 {font-family: Arial, Helvetica, sans-serif}
.style11 {font-family: Arial, Helvetica, sans-serif;
	font-size: 24px;
	color: #333333;
}
.style13 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-weight: bold;
	color: #000066;
}
-->
</style>
<script type="text/JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += ' - '+nm+' é requerido.\n'; }
  } if (errors) alert('Preenchimento Obrigatório'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
</head>

<body>
<p align="center"><strong><span class="style11">Cadastro de PIs</span></strong></p>
<form method="POST" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">PI:</span></td>
      <td bgcolor="#CCCCCC"><input name="PI" type="text" class="style9" onblur="MM_validateForm('PI','','R');return document.MM_returnValue" value="" size="18">
        <span class="style13">ex: 2005/00468-0</span> </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Selecione o Prédio:</span></td>
      <td bgcolor="#CCCCCC"><select name="cod_predio" class="style9">
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
      </select></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Selecione a Construtora:</span></td>
      <td bgcolor="#CCCCCC"><select name="cod_construtora" class="style9">
        <%
While (NOT rs_construtora.EOF)
%>
        <option value="<%=(rs_construtora.Fields.Item("cod_construtora").Value)%>"><%=(rs_construtora.Fields.Item("Construtora").Value)%></option>
        <%
  rs_construtora.MoveNext()
Wend
If (rs_construtora.CursorType > 0) Then
  rs_construtora.MoveFirst
Else
  rs_construtora.Requery
End If
%>
      </select>      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Selecione o Fiscal:</span></td>
      <td bgcolor="#CCCCCC"><select name="cod_fiscal" class="style9">
        <%
While (NOT rs_fiscal.EOF)
%>
        <option value="<%=(rs_fiscal.Fields.Item("cod_fiscal").Value)%>"><%=(rs_fiscal.Fields.Item("Responsável").Value)%></option>
        <%
  rs_fiscal.MoveNext()
Wend
If (rs_fiscal.CursorType > 0) Then
  rs_fiscal.MoveFirst
Else
  rs_fiscal.Requery
End If
%>
      </select>      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10"></span></td>
      <td bgcolor="#CCCCCC"><input type="submit" value="Salvar">      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1">
</form>
<p>&nbsp;</p>

<div align="center"></div>
<div align="center">
  <table border="1">
    <tr bgcolor="#333333" class="style9">
      <td>&nbsp;</td>
      <td><span class="style7">PI</span></td>
      <td><span class="style7">Nome_Unidade</span></td>
      <td><span class="style7">Respons&aacute;vel</span></td>
      <td><span class="style7">Munic&iacute;pio</span></td>
      <td><span class="style7">Construtora</span></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rs_lista_pi.EOF)) %>
      <tr bgcolor="#CCCCCC" class="style9">
        <td><a href="delete_pi.asp?pi=<%=(rs_lista_pi.Fields.Item("PI").Value)%>"><img src="imagens/delete.gif" width="16" height="15" border="0" /></a></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("PI").Value)%></span></td>
        <td><div align="left"><%=(rs_lista_pi.Fields.Item("unidade").Value)%></div></td>
        <td><div align="left"><span class="style9"><%=(rs_lista_pi.Fields.Item("Responsável").Value)%></span></div></td>
        <td><div align="left"><%=(rs_lista_pi.Fields.Item("Município").Value)%></div></td>
        <td><div align="left"><span class="style9"><%=(rs_lista_pi.Fields.Item("Construtora").Value)%></span></div></td>
      </tr>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_lista_pi.MoveNext()
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
rs_predio.Close()
Set rs_predio = Nothing
%>
<%
rs_municipio.Close()
Set rs_municipio = Nothing
%>
<%
rs_construtora.Close()
Set rs_construtora = Nothing
%>
<%
rs_fiscal.Close()
Set rs_fiscal = Nothing
%>
<%
rs_inclui_pi.Close()
Set rs_inclui_pi = Nothing
%>
<%
rs_lista_pi.Close()
Set rs_lista_pi = Nothing
%>
