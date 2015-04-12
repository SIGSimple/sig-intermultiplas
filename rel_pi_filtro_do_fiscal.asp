<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
' technocurve arc 3 asp vb mv block1/3 start
Dim moColor1, moColor2, moColor
moColor1 = ""
moColor2 = ""
moColor3 = "#CCE6FF"
moColor = moColor1
' technocurve arc 3 asp vb mv block1/3 start
%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim rs_responsavel
Dim rs_responsavel_numRows

Set rs_responsavel = Server.CreateObject("ADODB.Recordset")
rs_responsavel.ActiveConnection = MM_cpf_STRING
rs_responsavel.Source = "SELECT tb_responsavel.cod_fiscal, tb_responsavel.Responsável  FROM tb_responsavel INNER JOIN tb_pi ON tb_responsavel.cod_fiscal = tb_pi.cod_fiscal  GROUP BY tb_responsavel.cod_fiscal, tb_responsavel.Responsável  ORDER BY tb_responsavel.Responsável;"
rs_responsavel.CursorType = 0
rs_responsavel.CursorLocation = 2
rs_responsavel.LockType = 1
rs_responsavel.Open()

rs_responsavel_numRows = 0
%>
<%
Dim rs_PIS__MMColParam
rs_PIS__MMColParam = "%%"
If (Request.Form("cod_fiscal")  <> "") Then 
  rs_PIS__MMColParam = Request.Form("cod_fiscal") 
End If
%>
<%
Dim rs_PIS__MMColParam1
rs_PIS__MMColParam1 = "%%"
If (Request.Form("cod_situacao")  <> "") Then 
  rs_PIS__MMColParam1 = Request.Form("cod_situacao") 
End If
%>
<%
Dim rs_PIS
Dim rs_PIS_numRows

Set rs_PIS = Server.CreateObject("ADODB.Recordset")
rs_PIS.ActiveConnection = MM_cpf_STRING
rs_PIS.Source = "SELECT tb_situacao_pi.desc_situacao, tb_pi.cod_predio, tb_pi.PI, tb_responsavel.Responsável, tb_predio.Nome_Unidade, tb_pi.cod_fiscal, tb_situacao_pi.cod_situacao  FROM tb_situacao_pi INNER JOIN (tb_responsavel INNER JOIN (tb_predio INNER JOIN tb_pi ON tb_predio.cod_predio = tb_pi.cod_predio) ON tb_responsavel.cod_fiscal = tb_pi.cod_fiscal) ON tb_situacao_pi.cod_situacao = tb_pi.cod_situacao  WHERE tb_pi.cod_fiscal like '" + Replace(rs_PIS__MMColParam, "'", "''") + "' and tb_pi.cod_situacao like '" + Replace(rs_PIS__MMColParam1, "'", "''") + "'  ORDER BY tb_pi.cod_predio"
rs_PIS.CursorType = 0
rs_PIS.CursorLocation = 2
rs_PIS.LockType = 1
rs_PIS.Open()

rs_PIS_numRows = 0
%>
<%
Dim rs_total__MMColParam
rs_total__MMColParam = "%%"
If (Request.Form("cod_fiscal")      <> "") Then 
  rs_total__MMColParam = Request.Form("cod_fiscal")     
End If
%>
<%
Dim rs_total__MMColParam1
rs_total__MMColParam1 = "%%"
If (Request.Form("cod_situacao")       <> "") Then 
  rs_total__MMColParam1 = Request.Form("cod_situacao")      
End If
%>
<%
Dim rs_total
Dim rs_total_numRows

Set rs_total = Server.CreateObject("ADODB.Recordset")
rs_total.ActiveConnection = MM_cpf_STRING
rs_total.Source = "SELECT tb_responsavel.Responsável,tb_situacao_pi.desc_situacao, tb_responsavel.cod_fiscal ,count([tb_pi.PI]) as conta  FROM tb_situacao_pi INNER JOIN (tb_responsavel INNER JOIN (tb_predio INNER JOIN tb_pi ON tb_predio.cod_predio = tb_pi.cod_predio) ON tb_responsavel.cod_fiscal = tb_pi.cod_fiscal) ON tb_situacao_pi.cod_situacao = tb_pi.cod_situacao  WHERE tb_pi.cod_fiscal like '" + Replace(rs_total__MMColParam, "'", "''") + "' and tb_situacao_pi.cod_situacao like '" + Replace(rs_total__MMColParam1, "'", "''") + "' GROUP BY tb_responsavel.Responsável,tb_situacao_pi.desc_situacao, tb_situacao_pi.cod_situacao, tb_responsavel.cod_fiscal"
rs_total.CursorType = 0
rs_total.CursorLocation = 2
rs_total.LockType = 1
rs_total.Open()

rs_total_numRows = 0
%>
<%
Dim rs_situacao
Dim rs_situacao_cmd
Dim rs_situacao_numRows

Set rs_situacao_cmd = Server.CreateObject ("ADODB.Command")
rs_situacao_cmd.ActiveConnection = MM_cpf_STRING
rs_situacao_cmd.CommandText = "SELECT tb_situacao_pi.cod_situacao, tb_situacao_pi.desc_situacao FROM tb_situacao_pi ORDER BY tb_situacao_pi.desc_situacao; " 
rs_situacao_cmd.Prepared = true

Set rs_situacao = rs_situacao_cmd.Execute
rs_situacao_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 25
Repeat1__index = 0
rs_PIS_numRows = rs_PIS_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
rs_total_numRows = rs_total_numRows + Repeat2__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rs_PIS_total
Dim rs_PIS_first
Dim rs_PIS_last

' set the record count
rs_PIS_total = rs_PIS.RecordCount

' set the number of rows displayed on this page
If (rs_PIS_numRows < 0) Then
  rs_PIS_numRows = rs_PIS_total
Elseif (rs_PIS_numRows = 0) Then
  rs_PIS_numRows = 1
End If

' set the first and last displayed record
rs_PIS_first = 1
rs_PIS_last  = rs_PIS_first + rs_PIS_numRows - 1

' if we have the correct record count, check the other stats
If (rs_PIS_total <> -1) Then
  If (rs_PIS_first > rs_PIS_total) Then
    rs_PIS_first = rs_PIS_total
  End If
  If (rs_PIS_last > rs_PIS_total) Then
    rs_PIS_last = rs_PIS_total
  End If
  If (rs_PIS_numRows > rs_PIS_total) Then
    rs_PIS_numRows = rs_PIS_total
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

Set MM_rs    = rs_PIS
MM_rsCount   = rs_PIS_total
MM_size      = rs_PIS_numRows
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
rs_PIS_first = MM_offset + 1
rs_PIS_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rs_PIS_first > MM_rsCount) Then
    rs_PIS_first = MM_rsCount
  End If
  If (rs_PIS_last > MM_rsCount) Then
    rs_PIS_last = MM_rsCount
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
.style5 {font-family: Arial, Helvetica, sans-serif; font-size: 14px; font-weight: bold; }
.style7 {font-family: Arial, Helvetica, sans-serif; font-size: 14px; font-weight: bold; color: #FFFFFF; }
.style9 {font-family: Arial, Helvetica, sans-serif; font-size: 18px; font-weight: bold; }
.style10 {font-size: 14px; font-family: Arial, Helvetica, sans-serif;}
.style12 {font-family: Arial, Helvetica, sans-serif; font-size: 18px; font-weight: bold; color: #003399; }
.style14 {font-size: 12px; font-family: Arial, Helvetica, sans-serif; }
.style19 {font-size: 12px}
.style21 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; color: #FFFF33; }
.style22 {color: #FFFF33}
-->
</style>
</head>

<body>
<form id="form1" name="form1" method="post" action="">
  <label for="cod_situacao"></label>
  <span class="style10">Fiscal</span>
  <select name="Cod_fiscal" class="style10" id="Cod_fiscal">
    <option value="">todos..</option>
    <%
While (NOT rs_responsavel.EOF)
%>
    <option value="<%=(rs_responsavel.Fields.Item("cod_fiscal").Value)%>"><%=(rs_responsavel.Fields.Item("Responsável").Value)%></option>
    <%
  rs_responsavel.MoveNext()
Wend
If (rs_responsavel.CursorType > 0) Then
  rs_responsavel.MoveFirst
Else
  rs_responsavel.Requery
End If
%>
  </select>
  <label>
  <span class="style10">Situa&ccedil;&atilde;o</span>
  <select name="cod_situacao" id="cod_situacao">
    <option value="">todas..</option>
    <%
While (NOT rs_situacao.EOF)
%>
    <option value="<%=(rs_situacao.Fields.Item("cod_situacao").Value)%>"><%=(rs_situacao.Fields.Item("desc_situacao").Value)%></option>
    <%
  rs_situacao.MoveNext()
Wend
If (rs_situacao.CursorType > 0) Then
  rs_situacao.MoveFirst
Else
  rs_situacao.Requery
End If
%>
  </select>
</label>
  <label for="Submit"></label>
  <input type="submit" name="Submit" value="Buscar" id="Submit" />
  <span class="style12"></span>
  <span class="style9"></span>
</form>

<p class="style14"><a href="rel_pi_filtrofiscal.asp" target="_blank"><U>RELAT&Oacute;RIO DE PLANEJAMENTO E CONTROLE DAS MEDI&Ccedil;&Otilde;ES</U></a></p>
<p class="style14">&nbsp;</p>
<table border="0">
  
  <tr bgcolor="#999999">
    <td width="127" class="style7">cod_predio</td>
    <td width="383" class="style7">Nome Unidade</td>
    <td width="131" class="style7">PI</td>
    <td colspan="2"><div align="center" class="style9">A&Ccedil;&Otilde;ES</div></td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT rs_PIS.EOF)) %>
    <tr <%
' technocurve arc 3 asp vb mv block2/3 start
Response.Write(" style='background-color:" & moColor & "' onMouseOver='this.style.backgroundColor=" & chr(34) & moColor3 & chr(34) & "' onMouseOut='this.style.backgroundColor=" & chr(34) & moColor & chr(34) & "'")
' technocurve arc 3 asp vb mv block2/3 start
%> class="style10">
      <td class="style10 style19"><%=(rs_PIS.Fields.Item("cod_predio").Value)%></td>
      <td class="style14"><%=(rs_PIS.Fields.Item("Nome_Unidade").Value)%></td>
      <td class="style14"><%=(rs_PIS.Fields.Item("PI").Value)%></td>
      <td width="146" class="style14 style19"><%=(rs_PIS.Fields.Item("desc_situacao").Value)%></td>
      <td width="156" class="style5"><div align="center"><a href="acompanhamento_inclui.asp?pi=<%=(rs_PIS.Fields.Item("PI").Value)%>" target="_blank" class="style19">Acompanhamento</a></div></td>
    </tr>
    <%
' technocurve arc 3 asp vb mv block3/3 start
if moColor = moColor1 then
	moColor = moColor2
else
	moColor = moColor1
end if
' technocurve arc 3 asp vb mv block3/3 start
%>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_PIS.MoveNext()
Wend
%>
</table>

<table border="0" width="50%" align="center">
  <tr>
    <td width="23%" align="center"><% If MM_offset <> 0 Then %>
        <a href="<%=MM_moveFirst%>"><img src="First.gif" border="0" /></a>
        <% End If ' end MM_offset <> 0 %>
    </td>
    <td width="31%" align="center"><% If MM_offset <> 0 Then %>
        <a href="<%=MM_movePrev%>"><img src="Previous.gif" border="0" /></a>
        <% End If ' end MM_offset <> 0 %>
    </td>
    <td width="23%" align="center"><% If Not MM_atTotal Then %>
        <a href="<%=MM_moveNext%>"><img src="Next.gif" border="0" /></a>
        <% End If ' end Not MM_atTotal %>
    </td>
    <td width="23%" align="center"><% If Not MM_atTotal Then %>
        <a href="<%=MM_moveLast%>"><img src="Last.gif" border="0" /></a>
        <% End If ' end Not MM_atTotal %>
    </td>
  </tr>
</table>
<table width="423" border="1">
  <tr>
    <td width="83" bgcolor="#333333" class="style12 style19 style22"><div align="center" class="style21"><%=(rs_total.Fields.Item("Responsável").Value)%></div></td>
    <td width="235" bgcolor="#333333" class="style12 style19 style22"><div align="center">Situa&ccedil;&atilde;o</div></td>
    <td width="83" bgcolor="#333333" class="style21"><div align="center">Obras</div></td>
  </tr>
  <% 
While ((Repeat2__numRows <> 0) AND (NOT rs_total.EOF)) 
%>
    <tr>
      <td colspan="2" bgcolor="#CCCCCC" class="style14"><%=(rs_total.Fields.Item("desc_situacao").Value)%></td>
      <td bgcolor="#CCCCCC" class="style14"><div align="center"><%=(rs_total.Fields.Item("conta").Value)%></div></td>
    </tr>
    <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rs_total.MoveNext()
Wend
%>
</table>
<p>&nbsp;</p>
</body>
</html>
<%
rs_responsavel.Close()
Set rs_responsavel = Nothing
%>
<%
rs_PIS.Close()
Set rs_PIS = Nothing
%>
<%
rs_total.Close()
Set rs_total = Nothing
%>
<%
rs_situacao.Close()
Set rs_situacao = Nothing
%>
