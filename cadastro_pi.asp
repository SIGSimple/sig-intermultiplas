<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="1,3,4"
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
  MM_fieldsStr  = "PI|value|cod_predio|value|cod_fiscal|value"
  MM_columnsStr = "PI|',none,''|cod_predio|',none,''|cod_fiscal|none,none,NULL"

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

  Dim frmField_PI
  Dim frmField_cod_predio
  Dim frmField_nome_empreendimento
  Dim frmField_cod_tipo_empreendimento
  Dim frmField_cod_programa
  Dim frmField_dsc_situacao_anterior
  Dim frmField_dsc_situacao_atual
  Dim frmField_dsc_resultado_obtido
  Dim frmField_endereco
  Dim frmField_cep
  Dim frmField_latitude_longitude
  Dim frmField_email
  Dim frmField_telefone
  Dim frmField_qtd_populacao_urbana_2010
  Dim frmField_qtd_populacao_urbana_2030
  Dim frmField_cod_fiscal
  Dim frmField_cod_engenheiro_daee
  Dim frmField_cod_engenheiro_plan_consorcio
  Dim frmField_cod_fiscal_consorcio
  Dim frmField_cod_engenheiro_construtora
  Dim frmField_cod_situacao
  Dim frmField_cod_situacao_externa

  Set frmField_PI = Request.Form("PI")
  Set frmField_cod_predio = Request.Form("cod_predio")
      frmField_cod_predio = Split(frmField_cod_predio, "-")
  Set frmField_nome_empreendimento = Request.Form("nome_empreendimento")
  Set frmField_descricao_empreendimento = Request.Form("descricao_empreendimento")
  Set frmField_cod_tipo_empreendimento = Request.Form("cod_tipo_empreendimento")
  Set frmField_cod_programa = Request.Form("cod_programa")
  Set frmField_dsc_resultado_obtido = Request.Form("dsc_resultado_obtido")
  'Set frmField_dsc_situacao_anterior = Request.Form("dsc_situacao_anterior")
  'Set frmField_dsc_situacao_atual = Request.Form("dsc_situacao_atual")
  Set frmField_endereco = Request.Form("endereco")
  Set frmField_cep = Request.Form("cep")
  Set frmField_latitude_longitude = Request.Form("latitude_longitude")
  Set frmField_email = Request.Form("email")
  Set frmField_telefone = Request.Form("telefone")
  Set frmField_qtd_populacao_urbana_2010 = Request.Form("qtd_populacao_urbana_2010")
  'Set frmField_qtd_populacao_urbana_2030 = Request.Form("qtd_populacao_urbana_2030")
  Set frmField_cod_fiscal = Request.Form("cod_fiscal")
  Set frmField_cod_engenheiro_daee = Request.Form("cod_engenheiro_daee")
  Set frmField_cod_engenheiro_plan_consorcio = Request.Form("cod_engenheiro_plan_consorcio")
  Set frmField_cod_fiscal_consorcio = Request.Form("cod_fiscal_consorcio")
  Set frmField_cod_engenheiro_medicao = Request.Form("cod_engenheiro_medicao")
  Set frmField_cod_engenheiro_construtora = Request.Form("cod_engenheiro_construtora")
  Set frmField_cod_situacao = Request.Form("cod_situacao")
  Set frmField_cod_situacao_externa = Request.Form("cod_situacao_externa")

  MM_editQuery = "insert into "& MM_editTable &" ("
  MM_editQuery = MM_editQuery & "PI,cod_predio,id_predio,municipio,nome_empreendimento,[Descrição da Intervenção FDE],cod_tipo_empreendimento,cod_programa"
  'MM_editQuery = MM_editQuery & ",dsc_situacao_anterior,dsc_situacao_atual,dsc_resultado_obtido,endereco,cep"
  MM_editQuery = MM_editQuery & ",dsc_resultado_obtido,endereco,cep"
  MM_editQuery = MM_editQuery & ",latitude_longitude"
  MM_editQuery = MM_editQuery & ",email,telefone,qtd_populacao_urbana_2010" ',qtd_populacao_urbana_2030'
  MM_editQuery = MM_editQuery & ",cod_fiscal,cod_engenheiro_daee,cod_engenheiro_plan_consorcio,cod_fiscal_consorcio,cod_engenheiro_medicao,cod_engenheiro_construtora,cod_situacao"
  MM_editQuery = MM_editQuery & ",cod_situacao_externa"
  MM_editQuery = MM_editQuery & ") values ("
  MM_editQuery = MM_editQuery & "'" & frmField_PI                              & "'," 'PI'
  MM_editQuery = MM_editQuery & "'" & frmField_cod_predio(1)                   & "'," 'cod_predio'
  MM_editQuery = MM_editQuery & ""  & frmField_cod_predio(0)                   & ","  'id_predio' 
  MM_editQuery = MM_editQuery & "'" & frmField_cod_predio(2)                   & "'," 'municipio'
  MM_editQuery = MM_editQuery & "'" & frmField_nome_empreendimento             & "'," 'nome_empreendimento'
  MM_editQuery = MM_editQuery & "'" & frmField_descricao_empreendimento        & "'," '[Descrição da Intervenção FDE]'
  MM_editQuery = MM_editQuery & ""  & frmField_cod_tipo_empreendimento         & ","  'cod_tipo_empreendimento'
  MM_editQuery = MM_editQuery & ""  & frmField_cod_programa                    & ""  'cod_programa'
  'MM_editQuery = MM_editQuery & ",'" & frmField_dsc_situacao_anterior         & "'" 'dsc_situacao_anterior'
  'MM_editQuery = MM_editQuery & ",'" & frmField_dsc_situacao_atual            & "'" 'dsc_situacao_atual'
  MM_editQuery = MM_editQuery & ",'" & frmField_dsc_resultado_obtido          & "'" 'dsc_resultado_obtido'
  MM_editQuery = MM_editQuery & ",'" & frmField_endereco                      & "'" 'endereco'
  MM_editQuery = MM_editQuery & ",'" & frmField_cep                           & "'" 'cep'
  MM_editQuery = MM_editQuery & ",'" & frmField_latitude_longitude              & "'" 'latitude_longitude'
  MM_editQuery = MM_editQuery & ",'" & frmField_email                         & "'" 'email'
  MM_editQuery = MM_editQuery & ",'" & frmField_telefone                      & "'" 'telefone'
  MM_editQuery = MM_editQuery & ","  & frmField_qtd_populacao_urbana_2010     & ""  'qtd_populacao_urbana_2010'
  'MM_editQuery = MM_editQuery & ","  & frmField_qtd_populacao_urbana_2030     & ""  'qtd_populacao_urbana_2030'
  MM_editQuery = MM_editQuery & ","  & frmField_cod_fiscal                    & ""  'cod_fiscal'
  MM_editQuery = MM_editQuery & ","  & frmField_cod_engenheiro_daee           & ""  'cod_engenheiro_daee'
  MM_editQuery = MM_editQuery & ","  & frmField_cod_engenheiro_plan_consorcio & ""  'cod_engenheiro_plan_consorcio'
  MM_editQuery = MM_editQuery & ","  & frmField_cod_fiscal_consorcio          & ""  'cod_fiscal_consorcio'
  MM_editQuery = MM_editQuery & ","  & frmField_cod_engenheiro_medicao        & ""  'cod_engenheiro_medicao'
  MM_editQuery = MM_editQuery & ","  & frmField_cod_engenheiro_construtora    & ""  'cod_engenheiro_construtora'
  MM_editQuery = MM_editQuery & ","  & frmField_cod_situacao                  & ""  'cod_situacao'
  MM_editQuery = MM_editQuery & ","  & frmField_cod_situacao_externa          & ""  'cod_situacao_externa'
  
  MM_editQuery = MM_editQuery & ")"

  If (Not MM_abortEdit) Then
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
Dim rs_tipo_empreendimento
Dim rs_tipo_empreendimento_numRows

Set rs_tipo_empreendimento = Server.CreateObject("ADODB.Recordset")
rs_tipo_empreendimento.ActiveConnection = MM_cpf_STRING
rs_tipo_empreendimento.Source = "SELECT * FROM tb_tipo_empreendimento;  "
rs_tipo_empreendimento.CursorType = 0
rs_tipo_empreendimento.CursorLocation = 2
rs_tipo_empreendimento.LockType = 1
rs_tipo_empreendimento.Open()

rs_tipo_empreendimento_numRows = 0
%>

<%
Dim rs_programa
Dim rs_programa_numRows

Set rs_programa = Server.CreateObject("ADODB.Recordset")
rs_programa.ActiveConnection = MM_cpf_STRING
rs_programa.Source = "SELECT * FROM tb_depto;  "
rs_programa.CursorType = 0
rs_programa.CursorLocation = 2
rs_programa.LockType = 1
rs_programa.Open()

rs_programa_numRows = 0
%>

<%
Dim rs_predio
Dim rs_predio_numRows

Set rs_predio = Server.CreateObject("ADODB.Recordset")
rs_predio.ActiveConnection = MM_cpf_STRING
rs_predio.Source = "SELECT tb_predio.id_predio, tb_predio.cod_predio, [tb_predio].[Município] AS Expr1, tb_predio.Município  FROM tb_predio LEFT JOIN tb_PI ON tb_predio.cod_predio = tb_PI.cod_predio  GROUP BY tb_predio.id_predio, tb_predio.cod_predio, [tb_predio].[Município]  ORDER BY [tb_predio].[Município];  "
rs_predio.CursorType = 0
rs_predio.CursorLocation = 2
rs_predio.LockType = 1
rs_predio.Open()

rs_predio_numRows = 0
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
Dim rs_situacao
Dim rs_situacao_numRows

Set rs_situacao = Server.CreateObject("ADODB.Recordset")
rs_situacao.ActiveConnection = MM_cpf_STRING
rs_situacao.Source = "SELECT *  FROM tb_situacao_pi  ORDER BY desc_situacao ASC"
rs_situacao.CursorType = 0
rs_situacao.CursorLocation = 2
rs_situacao.LockType = 1
rs_situacao.Open()

rs_situacao_numRows = 0
%>
<%
Dim rs_lista_pi
Dim rs_lista_pi_numRows

Set rs_lista_pi = Server.CreateObject("ADODB.Recordset")
rs_lista_pi.ActiveConnection = MM_cpf_STRING
rs_lista_pi.Source = "SELECT * FROM c_lista_dados_obras ORDER BY Município, nome_empreendimento"
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
<title>Cadastro de Empreendimentos</title>
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

function validateForm(event) {  
  var num_autos     = event.target.PI.value;
  var municipio     = event.target.cod_predio.value;
  var eng_consorcio = event.target.cod_fiscal.value;

  if (num_autos == "") {
    alert("você deve informar o campo 'Autos'");
    event.preventDefault();
    return false;
  }

  if (municipio == "") {
    alert("você deve informar o campo 'Município'");
    event.preventDefault();
    return false;
  }

  if (eng_consorcio == "") {
    alert("você deve informar o campo 'Eng. Obras Consórcio'");
    event.preventDefault();
    return false;
  }

  return true;
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
<script src="//code.jquery.com/jquery-1.11.2.min.js"></script>
<script type="text/javascript" src="js/jquery.floatThead.min.js"></script>
<script type="text/javascript">
  $(function(){
    $("table#data").floatThead();
  });
</script>
</head>

<body>
  <p align="center">
    <strong>
      <span class="style11">Cadastro de Empreendimentos</span>
    </strong>
  </p>

  <%
    If Session("MM_UserAuthorization") = 1 Or Session("MM_UserAuthorization") = 3 Then 
  %>
  <form method="POST" action="<%=MM_editAction%>" name="form1" onsubmit="validateForm(event)">
    <input type="hidden" name="MM_insert" value="form1">

    <table align="center">
      <!-- MUNICÍPIO -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10">Selecione o Município:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <select name="cod_predio" class="style9">
            <option value=""></option>
            <%
              While (NOT rs_predio.EOF)
            %>
            <option value="<%=(rs_predio.Fields.Item("id_predio").Value)%>-<%=(rs_predio.Fields.Item("cod_predio").Value)%>-<%=(rs_predio.Fields.Item("Expr1").Value)%>"><%=(rs_predio.Fields.Item("Expr1").Value)%></option>
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

      <!-- NOME DO EMPREENDIMENTO -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10">Localidade:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <input name="nome_empreendimento" type="text" class="style9" value="">
        </td>
      </tr>

      <!-- AUTOS -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10">Autos:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <input name="PI" type="text" class="style9" value="" size="18">
        </td>
      </tr>

      <!-- DESCRIÇÃO DO EMPREENDIMENTO -->
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10">Objeto da Obra:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <textarea name="descricao_empreendimento" cols="50" rows="5" class="style9" style="width: 98%;"></textarea>
        </td>
      </tr>

      <!-- TIPO -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10">Tipo:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <select name="cod_tipo_empreendimento" class="style9">
            <option value=""></option>
            <%
              While (NOT rs_tipo_empreendimento.EOF)
            %>
            <option value="<%=(rs_tipo_empreendimento.Fields.Item("id").Value)%>"><%=(rs_tipo_empreendimento.Fields.Item("desc_tipo").Value)%></option>
            <%
                rs_tipo_empreendimento.MoveNext()
              Wend
              If (rs_tipo_empreendimento.CursorType > 0) Then
                rs_tipo_empreendimento.MoveFirst
              Else
                rs_tipo_empreendimento.Requery
              End If
            %>
          </select>
        </td>
      </tr>

      <!-- PROGRAMA -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10">Programa:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <select name="cod_programa" class="style9">
            <option value=""></option>
            <%
              While (NOT rs_programa.EOF)
            %>
            <option value="<%=(rs_programa.Fields.Item("cod_depto").Value)%>"><%=(rs_programa.Fields.Item("sigla").Value)%> - <%=(rs_programa.Fields.Item("desc_depto").Value)%></option>
            <%
                rs_programa.MoveNext()
              Wend
              If (rs_programa.CursorType > 0) Then
                rs_programa.MoveFirst
              Else
                rs_programa.Requery
              End If
            %>
          </select>
        </td>
      </tr>

      <!-- BENEFÍCIO GERAL DA OBRA -->
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Benefício Geral da Obra:</span></td>
        <td bgcolor="#CCCCCC">
          <textarea name="dsc_resultado_obtido"cols="50" rows="5" class="style9" style="width: 98%;"></textarea>
        </td>
      </tr>

      <!-- SITUAÇÃO ANTERIOR -->
      <!-- <tr valign="baseline">
        <td align="right" valign="middle" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Situação Anterior:</span></td>
        <td bgcolor="#CCCCCC">
          <textarea name="dsc_situacao_anterior"cols="50" rows="5" class="style9" style="width: 98%;"></textarea>
        </td>
      </tr> -->

      <!-- SITUAÇÃO ATUAL -->
      <!-- <tr valign="baseline">
        <td align="right" valign="middle" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Situação Atual:</span></td>
        <td bgcolor="#CCCCCC">
          <textarea name="dsc_situacao_atual"cols="50" rows="5" class="style9" style="width: 98%;"></textarea>
        </td>
      </tr> -->

      <!-- ENDEREÇO -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Endereço:</span></td>
        <td bgcolor="#CCCCCC">
          <input name="endereco" type="text" class="style9" value="">
        </td>
      </tr>

      <!-- CEP -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">CEP:</span></td>
        <td bgcolor="#CCCCCC">
          <input name="cep" type="text" class="style9" value="">
        </td>
      </tr>

      <!-- LOCALIZAÇÃO GEOGRÁFICA (LAT,LONG) -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Localização Geográfica (Lat, Long):</span></td>
        <td bgcolor="#CCCCCC">
          <input name="latitude_longitude" type="text" class="style9" value="">
        </td>
      </tr>

      <!-- EMAIL -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Email:</span></td>
        <td bgcolor="#CCCCCC">
          <input name="email" type="text" class="style9" value="">
        </td>
      </tr>

      <!-- TELEFONE -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Telefone:</span></td>
        <td bgcolor="#CCCCCC">
          <input name="telefone" type="text" class="style9" value="">
        </td>
      </tr>

      <!-- POPULAÇÃO 2010 -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">População Urbana - IBGE (2010) (hab):</span></td>
        <td bgcolor="#CCCCCC">
          <input name="qtd_populacao_urbana_2010" type="text" class="style9" value="">
        </td>
      </tr>

      <!-- PROJEÇÃO POPULAÇÃO 2030 -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Projeção de População (2030):</span></td>
        <td bgcolor="#CCCCCC">
          <input name="qtd_populacao_urbana_2030" type="text" class="style9" value="" disabled="disabled">
        </td>
      </tr>

      <!-- ENG. OBRAS CONSÓRCIO -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10">Selecione o Eng. Obras Consórcio:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <select name="cod_fiscal" class="style9">
            <option value=""></option>
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
          </select>
        </td>
      </tr>

      <!-- ENG. DAEE -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10">Selecione o Eng. DAEE:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <select name="cod_engenheiro_daee" class="style9">
            <option value=""></option>
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
          </select>
        </td>
      </tr>

      <!-- ENG. PLAN. OBRAS CONSÓRCIO -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10">Selecione o Eng. Plan. Obras Consórcio:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <select name="cod_engenheiro_plan_consorcio" class="style9">
            <option value=""></option>
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
          </select>
        </td>
      </tr>

      <!-- FISCAL -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Selecione o Fiscal do Consórcio:</span></td>
        <td bgcolor="#CCCCCC">
          <select name="cod_fiscal_consorcio" class="style9">
            <option value=""></option>
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
          </select>
        </td>
      </tr>

      <!-- ENG. RESP. MEDIÇÕES -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Selecione o Eng. Resp. Medições:</span></td>
        <td bgcolor="#CCCCCC">
          <select name="cod_engenheiro_medicao" class="style9">
            <option value=""></option>
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
          </select>
        </td>
      </tr>

      <!-- ENG. OBRAS CONSTRUTORA -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10">Selecione o Eng. Obras Construtora:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <select name="cod_engenheiro_construtora" class="style9">
            <option value=""></option>
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
          </select>
        </td>
      </tr>

      <!-- SITUAÇÃO INTERNA -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10">Situação da Obra:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <select name="cod_situacao" class="style9">
            <option value=""></option>
            <%
              While (NOT rs_situacao.EOF)
            %>
            <option value="<%=(rs_situacao.Fields.Item("cod_situacao").Value)%>">Status: <%=(rs_situacao.Fields.Item("desc_situacao").Value)%> - Situação: <%=(rs_situacao.Fields.Item("cod_atendimento").Value)%></option>
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
        </td>
      </tr>

      <!-- SITUAÇÃO EXTERNA -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10">Situação Atual do Empreendimento:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <select name="cod_situacao_externa" class="style9">
            <option value=""></option>
            <%
              While (NOT rs_situacao.EOF)
            %>
            <option value="<%=(rs_situacao.Fields.Item("cod_situacao").Value)%>">Status: <%=(rs_situacao.Fields.Item("desc_situacao").Value)%> - Situação: <%=(rs_situacao.Fields.Item("cod_atendimento").Value)%></option>
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
        </td>
      </tr>

      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10"></span>
        </td>
        <td bgcolor="#CCCCCC">
          <input type="submit" value="Salvar">
        </td>
      </tr>
    </table>
  </form>

  <%
    End If
  %>

  <div align="center">
    <table id="data" border="0">
      <thead>
        <tr bgcolor="#333333" class="style9 table-title">
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td><span class="style7">Munic&iacute;pio</span></td>
          <td style="min-width: 150px;"><span class="style7">Localidade</span></td>
          <td style="min-width: 80px;"><span class="style7">Autos</span></td>
          <td style="min-width: 80px;"><span class="style7">Bacia DAEE</span></td>
          <td style="min-width: 150px;"><span class="style7">Objeto da Obra</span></td>
          <td style="min-width: 100px;"><span class="style7">Tipo</span></td>
          <td style="min-width: 150px;"><span class="style7">Programa</span></td>
          <td style="min-width: 150px;"><span class="style7">Localização Geográfica (Lat,Long)</span></td>
          <!-- <td style="min-width: 150px;"><span class="style7">Situação Anterior</span></td>
          <td style="min-width: 150px;"><span class="style7">Situação Atual</span></td> -->
          <td style="min-width: 150px;"><span class="style7">Beneficio Geral da Obra</span></td>
          <td style="min-width: 150px;"><span class="style7">Endereço</span></td>
          <td style="min-width: 80px;"><span class="style7">CEP</span></td>
          <td style="min-width: 100px;"><span class="style7">E-Mail</span></td>
          <td style="min-width: 100px;"><span class="style7">Telefone</span></td>
          <td style="min-width: 150px;"><span class="style7">População Urbana - IBGE (2010) (hab)</span></td>
          <td style="min-width: 150px;"><span class="style7">Projeção de População (2030)</span></td>
          <td style="min-width: 150px;"><span class="style7">Investimento Governo SP</span></td>
          <td style="min-width: 150px;"><span class="style7">Início das Obras</span></td>
          <td style="min-width: 150px;"><span class="style7">% Executado</span></td>
          <td style="min-width: 150px;"><span class="style7">Previsão de Término</span></td>
          <td style="min-width: 150px;"><span class="style7">Concluída/Inaugurada em</span></td>
          <td style="min-width: 150px;"><span class="style7">Previsão de Inauguração</span></td>
          <td style="min-width: 150px;"><span class="style7">Carga Orgânica Retirada (ton./mês)</span></td>
          <td style="min-width: 150px;"><span class="style7">Eng. Obras Consórcio</span></td>
          <td style="min-width: 150px;"><span class="style7">Eng. DAEE</span></td>
          <td style="min-width: 150px;"><span class="style7">Eng. Plan. Obras Consórcio</span></td>
          <td style="min-width: 150px;"><span class="style7">Fiscal do Consórcio</span></td>
          <td style="min-width: 150px;"><span class="style7">Eng. Resp. Medições</span></td>
          <td style="min-width: 150px;"><span class="style7">Eng. Obras Construtora</span></td>
          <td style="min-width: 150px;"><span class="style7">Situação Interna</span></td>
          <td style="min-width: 150px;"><span class="style7">Situação Atual do Empreendimento</span></td>
          <td style="min-width: 150px;"><span class="style7">Observações Rel. Mensal</span></td>
        </tr>
      </thead>
      <%
        While ((Repeat1__numRows <> 0) AND (NOT rs_lista_pi.EOF))
      %>
      <tr bgcolor="#CCCCCC" class="style9">
        <td><a href="altera_pi.asp?pi=<%=(rs_lista_pi.Fields.Item("PI").Value)%>"><img src="depto/imagens/edit.gif" width="16" height="15" border="0" /></a></td>
        <td><a href="delete_pi.asp?pi=<%=(rs_lista_pi.Fields.Item("PI").Value)%>"><img src="const/imagens/delete.gif" width="16" height="15" border="0" /></a></td>
        <td><a href="image_tool.asp?cod_empreendimento=<%=(rs_lista_pi.Fields.Item("PI").Value)%>&pg=1">Fotos</a></td>
        <td><div align="left"><%=(rs_lista_pi.Fields.Item("Município").Value)%></div></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("nome_empreendimento").Value)%></span></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("PI").Value)%></span></td>
        <td align="center"><span class="style9"><%=( Mid(rs_lista_pi.Fields.Item("bacia_daee").Value, 1, 3) )%></span></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("Descrição da Intervenção FDE").Value)%></span></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("tipo_empreendimento").Value)%></span></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("programa").Value)%></span></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("latitude_longitude").Value)%></span></td>
        <!-- <td><span class="style9"><%=(rs_lista_pi.Fields.Item("dsc_situacao_anterior").Value)%></span></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("dsc_situacao_atual").Value)%></span></td> -->
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("dsc_resultado_obtido").Value)%></span></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("endereco").Value)%></span></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("cep").Value)%></span></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("email").Value)%></span></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("telefone").Value)%></span></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("qtd_populacao_urbana_2010").Value)%></span></td>
        <td>
          <span class="style9">
            <%
              If Not IsNull(rs_lista_pi.Fields.Item("qtd_populacao_urbana_2010").Value) Then

                data = rs_lista_pi.Fields.Item("qtd_populacao_urbana_2010").Value
                data = data * 1.25
                a = Round(data/100, 0)
                b = a * 100

                Response.Write b
              End If
            %>
          </span>
        </td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("Valor do Contrato").Value)%></span></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("dta_inicio_obras").Value)%></span></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("num_percentual_executado").Value)%></span></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("dta_previsao_termino").Value)%></span></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("dta_inauguracao").Value)%></span></td>
        <td><span class="style9"><%=(rs_lista_pi.Fields.Item("dta_previsao_inauguracao").Value)%></span></td>
        <td>
          <span class="style9">
            <%
              If Not IsNull(rs_lista_pi.Fields.Item("qtd_populacao_urbana_2010").Value) Then
                If Not IsNull(b) Then
                  ' Base de cálculo = qtd_populacao_urbana_2030 * 0,06 * 30 / 1000
                  Response.Write b * 0.0018
                End If
              End If
            %>
          </span>
        </td>
        <td><div align="left"><span class="style9"><%=(rs_lista_pi.Fields.Item("eng_obras_consorcio").Value)%></span></div></td>
        <td><div align="left"><span class="style9"><%=(rs_lista_pi.Fields.Item("eng_daee").Value)%></span></div></td>
        <td><div align="left"><span class="style9"><%=(rs_lista_pi.Fields.Item("eng_plan_consorcio").Value)%></span></div></td>
        <td><div align="left"><span class="style9"><%=(rs_lista_pi.Fields.Item("fiscal_consorcio").Value)%></span></div></td>
        <td><div align="left"><span class="style9"><%=(rs_lista_pi.Fields.Item("eng_medicao").Value)%></span></div></td>
        <td><div align="left"><span class="style9"><%=(rs_lista_pi.Fields.Item("eng_obras_construtora").Value)%></span></div></td>
        <td><div align="left"><span class="style9"><%=(rs_lista_pi.Fields.Item("desc_situacao_interna").Value)%></span></div></td>
        <td><div align="left"><span class="style9"><%=(rs_lista_pi.Fields.Item("desc_situacao_externa").Value)%></span></div></td>
        <td><div align="left"><span class="style9"><%=(rs_lista_pi.Fields.Item("dsc_observacoes_relatorio_mensal").Value)%></span></div></td>
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
