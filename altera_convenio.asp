<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
' *** Edit Operations: declare variables

Response.CharSet = "UTF-8"
  
Dim objCon
Set objCon = Server.CreateObject("ADODB.Connection")
    objCon.Open MM_cpf_STRING

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
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_cpf_STRING
  MM_editTable = "tb_convenio"
  MM_editColumn = "id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "cad_convenio.asp"
  MM_fieldsStr  = "num_autos|value|cod_projetista_convenio|value|cod_enquadramento|value|cod_programa|value|num_convenio|value|dta_assinatura|value|dta_publicacao_doe|value|vlr_original|value|prz_meses|value|dta_vigencia|value|nme_fonte_recurso|value|cod_coordenador_daee|value|dsc_observacoes|value"
  MM_columnsStr = "num_autos|',none,''|cod_projetista_convenio|none,none,NULL|cod_enquadramento|none,none,NULL|cod_programa|none,none,NULL|num_convenio|',none,''|dta_assinatura|',none,NULL|dta_publicacao_doe|',none,NULL|vlr_original|',none,NULL|prz_meses|none,none,NULL|dta_vigencia|',none,NULL|nme_fonte_recurso|',none,''|cod_coordenador_daee|none,none,NULL|dsc_observacoes|',none,''"

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
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
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
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
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
Dim rs__MMColParam
rs__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  rs__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim rs
Dim rs_numRows

Set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_cpf_STRING
rs.Source = "SELECT * FROM tb_convenio WHERE id = " + Replace(rs__MMColParam, "'", "''")
rs.CursorType = 0
rs.CursorLocation = 2
rs.LockType = 1
rs.Open()

rs_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style17 {	font-family: Arial, Helvetica, sans-serif;
	font-size: 16px;
}
.style44 {font-family: Arial, Helvetica, sans-serif; font-size: 9; font-weight: bold; }
.style45 {font-size: 9}
.style22 {font-family: Arial, Helvetica, sans-serif; font-size: 9; }
.style7 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
-->
</style>
<link rel="stylesheet" href="//code.jquery.com/ui/1.11.3/themes/smoothness/jquery-ui.css">
    <script src="//code.jquery.com/jquery-1.11.2.min.js"></script>
    <script type="text/javascript" src="//code.jquery.com/ui/1.11.4/jquery-ui.min.js"></script>
    <script type="text/javascript" src="js/datepicker-pt-BR.js"></script>
    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.9.0/moment.min.js"></script>
    <script type="text/javascript">
      $(document).ready(function() {
        $(".datepicker").datepicker($.datepicker.regional["pt-BR"]);

        $("input[name=prz_meses]").on("blur", function() {
          dta_assinatura = moment($("input[name=dta_assinatura]").val(), "DD/MM/YYYY");
          prz_meses = $("input[name=prz_meses]").val()
          $("input[name=dta_vigencia]").val(dta_assinatura.add(prz_meses, "M").format("DD/MM/YYYY"));
        });
      });
    </script>
</head>

<body>
<p align="center"><strong><span class="style17">Alteração de Convênio </span></strong></p>
    <form method="post" action="<%=MM_editAction%>" name="form1">
      <input type="hidden" name="MM_update" value="form1">
      <input type="hidden" name="MM_recordId" value="<%= rs.Fields.Item("id").Value %>">
      <table align="center">
        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Autos:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" name="num_autos" value="<%=(rs.Fields.Item("num_autos").Value)%>" size="32">
          </td>
        </tr>
        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Projetista:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <select name="cod_projetista_convenio">
              <option value=""></option>
              <%
                strQ = "SELECT * FROM tb_Construtora ORDER BY Construtora ASC "

                Set rs_combo = Server.CreateObject("ADODB.Recordset")
                  rs_combo.CursorLocation = 3
                  rs_combo.CursorType = 3
                  rs_combo.LockType = 1
                  rs_combo.Open strQ, objCon, , , &H0001

                If Not rs_combo.EOF Then
                  While Not rs_combo.EOF
                    If Trim(rs_combo.Fields.Item("Construtora").Value) <> "" Then
                      Response.Write "      <OPTION value='" & (rs_combo.Fields.Item("cod_construtora").Value) & "'"
                      If Lcase(rs_combo.Fields.Item("cod_construtora").Value) = Lcase(rs.Fields.Item("cod_projetista_convenio").Value) then
                        Response.Write "selected"
                      End If
                      Response.Write ">" & (rs_combo.Fields.Item("Construtora").Value) & "</OPTION>"
                    End If
                    rs_combo.MoveNext
                  Wend
                End If
              %>
            </select>
          </td>
        </tr>
        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Enquadramento:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <select name="cod_enquadramento">
              <option value=""></option>
              <%
                strQ = "SELECT * FROM tb_convenio_enquadramento "

                Set rs_combo = Server.CreateObject("ADODB.Recordset")
                  rs_combo.CursorLocation = 3
                  rs_combo.CursorType = 3
                  rs_combo.LockType = 1
                  rs_combo.Open strQ, objCon, , , &H0001

                If Not rs_combo.EOF Then
                  While Not rs_combo.EOF
                    If Trim(rs_combo.Fields.Item("dsc_enquadramento").Value) <> "" Then
                       Response.Write "      <OPTION value='" & (rs_combo.Fields.Item("id").Value) & "'"
                       If Lcase(rs_combo.Fields.Item("id").Value) = Lcase(rs.Fields.Item("cod_enquadramento").Value) then
                         Response.Write "selected"
                       End If
                       Response.Write ">" & (rs_combo.Fields.Item("dsc_enquadramento").Value) & "</OPTION>"
                    End If
                    rs_combo.MoveNext
                  Wend
                End If
              %>
            </select>
          </td>
        </tr>
        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Programa:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <select name="cod_programa">
              <option value=""></option>
              <%
                strQ = "SELECT * FROM tb_depto "

                Set rs_combo = Server.CreateObject("ADODB.Recordset")
                  rs_combo.CursorLocation = 3
                  rs_combo.CursorType = 3
                  rs_combo.LockType = 1
                  rs_combo.Open strQ, objCon, , , &H0001

                If Not rs_combo.EOF Then
                  While Not rs_combo.EOF
                    If Trim(rs_combo.Fields.Item("desc_depto").Value) <> "" Then
                      Response.Write "      <OPTION value='" & (rs_combo.Fields.Item("cod_depto").Value) & "'"
                      If Lcase(rs_combo.Fields.Item("cod_depto").Value) = Lcase(rs.Fields.Item("cod_programa").Value) then
                        Response.Write "selected"
                      End If
                      Response.Write ">" & (rs_combo.Fields.Item("sigla").Value) & " - " & (rs_combo.Fields.Item("desc_depto").Value) & "</OPTION>"
                    End If
                    rs_combo.MoveNext
                  Wend
                End If
              %>
            </select>
          </td>
        </tr>
        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Núm. Convênio:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" name="num_convenio" value="<%=(rs.Fields.Item("num_convenio").Value)%>" size="32">
          </td>
        </tr>
        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Data Assinatura:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" class="datepicker" name="dta_assinatura" value="<%=(rs.Fields.Item("dta_assinatura").Value)%>" size="32">
          </td>
        </tr>
        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Data Publicação DOE:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" class="datepicker" name="dta_publicacao_doe" value="<%=(rs.Fields.Item("dta_publicacao_doe").Value)%>" size="32">
          </td>
        </tr>
        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Valor Original:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" name="vlr_original" value="<%=(rs.Fields.Item("vlr_original").Value)%>" size="32">
          </td>
        </tr>
        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Prazo (Meses):</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" name="prz_meses" value="<%=(rs.Fields.Item("prz_meses").Value)%>" size="32">
          </td>
        </tr>
        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Vigência até:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" class="datepicker" name="dta_vigencia" value="<%=(rs.Fields.Item("dta_vigencia").Value)%>" size="32">
          </td>
        </tr>
        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Fonte de Recurso:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" name="nme_fonte_recurso" value="<%=(rs.Fields.Item("nme_fonte_recurso").Value)%>" size="32">
          </td>
        </tr>
        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Coord. DAEE:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <select name="cod_coordenador_daee">
              <option value=""></option>
              <%
                strQ = "SELECT * FROM tb_responsavel "

                Set rs_combo = Server.CreateObject("ADODB.Recordset")
                  rs_combo.CursorLocation = 3
                  rs_combo.CursorType = 3
                  rs_combo.LockType = 1
                  rs_combo.Open strQ, objCon, , , &H0001

                If Not rs_combo.EOF Then
                  While Not rs_combo.EOF
                    If Trim(rs_combo.Fields.Item("Responsável").Value) <> "" Then
                      Response.Write "      <OPTION value='" & (rs_combo.Fields.Item("cod_fiscal").Value) & "'"
                      If Lcase(rs_combo.Fields.Item("cod_fiscal").Value) = Lcase(rs.Fields.Item("cod_coordenador_daee").Value) then
                        Response.Write "selected"
                      End If
                      Response.Write ">" & (rs_combo.Fields.Item("Responsável").Value) & "</OPTION>"
                    End If
                    rs_combo.MoveNext
                  Wend
                End If
              %>
            </select>
          </td>
        </tr>
        <tr valign="middle">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Observações:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <textarea name="dsc_observacoes" cols="25" style="width: 98%;"><%=(rs.Fields.Item("dsc_observacoes").Value)%></textarea>
          </td>
        </tr>
        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">&nbsp;</td>
          <td bgcolor="#CCCCCC">
            <input type="submit" value="Salvar">
          </td>
        </tr>
      </table>
    </form>
<p>&nbsp;</p>
</body>
</html>
<%
rs.Close()
Set rs = Nothing
%>