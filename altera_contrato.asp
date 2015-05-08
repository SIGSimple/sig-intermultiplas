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
  MM_editTable = "tb_contrato"
  MM_editColumn = "id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "cad_contrato.asp"
  MM_fieldsStr  = "cod_projetista|value|cod_empresa_contratada|value|cod_engenheiro_empresa_contratada|value|num_autos|value|num_contrato|value|dta_assinatura|value|dta_publicacao_doe|value|dta_pedido_empenho|value|dta_os|value|prz_original_contrato_meses|value|prz_original_execucao_meses|value|dta_vigencia|value|dta_base|value|dta_inauguracao|value|dta_termo_recebimento_provisorio|value|dta_termo_recebimento_definitivo|value|dta_encerramento_contrato|value|dta_recisao_contratual|value|cod_contratante|value|cod_situacao|value"
  MM_columnsStr = "cod_projetista|none,none,NULL|cod_empresa_contratada|none,none,NULL|cod_engenheiro_empresa_contratada|none,none,NULL|num_autos|',none,''|num_contrato|',none,''|dta_assinatura|',none,NULL|dta_publicacao_doe|',none,NULL|dta_pedido_empenho|',none,NULL|dta_os|',none,NULL|prz_original_contrato_meses|',none,''|prz_original_execucao_meses|',none,''|dta_vigencia|',none,NULL|dta_base|',none,NULL|dta_inauguracao|',none,NULL|dta_termo_recebimento_provisorio|',none,NULL|dta_termo_recebimento_definitivo|',none,NULL|dta_encerramento_contrato|',none,NULL|dta_recisao_contratual|',none,NULL|cod_contratante|none,none,NULL|cod_situacao|none,none,NULL"

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
rs.Source = "SELECT * FROM tb_contrato WHERE id = " + Replace(rs__MMColParam, "'", "''")
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

        $("input[name=prz_original_contrato_meses]").on("blur", function() {
          dta_assinatura = moment($("input[name=dta_assinatura]").val(), "DD/MM/YYYY");
          prz_meses = $("input[name=prz_original_contrato_meses]").val()
          $("input[name=dta_vigencia]").val(dta_assinatura.add(prz_meses, "M").format("DD/MM/YYYY"));
        });
      });
    </script>
</head>

<body>
<p align="center"><strong><span class="style17">Alteração de Contrato </span></strong></p>
    <form method="post" action="<%=MM_editAction%>" name="form1">
      <input type="hidden" name="MM_update" value="form1">
      <input type="hidden" name="MM_recordId" value="<%= rs.Fields.Item("id").Value %>">
      <table align="center">
        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Empresa Contratada:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <select name="cod_empresa_contratada">
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
                       If Lcase(rs_combo.Fields.Item("cod_construtora").Value) = Lcase(rs.Fields.Item("cod_empresa_contratada").Value) then
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
            <span class="style22">Engenheiro Responsável:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <select name="cod_engenheiro_empresa_contratada">
              <option value=""></option>
              <%
                strQ = "SELECT * FROM c_lista_fiscal "

                Set rs_combo = Server.CreateObject("ADODB.Recordset")
                  rs_combo.CursorLocation = 3
                  rs_combo.CursorType = 3
                  rs_combo.LockType = 1
                  rs_combo.Open strQ, objCon, , , &H0001

                If Not rs_combo.EOF Then
                  While Not rs_combo.EOF
                    If Trim(rs_combo.Fields.Item("nme_interessado").Value) <> "" Then
                       Response.Write "      <OPTION value='" & (rs_combo.Fields.Item("cod_fiscal").Value) & "'"
                       If Lcase(rs_combo.Fields.Item("cod_fiscal").Value) = Lcase(rs.Fields.Item("cod_engenheiro_empresa_contratada").Value) then
                         Response.Write "selected"
                       End If
                       Response.Write ">" & (rs_combo.Fields.Item("nme_interessado").Value) & "</OPTION>"
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
            <span class="style22">Licitação:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <select name="cod_licitacao">
              <option value=""></option>
              <%
                strQ = "SELECT * FROM tb_licitacao ORDER BY num_autos ASC"

                Set rs_combo = Server.CreateObject("ADODB.Recordset")
                  rs_combo.CursorLocation = 3
                  rs_combo.CursorType = 3
                  rs_combo.LockType = 1
                  rs_combo.Open strQ, objCon, , , &H0001

                If Not rs_combo.EOF Then
                  While Not rs_combo.EOF
                       Response.Write "      <OPTION value='" & (rs_combo.Fields.Item("id").Value) & "'"
                       If Lcase(rs_combo.Fields.Item("id").Value) = Lcase(rs.Fields.Item("cod_licitacao").Value) then
                         Response.Write "selected"
                       End If
                       Response.Write ">Nº Autos: " & (rs_combo.Fields.Item("num_autos").Value) & " - Nº Edital: " & (rs_combo.Fields.Item("num_autos").Value) & "</OPTION>"
                    rs_combo.MoveNext
                  Wend
                End If
              %>
            </select>
          </td>
        </tr>

        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Nº Autos:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" name="num_autos" value="<%=(rs.Fields.Item("num_autos").Value)%>" size="32">
          </td>
        </tr>

        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Nº Contrato:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" name="num_contrato" value="<%=(rs.Fields.Item("num_contrato").Value)%>" size="32">
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
            <span class="style22">Data Pedido Empenho:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" class="datepicker" name="dta_pedido_empenho" value="<%=(rs.Fields.Item("dta_pedido_empenho").Value)%>" size="32">
          </td>
        </tr>

        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Data Ordem de Serviço:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" class="datepicker" name="dta_os" value="<%=(rs.Fields.Item("dta_os").Value)%>" size="32">
          </td>
        </tr>

        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Prazo Original (Meses):</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" name="prz_original_contrato_meses" value="<%=(rs.Fields.Item("prz_original_contrato_meses").Value)%>" size="32">
          </td>
        </tr>

        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Prazo de Execução (Meses):</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" name="prz_original_execucao_meses" value="<%=(rs.Fields.Item("prz_original_execucao_meses").Value)%>" size="32">
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
            <span class="style22">Data Base:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" class="datepicker" name="dta_base" value="<%=(rs.Fields.Item("dta_base").Value)%>" size="32">
          </td>
        </tr>

        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Data Inauguração:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" class="datepicker" name="dta_inauguracao" value="<%=(rs.Fields.Item("dta_inauguracao").Value)%>" size="32">
          </td>
        </tr>

        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Data Rec. Termo Provisório:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" class="datepicker" name="dta_termo_recebimento_provisorio" value="<%=(rs.Fields.Item("dta_termo_recebimento_provisorio").Value)%>" size="32">
          </td>
        </tr>

        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Data Rec. Termo Definitivo:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" class="datepicker" name="dta_termo_recebimento_definitivo" value="<%=(rs.Fields.Item("dta_termo_recebimento_definitivo").Value)%>" size="32">
          </td>
        </tr>

        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Data Encerramento Contrato:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" class="datepicker" name="dta_encerramento_contrato" value="<%=(rs.Fields.Item("dta_encerramento_contrato").Value)%>" size="32">
          </td>
        </tr>

        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Data Recisão Contratual:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <input type="text" class="datepicker" name="dta_recisao_contratual" value="<%=(rs.Fields.Item("dta_recisao_contratual").Value)%>" size="32">
          </td>
        </tr>

        <tr valign="baseline">
          <td align="right" nowrap bgcolor="#CCCCCC" class="style7">
            <span class="style22">Contratante:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <select name="cod_contratante">
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
                       If Lcase(rs_combo.Fields.Item("cod_construtora").Value) = Lcase(rs.Fields.Item("cod_contratante").Value) then
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
            <span class="style22">Situação:</span>
          </td>
          <td bgcolor="#CCCCCC">
            <select name="cod_situacao">
              <option value=""></option>
              <%
                strQ = "SELECT * FROM tb_situacao_pi ORDER BY desc_situacao ASC "

                Set rs_combo = Server.CreateObject("ADODB.Recordset")
                  rs_combo.CursorLocation = 3
                  rs_combo.CursorType = 3
                  rs_combo.LockType = 1
                  rs_combo.Open strQ, objCon, , , &H0001

                If Not rs_combo.EOF Then
                  While Not rs_combo.EOF
                    If Trim(rs_combo.Fields.Item("desc_situacao").Value) <> "" Then
                       Response.Write "      <OPTION value='" & (rs_combo.Fields.Item("cod_situacao").Value) & "'"
                       If Lcase(rs_combo.Fields.Item("cod_situacao").Value) = Lcase(rs.Fields.Item("cod_situacao").Value) then
                         Response.Write "selected"
                       End If
                       Response.Write ">" & (rs_combo.Fields.Item("desc_situacao").Value) & "</OPTION>"
                    End If
                    rs_combo.MoveNext
                  Wend
                End If
              %>
            </select>
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