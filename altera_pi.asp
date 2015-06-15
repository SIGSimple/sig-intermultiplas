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
  MM_editTable = "tb_pi"
  MM_editColumn = "PI"
  MM_recordId = "'" + Request.Form("MM_recordId") + "'"
  MM_editRedirectUrl = ""
  MM_fieldsStr  = "PI|value"
  MM_columnsStr = "PI|',none,''"

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

  Dim frmField_PI
  Dim frmField_cod_predio
  Dim frmField_id_predio
  Dim frmField_municipio
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
  Dim frmField_valor_contrato
  Dim frmField_dta_inicio_obras
  Dim frmField_num_percentual_executado
  Dim frmField_dta_previsao_termino
  Dim frmField_dta_inauguracao
  Dim frmField_dta_previsao_inauguracao
  Dim frmField_cod_fiscal
  Dim frmField_cod_engenheiro_daee
  Dim frmField_cod_engenheiro_plan_consorcio
  Dim frmField_cod_fiscal_consorcio
  Dim frmField_cod_engenheiro_construtora
  Dim frmField_cod_situacao
  Dim frmField_cod_situacao_externa
  Dim frmField_dsc_observacoes_relatorio_mensal

  Dim frmField_cod_bacia_hidrografica
  Dim frmField_cod_manancial_lancamento
  Dim frmField_qtd_metragem_coletor_tronco
  Dim frmField_qtd_metragem_interceptor
  Dim frmField_qtd_metragem_emissario_fluente_bruto
  Dim frmField_qtd_eee
  Dim frmField_qtd_metragem_linha_recalque
  Dim frmField_cod_tipo_ete
  Dim frmField_dsc_estacao_tratamento
  Dim frmField_qtd_metragem_emissario_efluente_tratado
  Dim frmField_flg_estudo_elaborado_daee
  Dim frmField_dsc_observacoes
  Dim frmField_cod_parceiro
  Dim frmField_dsc_parceria_realizacao
  Dim frmField_dsc_observacoes_gestor

  Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
      IIf = sTrue
    Else
      IIf = sFalse
    End If
  End Function

  frmField_cod_tipo_empreendimento                 = IIf(Request.Form("cod_tipo_empreendimento") = "", "NULL", Request.Form("cod_tipo_empreendimento"))
  frmField_cod_programa                            = IIf(Request.Form("cod_programa") = "", "NULL", Request.Form("cod_programa"))
  'frmField_dsc_situacao_anterior                  = IIf(Request.Form("dsc_situacao_anterior") = "", "", Request.Form("dsc_situacao_anterior"))
  'frmField_dsc_situacao_atual                     = IIf(Request.Form("dsc_situacao_atual") = "", "", Request.Form("dsc_situacao_atual"))
  frmField_dsc_resultado_obtido                    = IIf(Request.Form("dsc_resultado_obtido") = "", "", Request.Form("dsc_resultado_obtido"))
  frmField_endereco                                = IIf(Request.Form("endereco") = "", "", Request.Form("endereco"))
  frmField_cep                                     = IIf(Request.Form("cep") = "", "", Request.Form("cep"))
  frmField_latitude_longitude                      = IIf(Request.Form("latitude_longitude") = "", "", Request.Form("latitude_longitude"))
  frmField_email                                   = IIf(Request.Form("email") = "", "", Request.Form("email"))
  frmField_telefone                                = IIf(Request.Form("telefone") = "", "", Request.Form("telefone"))
  frmField_qtd_populacao_urbana_2010               = IIf(Request.Form("qtd_populacao_urbana_2010") = "", "NULL", Request.Form("qtd_populacao_urbana_2010"))
  'frmField_qtd_populacao_urbana_2030              = IIf(Request.Form("qtd_populacao_urbana_2030") = "", "NULL", Request.Form("qtd_populacao_urbana_2030"))
  frmField_valor_contrato                          = IIf(Request.Form("ValorContrato") = "", "0", Replace(Request.Form("ValorContrato"), ",","."))
  frmField_dta_inicio_obras                        = IIf(Request.Form("dta_inicio_obras") = "", "NULL", "'"& Request.Form("dta_inicio_obras") &"'")
  frmField_num_percentual_executado                = IIf(Request.Form("num_percentual_executado") = "", "0", Replace(Request.Form("num_percentual_executado"), ",","."))
  frmField_dta_previsao_termino                    = IIf(Request.Form("dta_previsao_termino") = "", "NULL", "'"& Request.Form("dta_previsao_termino") &"'")
  frmField_dta_inauguracao                         = IIf(Request.Form("dta_inauguracao") = "", "NULL", "'"& Request.Form("dta_inauguracao") &"'")
  frmField_dta_previsao_inauguracao                = IIf(Request.Form("dta_previsao_inauguracao") = "", "NULL", "'"& Request.Form("dta_previsao_inauguracao") &"'")
  frmField_cod_fiscal                              = IIf(Request.Form("cod_fiscal") = "", "NULL", Request.Form("cod_fiscal"))
  frmField_cod_engenheiro_daee                     = IIf(Request.Form("cod_engenheiro_daee") = "", "NULL", Request.Form("cod_engenheiro_daee"))
  frmField_cod_engenheiro_plan_consorcio           = IIf(Request.Form("cod_engenheiro_plan_consorcio") = "", "NULL", Request.Form("cod_engenheiro_plan_consorcio"))
  frmField_cod_fiscal_consorcio                    = IIf(Request.Form("cod_fiscal_consorcio") = "", "NULL", Request.Form("cod_fiscal_consorcio"))
  frmField_cod_engenheiro_medicao                  = IIf(Request.Form("cod_engenheiro_medicao") = "", "NULL", Request.Form("cod_engenheiro_medicao"))
  frmField_cod_engenheiro_construtora              = IIf(Request.Form("cod_engenheiro_construtora") = "", "NULL", Request.Form("cod_engenheiro_construtora"))
  frmField_cod_situacao                            = IIf(Request.Form("cod_situacao") = "", "NULL", Request.Form("cod_situacao"))
  frmField_cod_situacao_externa                    = IIf(Request.Form("cod_situacao_externa") = "", "NULL", Request.Form("cod_situacao_externa"))
  frmField_dsc_observacoes_relatorio_mensal        = IIf(Request.Form("dsc_observacoes_relatorio_mensal") = "", "", Request.Form("dsc_observacoes_relatorio_mensal"))
  frmField_cod_bacia_hidrografica                  = IIF(Request.Form("cod_bacia_hidrografica") = "", "NULL", Request.Form("cod_bacia_hidrografica"))
  frmField_cod_manancial_lancamento                = IIF(Request.Form("cod_manancial_lancamento") = "", "NULL", Request.Form("cod_manancial_lancamento"))
  frmField_qtd_metragem_coletor_tronco             = IIF(Request.Form("qtd_metragem_coletor_tronco") = "", "0", Request.Form("qtd_metragem_coletor_tronco"))
  frmField_qtd_metragem_interceptor                = IIF(Request.Form("qtd_metragem_interceptor") = "", "0", Request.Form("qtd_metragem_interceptor"))
  frmField_qtd_metragem_emissario_fluente_bruto    = IIF(Request.Form("qtd_metragem_emissario_fluente_bruto") = "", "0", Request.Form("qtd_metragem_emissario_fluente_bruto"))
  frmField_qtd_eee                                 = IIF(Request.Form("qtd_eee") = "", "0", Request.Form("qtd_eee"))
  frmField_qtd_metragem_linha_recalque             = IIF(Request.Form("qtd_metragem_linha_recalque") = "", "0", Request.Form("qtd_metragem_linha_recalque"))
  frmField_cod_tipo_ete                            = IIF(Request.Form("cod_tipo_ete") = "", "NULL", Request.Form("cod_tipo_ete"))
  frmField_dsc_estacao_tratamento                  = IIF(Request.Form("dsc_estacao_tratamento") = "", "", Request.Form("dsc_estacao_tratamento"))
  frmField_qtd_metragem_emissario_efluente_tratado = IIF(Request.Form("qtd_metragem_emissario_efluente_tratado") = "", "0", Request.Form("qtd_metragem_emissario_efluente_tratado"))
  frmField_flg_estudo_elaborado_daee               = IIF(Request.Form("flg_estudo_elaborado_daee") = "", "", Request.Form("flg_estudo_elaborado_daee"))
  frmField_dsc_observacoes                         = IIF(Request.Form("dsc_observacoes") = "", "", Request.Form("dsc_observacoes"))
  frmField_cod_parceiro                            = IIF(Request.Form("cod_parceiro") = "", "NULL", Request.Form("cod_parceiro"))
  frmField_dsc_parceria_realizacao                 = IIF(Request.Form("dsc_parceria_realizacao") = "", "", Request.Form("dsc_parceria_realizacao"))
  frmField_dsc_observacoes_gestor                  = IIF(Request.Form("dsc_observacoes_gestor") = "", "", Request.Form("dsc_observacoes_gestor"))

  MM_editQuery = "UPDATE " & MM_editTable & " SET "
  MM_editQuery = MM_editQuery & "cod_tipo_empreendimento="& frmField_cod_tipo_empreendimento &",cod_programa="& frmField_cod_programa
  'MM_editQuery = MM_editQuery & "dsc_situacao_anterior='"& frmField_dsc_situacao_anterior &"',dsc_situacao_atual='"& frmField_dsc_situacao_atual & "'"
  MM_editQuery = MM_editQuery & ",dsc_resultado_obtido='"& frmField_dsc_resultado_obtido & "',endereco='"& frmField_endereco &"',cep='"& frmField_cep & "'"
  MM_editQuery = MM_editQuery & ",latitude_longitude='"& frmField_latitude_longitude &"'"
  MM_editQuery = MM_editQuery & ",email='"& frmField_email &"',telefone='"& frmField_telfone &"',qtd_populacao_urbana_2010="& frmField_qtd_populacao_urbana_2010 &",[Valor do Contrato]="& frmField_valor_contrato &",dta_inicio_obras="& frmField_dta_inicio_obras &",num_percentual_executado="& frmField_num_percentual_executado &",dta_previsao_termino="& frmField_dta_previsao_termino & ",dta_inauguracao=" & frmField_dta_inauguracao & ",dta_previsao_inauguracao=" & frmField_dta_previsao_inauguracao
    '&",qtd_populacao_urbana_2030="& frmField_qtd_populacao_urbana_2030
  MM_editQuery = MM_editQuery & ",cod_fiscal="& frmField_cod_fiscal &",cod_engenheiro_daee="& frmField_cod_engenheiro_daee &",cod_engenheiro_plan_consorcio="& frmField_cod_engenheiro_plan_consorcio &",cod_fiscal_consorcio="& frmField_cod_fiscal_consorcio &",cod_engenheiro_medicao="& frmField_cod_engenheiro_medicao &",cod_engenheiro_construtora="& frmField_cod_engenheiro_construtora
  MM_editQuery = MM_editQuery & ",cod_situacao="& frmField_cod_situacao
  MM_editQuery = MM_editQuery & ",cod_situacao_externa="& frmField_cod_situacao_externa & ",dsc_observacoes_relatorio_mensal='"& frmField_dsc_observacoes_relatorio_mensal &"'"
  MM_editQuery = MM_editQuery & ",cod_bacia_hidrografica="& frmField_cod_bacia_hidrografica &",cod_manancial_lancamento="& frmField_cod_manancial_lancamento &",qtd_metragem_coletor_tronco="& frmField_qtd_metragem_coletor_tronco &",qtd_metragem_interceptor="& frmField_qtd_metragem_interceptor &",qtd_metragem_emissario_fluente_bruto="& frmField_qtd_metragem_emissario_fluente_bruto &",qtd_eee="& frmField_qtd_eee &",qtd_metragem_linha_recalque="& frmField_qtd_metragem_linha_recalque &",cod_tipo_ete="& frmField_cod_tipo_ete &",dsc_estacao_tratamento='"& frmField_dsc_estacao_tratamento &"',qtd_metragem_emissario_efluente_tratado="& frmField_qtd_metragem_emissario_efluente_tratado &",flg_estudo_elaborado_daee='"& frmField_flg_estudo_elaborado_daee &"',dsc_observacoes_obra='"& frmField_dsc_observacoes &"',cod_parceiro="& frmField_cod_parceiro &"" &",dsc_parceria_realizacao='"& frmField_dsc_parceria_realizacao &"',dsc_observacoes_gestor='"& frmField_dsc_observacoes_gestor & "'"
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    'Response.Write MM_editQuery
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
Dim rs_pi__MMColParam
rs_pi__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  rs_pi__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim rs_pi
Dim rs_pi_numRows

Set rs_pi = Server.CreateObject("ADODB.Recordset")
rs_pi.ActiveConnection = MM_cpf_STRING
rs_pi.Source = "SELECT tb_pi.*, tb_predio.[Diretoria de Ensino] AS bacia_daee FROM tb_pi LEFT JOIN tb_predio ON (tb_pi.cod_predio = tb_predio.cod_predio) AND (tb_pi.id_predio = tb_predio.id_predio) WHERE PI = '" + Replace(rs_pi__MMColParam, "'", "''") + "'"
rs_pi.CursorType = 0
rs_pi.CursorLocation = 2
rs_pi.LockType = 1
rs_pi.Open()

rs_pi_numRows = 0
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

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style7 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; color: #FFFFFF; }
.style9 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; }
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
<link rel="stylesheet" href="//code.jquery.com/ui/1.11.3/themes/smoothness/jquery-ui.css">
<script src="//code.jquery.com/jquery-1.11.2.min.js"></script>
<script type="text/javascript" src="//code.jquery.com/ui/1.11.4/jquery-ui.min.js"></script>
<script type="text/javascript" src="js/datepicker-pt-BR.js"></script>
<script type="text/javascript">
  $(function() {
    $(".datepicker").datepicker($.datepicker.regional["pt-BR"]);
  });
</script>
</head>

<body>
  <p align="center">
    <strong>
      <span class="style11">Atualização de Dados de Empreendimentos</span>
    </strong>
  </p>

  <form method="post" action="<%=MM_editAction%>" name="form1">
    <input type="hidden" name="MM_update" value="form1">
    <input type="hidden" name="MM_recordId" value="<%= rs_pi.Fields.Item("PI").Value %>">

    <table align="center">
      <!-- MUNICÍPIO -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10">Selecione o Município:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <select name="cod_predio" class="style9" <% If Session("MM_UserAuthorization") <> 1 Then Response.Write "disabled='disabled'" End If %>>
            <option value=""></option>
            <%
              While (NOT rs_predio.EOF)
                If Trim(rs_predio.Fields.Item("Expr1").Value) <> "" Then
                  Response.Write "      <OPTION value='" & (rs_predio.Fields.Item("id_predio").Value) & "-"& (rs_predio.Fields.Item("cod_predio").Value) &"-"& (rs_predio.Fields.Item("Expr1").Value) &"'"
                  If Lcase(rs_predio.Fields.Item("cod_predio").Value) = Lcase(rs_pi.Fields.Item("cod_predio").Value) then
                    Response.Write "selected"
                  End If
                  Response.Write ">" & (rs_predio.Fields.Item("Expr1").Value) & "</OPTION>"
                End If
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
          <input name="nome_empreendimento" type="text" class="style9" value="<%=(rs_pi.Fields.Item("nome_empreendimento").Value)%>" <% If Session("MM_UserAuthorization") <> 1 Then Response.Write "disabled='disabled'" End If %>>
        </td>
      </tr>

      <!-- AUTOS -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10">Autos:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <input name="PI" type="text" class="style9" value="<%=(rs_pi.Fields.Item("PI").Value)%>" size="18" <% If Session("MM_UserAuthorization") <> 1 Then Response.Write "disabled='disabled'" End If %>>
        </td>
      </tr>

      <!-- BACIA DAEE -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10">Bacia DAEE:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <input name="PI" type="text" class="style9" value="<%=(rs_pi.Fields.Item("bacia_daee").Value)%>" style="width: 98%;" disabled="disabled">
        </td>
      </tr>

      <!-- DESCRIÇÃO DO EMPREENDIMENTO -->
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10">Objeto da Obra:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <textarea name="descricao_empreendimento" cols="50" rows="5" class="style9" style="width: 98%;" <% If Session("MM_UserAuthorization") <> 1 Then Response.Write "disabled='disabled'" End If %>><%=(rs_pi.Fields.Item("Descrição da Intervenção FDE").Value)%></textarea>
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
                If Trim(rs_tipo_empreendimento.Fields.Item("desc_tipo").Value) <> "" Then
                  Response.Write "      <OPTION value='" & (rs_tipo_empreendimento.Fields.Item("id").Value) & "'"
                  If Lcase(rs_tipo_empreendimento.Fields.Item("id").Value) = Lcase(rs_pi.Fields.Item("cod_tipo_empreendimento").Value) then
                    Response.Write "selected"
                  End If
                  Response.Write ">" & (rs_tipo_empreendimento.Fields.Item("desc_tipo").Value) & "</OPTION>"
                End If
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
                If Trim(rs_programa.Fields.Item("desc_depto").Value) <> "" Then
                  Response.Write "      <OPTION value='" & (rs_programa.Fields.Item("cod_depto").Value) & "'"
                  If Lcase(rs_programa.Fields.Item("cod_depto").Value) = Lcase(rs_pi.Fields.Item("cod_programa").Value) then
                    Response.Write "selected"
                  End If
                  Response.Write ">" & (rs_programa.Fields.Item("sigla").Value) &"-"& (rs_programa.Fields.Item("desc_depto").Value) & "</OPTION>"
                End If
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

      <!-- SITUAÇÃO ANTERIOR -->
      <!-- <tr valign="baseline">
        <td align="right" valign="middle" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Situação Anterior:</span></td>
        <td bgcolor="#CCCCCC">
          <textarea name="dsc_situacao_anterior"cols="50" rows="5" class="style9" style="width: 98%;"><%=(rs_pi.Fields.Item("dsc_situacao_anterior").Value)%></textarea>
        </td>
      </tr> -->

      <!-- SITUAÇÃO ATUAL -->
      <!-- <tr valign="baseline">
        <td align="right" valign="middle" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Situação Atual:</span></td>
        <td bgcolor="#CCCCCC">
          <textarea name="dsc_situacao_atual"cols="50" rows="5" class="style9" style="width: 98%;"><%=(rs_pi.Fields.Item("dsc_situacao_atual").Value)%></textarea>
        </td>
      </tr> -->

      <!-- BENEFÍCIO -->
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Beneficio Geral da Obra:</span></td>
        <td bgcolor="#CCCCCC">
          <textarea name="dsc_resultado_obtido"cols="50" rows="5" class="style9" style="width: 98%;"><%=(rs_pi.Fields.Item("dsc_resultado_obtido").Value)%></textarea>
        </td>
      </tr>

      <!-- ENDEREÇO -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Endereço:</span></td>
        <td bgcolor="#CCCCCC">
          <input name="endereco" type="text" class="style9" value="<%=(rs_pi.Fields.Item("endereco").Value)%>">
        </td>
      </tr>

      <!-- CEP -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">CEP:</span></td>
        <td bgcolor="#CCCCCC">
          <input name="cep" type="text" class="style9" value="<%=(rs_pi.Fields.Item("cep").Value)%>">
        </td>
      </tr>

      <!-- LOCALIZAÇÃO GEOGRÁFICA (LAT,LONG) -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Localização Geográfica (Lat, Long):</span></td>
        <td bgcolor="#CCCCCC">
          <input name="latitude_longitude" type="text" class="style9" value="<%=(rs_pi.Fields.Item("latitude_longitude").Value)%>">
        </td>
      </tr>

      <!-- EMAIL -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Email:</span></td>
        <td bgcolor="#CCCCCC">
          <input name="email" type="text" class="style9" value="<%=(rs_pi.Fields.Item("email").Value)%>">
        </td>
      </tr>

      <!-- TELEFONE -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Telefone:</span></td>
        <td bgcolor="#CCCCCC">
          <input name="telefone" type="text" class="style9" value="<%=(rs_pi.Fields.Item("telefone").Value)%>">
        </td>
      </tr>

      <!-- POPULAÇÃO 2010 -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">População Urbana - IBGE (2010) (hab):</span></td>
        <td bgcolor="#CCCCCC">
          <input name="qtd_populacao_urbana_2010" type="text" class="style9" value="<%=(rs_pi.Fields.Item("qtd_populacao_urbana_2010").Value)%>">
        </td>
      </tr>

      <!-- PROJEÇÃO POPULAÇÃO 2030 -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Projeção de População (2030):</span></td>
        <td bgcolor="#CCCCCC">
          <%
            If Not IsNull(rs_pi.Fields.Item("qtd_populacao_urbana_2010").Value) then
              projecao_inicial = rs_pi.Fields.Item("qtd_populacao_urbana_2010").Value
              projecao_inicial = projecao_inicial * 1.25
              a = Round(projecao_inicial/100, 0)
              b = a * 100
            End If
          %>
          <input name="qtd_populacao_urbana_2030" type="text" class="style9" value="<%=(b)%>" disabled="disabled">
        </td>
      </tr>

      <!-- INVESTIMENTO GOVERNO SP -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Investimento Governo SP:</span></td>
        <td bgcolor="#CCCCCC">
          <input name="ValorContrato" type="text" class="style9" value="<%=(rs_pi.Fields.Item("Valor do Contrato").Value)%>">
        </td>
      </tr>

      <!-- INÍCIO DAS OBRAS -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Início das Obras:</span></td>
        <td bgcolor="#CCCCCC">
          <input name="dta_inicio_obras" type="text" class="datepicker" class="style9" value="<%=(rs_pi.Fields.Item("dta_inicio_obras").Value)%>">
        </td>
      </tr>

      <!-- % EXECUTADO -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">% Executado:</span></td>
        <td bgcolor="#CCCCCC">
          <input name="num_percentual_executado" type="text" class="style9" value="<%=(rs_pi.Fields.Item("num_percentual_executado").Value)%>">
        </td>
      </tr>

      <!-- PREVISÃO DE TÉRMINO -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Previsão de Término:</span></td>
        <td bgcolor="#CCCCCC">
          <input name="dta_previsao_termino" type="text" class="datepicker" class="style9" value="<%=(rs_pi.Fields.Item("dta_previsao_termino").Value)%>">
        </td>
      </tr>

      <!-- CONCLUÍDA INAUGURADA -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Concluída/Inaugurada em:</span></td>
        <td bgcolor="#CCCCCC">
          <input name="dta_inauguracao" type="text" class="datepicker" class="style9" value="<%=(rs_pi.Fields.Item("dta_inauguracao").Value)%>">
        </td>
      </tr>

      <!-- INÍCIO DAS OBRAS -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Previsão de Inauguração:</span></td>
        <td bgcolor="#CCCCCC">
          <input name="dta_previsao_inauguracao" type="text" class="style9" value="<%=(rs_pi.Fields.Item("dta_previsao_inauguracao").Value)%>">
        </td>
      </tr>

      <!-- CARGA ORGÂNICA RETIRADA -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Carga Orgânica Retirada (ton./mês):</span></td>
        <td bgcolor="#CCCCCC">
          <%
            If Not IsNull(rs_pi.Fields.Item("qtd_populacao_urbana_2010").Value) Then
              If Not IsNull(b) Then
                ' Base de cálculo = qtd_populacao_urbana_2030 * 0,06 * 30 / 1000
              End If
            End If
          %>
          <input name="qtd_populacao_urbana_2030" type="text" class="style9" value="<%=(b * 0.0018)%>" disabled="disabled">
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
                If Trim(rs_fiscal.Fields.Item("Responsável").Value) <> "" Then
                  Response.Write "      <OPTION value='" & (rs_fiscal.Fields.Item("cod_fiscal").Value) & "'"
                  If Lcase(rs_fiscal.Fields.Item("cod_fiscal").Value) = Lcase(rs_pi.Fields.Item("cod_fiscal").Value) then
                    Response.Write "selected"
                  End If
                  Response.Write ">" & (rs_fiscal.Fields.Item("Responsável").Value) & "</OPTION>"
                End If
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

      <!-- ENG. DAEE  -->
      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style10">Selecione o Eng. DAEE:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <select name="cod_engenheiro_daee" class="style9">
            <option value=""></option>
            <%
              While (NOT rs_fiscal.EOF)
                If Trim(rs_fiscal.Fields.Item("Responsável").Value) <> "" Then
                  Response.Write "      <OPTION value='" & (rs_fiscal.Fields.Item("cod_fiscal").Value) & "'"
                  If Lcase(rs_fiscal.Fields.Item("cod_fiscal").Value) = Lcase(rs_pi.Fields.Item("cod_engenheiro_daee").Value) then
                    Response.Write "selected"
                  End If
                  Response.Write ">" & (rs_fiscal.Fields.Item("Responsável").Value) & "</OPTION>"
                End If
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
                If Trim(rs_fiscal.Fields.Item("Responsável").Value) <> "" Then
                  Response.Write "      <OPTION value='" & (rs_fiscal.Fields.Item("cod_fiscal").Value) & "'"
                  If Lcase(rs_fiscal.Fields.Item("cod_fiscal").Value) = Lcase(rs_pi.Fields.Item("cod_engenheiro_plan_consorcio").Value) then
                    Response.Write "selected"
                  End If
                  Response.Write ">" & (rs_fiscal.Fields.Item("Responsável").Value) & "</OPTION>"
                End If
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
                If Trim(rs_fiscal.Fields.Item("Responsável").Value) <> "" Then
                  Response.Write "      <OPTION value='" & (rs_fiscal.Fields.Item("cod_fiscal").Value) & "'"
                  If Lcase(rs_fiscal.Fields.Item("cod_fiscal").Value) = Lcase(rs_pi.Fields.Item("cod_fiscal_consorcio").Value) then
                    Response.Write "selected"
                  End If
                  Response.Write ">" & (rs_fiscal.Fields.Item("Responsável").Value) & "</OPTION>"
                End If
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
                If Trim(rs_fiscal.Fields.Item("Responsável").Value) <> "" Then
                  Response.Write "      <OPTION value='" & (rs_fiscal.Fields.Item("cod_fiscal").Value) & "'"
                  If Lcase(rs_fiscal.Fields.Item("cod_fiscal").Value) = Lcase(rs_pi.Fields.Item("cod_engenheiro_medicao").Value) then
                    Response.Write "selected"
                  End If
                  Response.Write ">" & (rs_fiscal.Fields.Item("Responsável").Value) & "</OPTION>"
                End If
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
                If Trim(rs_fiscal.Fields.Item("Responsável").Value) <> "" Then
                  Response.Write "      <OPTION value='" & (rs_fiscal.Fields.Item("cod_fiscal").Value) & "'"
                  If Lcase(rs_fiscal.Fields.Item("cod_fiscal").Value) = Lcase(rs_pi.Fields.Item("cod_engenheiro_construtora").Value) then
                    Response.Write "selected"
                  End If
                  Response.Write ">" & (rs_fiscal.Fields.Item("Responsável").Value) & "</OPTION>"
                End If
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
                If Trim(rs_situacao.Fields.Item("desc_situacao").Value) <> "" Then
                  Response.Write "      <OPTION value='" & (rs_situacao.Fields.Item("cod_situacao").Value) & "'"
                  If Lcase(rs_situacao.Fields.Item("cod_situacao").Value) = Lcase(rs_pi.Fields.Item("cod_situacao").Value) then
                    Response.Write "selected"
                  End If
                  Response.Write ">Status: " & (rs_situacao.Fields.Item("desc_situacao").Value) & " - Situação: " & (rs_situacao.Fields.Item("cod_atendimento").Value) & "</OPTION>"
                End If
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
                If Trim(rs_situacao.Fields.Item("desc_situacao").Value) <> "" Then
                  Response.Write "      <OPTION value='" & (rs_situacao.Fields.Item("cod_situacao").Value) & "'"
                  If Lcase(rs_situacao.Fields.Item("cod_situacao").Value) = Lcase(rs_pi.Fields.Item("cod_situacao_externa").Value) then
                    Response.Write "selected"
                  End If
                  Response.Write ">Status: " & (rs_situacao.Fields.Item("desc_situacao").Value) & " - Situação: " & (rs_situacao.Fields.Item("cod_atendimento").Value) & "</OPTION>"
                End If
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

      <!-- OBSERVAÇÕES RELATÓRIO MENSAL -->
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap bgcolor="#CCCCCC" class="style9"><span class="style10">Observações Relatório Mensal:</span></td>
        <td bgcolor="#CCCCCC">
          <textarea name="dsc_observacoes_relatorio_mensal"cols="50" rows="5" class="style9" style="width: 98%;"><%=(rs_pi.Fields.Item("dsc_observacoes_relatorio_mensal").Value)%></textarea>
        </td>
      </tr>

      <tr valign="baseline">
        <td align="center" bgcolor="#CCCCCC" class="style9" colspan="2">
          <strong>Informações Complementares</strong>
        </td>
      </tr>

      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style22">Bacia Hidrográfica:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <select name="cod_bacia_hidrografica">
            <option value=""></option>
            <%
              strQ = "SELECT * FROM tb_bacia_hidrografica "

              Set rs_combo = Server.CreateObject("ADODB.Recordset")
                rs_combo.CursorLocation = 3
                rs_combo.CursorType = 3
                rs_combo.LockType = 1
                rs_combo.Open strQ, objCon, , , &H0001

              If Not rs_combo.EOF Then
                While Not rs_combo.EOF
                   If Trim(rs_combo.Fields.Item("nme_bacia_hidrografica").Value) <> "" Then
                    Response.Write "      <OPTION value='" & (rs_combo.Fields.Item("id").Value) & "'"
                    If Lcase(rs_combo.Fields.Item("id").Value) = Lcase(rs_pi.Fields.Item("cod_bacia_hidrografica").Value) then
                      Response.Write "selected"
                    End If
                    Response.Write ">" & (rs_combo.Fields.Item("nme_bacia_hidrografica").Value) & "</OPTION>"
                  End If
                  rs_combo.MoveNext
                Wend
              End If
            %>
          </select>
        </td>
      </tr>

      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style22">Manancial de Lançamento:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <select name="cod_manancial_lancamento">
            <option value=""></option>
            <%
              strQ = "SELECT * FROM tb_manancial_lancamento"

              Set rs_combo = Server.CreateObject("ADODB.Recordset")
                rs_combo.CursorLocation = 3
                rs_combo.CursorType = 3
                rs_combo.LockType = 1
                rs_combo.Open strQ, objCon, , , &H0001

              If Not rs_combo.EOF Then
                While Not rs_combo.EOF
                  If Trim(rs_combo.Fields.Item("nme_manancial").Value) <> "" Then
                    Response.Write "      <OPTION value='" & (rs_combo.Fields.Item("id").Value) & "'"
                    If Lcase(rs_combo.Fields.Item("id").Value) = Lcase(rs_pi.Fields.Item("cod_manancial_lancamento").Value) then
                      Response.Write "selected"
                    End If
                    Response.Write ">" & (rs_combo.Fields.Item("nme_manancial").Value) & "</OPTION>"
                  End If
                  rs_combo.MoveNext
                Wend
              End If
            %>
          </select>
        </td>
      </tr>

      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style22">Coletor Tronco (m):</span>
        </td>
        <td bgcolor="#CCCCCC">
          <input type="text" name="qtd_metragem_coletor_tronco" value="<%=(rs_pi.Fields.Item("qtd_metragem_coletor_tronco").Value)%>" size="32">
        </td>
      </tr>

      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style22">Interceptor (m):</span>
        </td>
        <td bgcolor="#CCCCCC">
          <input type="text" name="qtd_metragem_interceptor" value="<%=(rs_pi.Fields.Item("qtd_metragem_interceptor").Value)%>" size="32"/>
        </td>
      </tr>

      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style22">Emissário de Efluente Bruto (m):</span>
        </td>
        <td bgcolor="#CCCCCC">
          <input type="text" name="qtd_metragem_emissario_fluente_bruto" value="<%=(rs_pi.Fields.Item("qtd_metragem_emissario_fluente_bruto").Value)%>" size="32"/>
        </td>
      </tr>

      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style22">Estação Elevatória de Esgoto (und):</span>
        </td>
        <td bgcolor="#CCCCCC">
          <input type="text" name="qtd_eee" value="<%=(rs_pi.Fields.Item("qtd_eee").Value)%>" size="32">
        </td>
      </tr>

      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style22">Linha de Recalque (m):</span>
        </td>
        <td bgcolor="#CCCCCC">
          <input type="text" name="qtd_metragem_linha_recalque" value="<%=(rs_pi.Fields.Item("qtd_metragem_linha_recalque").Value)%>" size="32">
        </td>
      </tr>

      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style22">Tipo de ETE:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <select name="cod_tipo_ete">
            <option value=""></option>
            <%
              strQ = "SELECT * FROM tb_tipo_ete "

              Set rs_combo = Server.CreateObject("ADODB.Recordset")
                rs_combo.CursorLocation = 3
                rs_combo.CursorType = 3
                rs_combo.LockType = 1
                rs_combo.Open strQ, objCon, , , &H0001

              If Not rs_combo.EOF Then
                While Not rs_combo.EOF
                  If Trim(rs_combo.Fields.Item("nme_tipo_ete").Value) <> "" Then
                    Response.Write "      <OPTION value='" & (rs_combo.Fields.Item("id").Value) & "'"
                    If Lcase(rs_combo.Fields.Item("id").Value) = Lcase(rs_pi.Fields.Item("cod_tipo_ete").Value) then
                      Response.Write "selected"
                    End If
                    Response.Write ">" & (rs_combo.Fields.Item("nme_tipo_ete").Value) & "</OPTION>"
                  End If
                  rs_combo.MoveNext
                Wend
              End If
            %>
          </select>
        </td>
      </tr>

      <tr valign="middle">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style22">Estação de Tratamento:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <textarea name="dsc_estacao_tratamento" cols="25" style="width: 98%;"><%=(rs_pi.Fields.Item("dsc_estacao_tratamento").Value)%></textarea>
        </td>
      </tr>

      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style22">Emissário de Efluente Tratado (m):</span>
        </td>
        <td bgcolor="#CCCCCC">
          <input type="text" name="qtd_metragem_emissario_efluente_tratado" value="<%=(rs_pi.Fields.Item("qtd_metragem_emissario_efluente_tratado").Value)%>" size="32">
        </td>
      </tr>

      <tr valign="baseline">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style22">Estudo Elaborado pelo DAEE?:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <select name="flg_estudo_elaborado_daee">
            <option value=""></option>
            <option value="Sim"
              <%
                If Lcase("Sim") = Lcase(rs_pi.Fields.Item("flg_estudo_elaborado_daee").Value) then
                  Response.Write "selected"
                End If
              %>
            >Sim</option>
            <option value="Não"
              <%
                If Lcase("Não") = Lcase(rs_pi.Fields.Item("flg_estudo_elaborado_daee").Value) then
                  Response.Write "selected"
                End If
              %>
            >Não</option>
          </select>
        </td>
      </tr>

      <tr valign="middle">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style22">Observações da Bacia:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <textarea name="dsc_observacoes" cols="50" rows="5" class="style9" style="width: 98%;"><%=(rs_pi.Fields.Item("dsc_observacoes_obra").Value)%></textarea>
        </td>
      </tr>

      <tr valign="middle">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style22">Parceria/Realização:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <select name="cod_parceiro">
            <option value=""></option>
            <%
              strQ = "SELECT * FROM tb_parceiro "

              Set rs_combo = Server.CreateObject("ADODB.Recordset")
                rs_combo.CursorLocation = 3
                rs_combo.CursorType = 3
                rs_combo.LockType = 1
                rs_combo.Open strQ, objCon, , , &H0001

              If Not rs_combo.EOF Then
                While Not rs_combo.EOF
                  If Trim(rs_combo.Fields.Item("nme_parceiro").Value) <> "" Then
                    Response.Write "      <OPTION value='" & (rs_combo.Fields.Item("id").Value) & "'"
                    If Lcase(rs_combo.Fields.Item("id").Value) = Lcase(rs_pi.Fields.Item("cod_parceiro").Value) then
                      Response.Write "selected"
                    End If
                    Response.Write ">" & (rs_combo.Fields.Item("nme_parceiro").Value) & "</OPTION>"
                  End If
                  rs_combo.MoveNext
                Wend
              End If
            %>
          </select>
          <br/>
          <textarea name="dsc_parceria_realizacao" cols="25" style="width: 98%;"><%=(rs_pi.Fields.Item("dsc_parceria_realizacao").Value)%></textarea>
        </td>
      </tr>

      <tr valign="middle">
        <td align="right" nowrap bgcolor="#CCCCCC" class="style9">
          <span class="style22">Observações Gerais:</span>
        </td>
        <td bgcolor="#CCCCCC">
          <textarea name="dsc_observacoes_gestor" cols="50" rows="5" class="style9" style="width: 98%;"><%=(rs_pi.Fields.Item("dsc_observacoes_gestor").Value)%></textarea>
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
</body>
</html>
<%
rs_pi.Close()
Set rs_pi = Nothing
%>