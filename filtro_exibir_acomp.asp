<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim rs_filtro_combo__MMColParam
rs_filtro_combo__MMColParam = "1"
If (Request.Form("cod_predio") <> "") Then 
  rs_filtro_combo__MMColParam = Request.Form("cod_predio")
End If
%>
<%
Dim rs_filtro_combo
Dim rs_filtro_combo_cmd
Dim rs_filtro_combo_numRows

Set rs_filtro_combo_cmd = Server.CreateObject ("ADODB.Command")
rs_filtro_combo_cmd.ActiveConnection = MM_cpf_STRING
rs_filtro_combo_cmd.CommandText = "SELECT tb_pi.*, tb_situacao_pi.desc_situacao AS situacao_interna, tb_situacao_pi_1.desc_situacao AS situacao_externa, tb_predio.Município FROM ((tb_situacao_pi RIGHT JOIN tb_pi ON tb_situacao_pi.cod_situacao = tb_pi.cod_situacao) LEFT JOIN tb_situacao_pi AS tb_situacao_pi_1 ON tb_pi.cod_situacao_externa = tb_situacao_pi_1.cod_situacao) INNER JOIN tb_predio ON (tb_predio.cod_predio = tb_pi.cod_predio) AND (tb_pi.id_predio = tb_predio.id_predio) WHERE tb_pi.cod_predio = ?" 
rs_filtro_combo_cmd.Prepared = true
rs_filtro_combo_cmd.Parameters.Append rs_filtro_combo_cmd.CreateParameter("param1", 200, 1, 255, rs_filtro_combo__MMColParam) ' adVarChar

Set rs_filtro_combo = rs_filtro_combo_cmd.Execute
rs_filtro_combo_numRows = 0
%>
<%
Dim rs_predio__MMColParam
rs_predio__MMColParam = "1"
If (Request.Form("cod_predio") <> "") Then 
  rs_predio__MMColParam = Request.Form("cod_predio")
End If
%>
<%
Dim rs_predio
Dim rs_predio_numRows

Set rs_predio = Server.CreateObject("ADODB.Recordset")
rs_predio.ActiveConnection = MM_cpf_STRING
rs_predio.Source = "SELECT cod_predio, Município FROM tb_predio WHERE cod_predio = '" + Replace(rs_predio__MMColParam, "'", "''") + "'"
rs_predio.CursorType = 0
rs_predio.CursorLocation = 2
rs_predio.LockType = 1
rs_predio.Open()

rs_predio_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rs_predio_numRows = rs_predio_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = 10
Repeat2__index = 0
rs_filtro_combo_numRows = rs_filtro_combo_numRows + Repeat2__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style9 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; color: #FFFFFF; }
.style13 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; color: #000000; }
.style15 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; color: #333333; }
-->
</style>
</head>

<body>
<div align="center">
  <table border="0">
    <tr bgcolor="#666666">
      <td width="311"><span class="style9">Município</span></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rs_predio.EOF)) %>
      <tr bgcolor="#CCCCCC">
        <td><span class="style13"><%=(rs_predio.Fields.Item("Município").Value)%></span></td>
      </tr>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_predio.MoveNext()
Wend
%>
  </table>
</div>
<p>&nbsp;</p>

<div align="center">
  <table border="0">
    <tr bgcolor="#666666">
      <td width="108"><span class="style9">Autos</span></td>
      <td width="308"><span class="style9">Nome do Empreendimento</span></td>
      <%
        If Session("MM_UserAuthorization") = 3 Then
      %>

      <td><span class="style9">Situação Externa</span></td>

      <%
        End If
        
        If Session("MM_UserAuthorization") = 1 Or Session("MM_UserAuthorization") = 4 Then
      %>

      <td><span class="style9">Situação Interna</span></td>

      <%
        End If

        If Session("MM_UserAuthorization") = 7 Then
      %>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <%
        End If
      %>
    </tr>
    <% While ((Repeat2__numRows <> 0) AND (NOT rs_filtro_combo.EOF)) %>
      <tr bgcolor="#CCCCCC">
        <td><a href="acompanhamento_inclui_admin.asp?pi=<%=(rs_filtro_combo.Fields.Item("PI").Value)%>" target="_blank" class="style15"><%=(rs_filtro_combo.Fields.Item("PI").Value)%></a></td>
        <td><div align="left"><span class="style15"><%=(rs_filtro_combo.Fields.Item("nome_empreendimento").Value)%></span></div></td>

        <%
          If Session("MM_UserAuthorization") = 3 Then
        %>

        <td><div align="left"><span class="style15"><%=(rs_filtro_combo.Fields.Item("situacao_externa").Value)%></span></div></td>

        <%
          End If

          If Session("MM_UserAuthorization") = 1 Or Session("MM_UserAuthorization") = 4 Then
        %>

        <td><div align="left"><span class="style15"><%=(rs_filtro_combo.Fields.Item("situacao_interna").Value)%></span></div></td>

        <%
          End If

          If Session("MM_UserAuthorization") = 7 Then
        %>

        <td>
          <span class="style15"><a href="cad_licenca.asp?cod_empreendimento=<%=(rs_filtro_combo.Fields.Item("PI").Value)%>&nme_municipio=<%=(rs_filtro_combo.Fields.Item("Município").Value)%>&nme_empreendimento=<%=(rs_filtro_combo.Fields.Item("nome_empreendimento").Value)%>">Licenças Ambientais</a></span>
        </td>
        <td>
          <span class="style15"><a href="cad_outorga.asp?cod_empreendimento=<%=(rs_filtro_combo.Fields.Item("PI").Value)%>&nme_municipio=<%=(rs_filtro_combo.Fields.Item("Município").Value)%>&nme_empreendimento=<%=(rs_filtro_combo.Fields.Item("nome_empreendimento").Value)%>">Outorgas</a></span>
        </td>
        <td>
          <span class="style15"><a href="cad_app.asp?cod_empreendimento=<%=(rs_filtro_combo.Fields.Item("PI").Value)%>&nme_municipio=<%=(rs_filtro_combo.Fields.Item("Município").Value)%>&nme_empreendimento=<%=(rs_filtro_combo.Fields.Item("nome_empreendimento").Value)%>">Intervenções em App</a></span>
        </td>
        <td>
          <span class="style15"><a href="cad_tcra.asp?cod_empreendimento=<%=(rs_filtro_combo.Fields.Item("PI").Value)%>&nme_municipio=<%=(rs_filtro_combo.Fields.Item("Município").Value)%>&nme_empreendimento=<%=(rs_filtro_combo.Fields.Item("nome_empreendimento").Value)%>">TCRA</a></span>
        </td>

        <%
          End If
        %>
      </tr>
      <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rs_filtro_combo.MoveNext()
Wend
%>
  </table>
 </div>
</body>
</html>
<%
rs_filtro_combo.Close()
Set rs_filtro_combo = Nothing
%>
<%
rs_predio.Close()
Set rs_predio = Nothing
%>