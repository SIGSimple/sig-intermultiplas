<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
' *** Edit Operations: declare variables

Response.CharSet = "UTF-8"

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
  MM_editTable = "tb_Acompanhamento"
  MM_editRedirectUrl = "acompanhamento_inclui_admin.asp"
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '                     CAMPOS TELA                     '
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''

  MM_fieldsStr = "cod_usuario_lancamento|value|PI|value|Data_do_Registro|value|Registro|value|Responsvel|value"
  
  If Session("MM_UserAuthorization") = 1 Or Session("MM_UserAuthorization") = 5 Then
    MM_fieldsStr = MM_fieldsStr & "|cod_situacao_sso|value|dsc_nota_sso|value|cod_situacao_clima_manha|value|cod_situacao_clima_tarde|value|cod_situacao_clima_noite|value|cod_situacao_limpeza_obra|value|cod_situacao_organizacao_obra|value|dsc_nota_clima_manha|value|dsc_nota_clima_tarde|value|dsc_nota_clima_noite|value|dsc_nota_limpeza_obra|value|dsc_nota_organizacao_obra|value|dsc_nota_dia_perdido|value|dsc_nota_dia_trabalhado|value"
  End If

  If Session("MM_UserAuthorization") <> 6 And Session("MM_UserAuthorization") <> 3 Then
    MM_fieldsStr = MM_fieldsStr & "|vistoria|value|dt_vistoria|value"
  End If
  
  MM_fieldsStr = MM_fieldsStr & "|flg_pendencia|value|cod_tipo_pendencia|value|dsc_pendencia|value|dta_limite_pendencia|value|cod_tipo_registro|value"

  '''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '                     COLUNAS BD                      '
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''

  MM_columnsStr = "cod_usuario_lancamento|none,none,NULL|PI|',none,''|[Data do Registro]|',none,NULL|Registro|',none,''|cod_fiscal|none,none,NULL"

  If Session("MM_UserAuthorization") = 1 Or Session("MM_UserAuthorization") = 5 Then
    MM_columnsStr = MM_columnsStr & "|cod_situacao_sso|none,none,NULL|dsc_nota_sso|',none,''|cod_situacao_clima_manha|none,none,NULL|cod_situacao_clima_tarde|none,none,NULL|cod_situacao_clima_noite|none,none,NULL|cod_situacao_limpeza_obra|none,none,NULL|cod_situacao_organizacao_obra|none,none,NULL|dsc_nota_clima_manha|',none,''|dsc_nota_clima_tarde|',none,''|dsc_nota_clima_noite|',none,''|dsc_nota_limpeza_obra|',none,''|dsc_nota_organizacao_obra|',none,''|dsc_nota_dia_perdido|',none,''|dsc_nota_dia_trabalhado|',none,''"
  End If

  If Session("MM_UserAuthorization") <> 6 And Session("MM_UserAuthorization") <> 3 Then
    MM_columnsStr = MM_columnsStr & "|vistoria|none,1,0|dt_vistoria|',none,NULL"
  End If

  MM_columnsStr = MM_columnsStr & "|flg_pendencia|none,1,0|cod_tipo_pendencia|none,none,NULL|dsc_pendencia|',none,''|dta_limite_pendencia|',none,NULL|cod_tipo_registro|none,none,NULL"

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
rs_pi.Source = "SELECT * FROM c_lista_pi WHERE PI = '" + Replace(rs_pi__MMColParam, "'", "''") + "'"
rs_pi.CursorType = 0
rs_pi.CursorLocation = 2
rs_pi.LockType = 1
rs_pi.Open()

rs_pi_numRows = 0
%>
<%
Dim rs_situacao_sso
Dim rs_situacao_sso_numRows

Set rs_situacao_sso = Server.CreateObject("ADODB.Recordset")
rs_situacao_sso.ActiveConnection = MM_cpf_STRING
rs_situacao_sso.Source = "SELECT *  FROM tb_situacao_sso"
rs_situacao_sso.CursorType = 0
rs_situacao_sso.CursorLocation = 2
rs_situacao_sso.LockType = 1
rs_situacao_sso.Open()

rs_situacao_sso_numRows = 0
%>
<%
Dim rs_situacao_clima
Dim rs_situacao_clima_numRows

Set rs_situacao_clima = Server.CreateObject("ADODB.Recordset")
rs_situacao_clima.ActiveConnection = MM_cpf_STRING
rs_situacao_clima.Source = "SELECT *  FROM tb_situacao_clima"
rs_situacao_clima.CursorType = 0
rs_situacao_clima.CursorLocation = 2
rs_situacao_clima.LockType = 1
rs_situacao_clima.Open()

rs_situacao_clima_numRows = 0
%>
<%
Dim rs_situacao_obra
Dim rs_situacao_obra_numRows

Set rs_situacao_obra = Server.CreateObject("ADODB.Recordset")
rs_situacao_obra.ActiveConnection = MM_cpf_STRING
rs_situacao_obra.Source = "SELECT *  FROM tb_situacao_obra"
rs_situacao_obra.CursorType = 0
rs_situacao_obra.CursorLocation = 2
rs_situacao_obra.LockType = 1
rs_situacao_obra.Open()

rs_situacao_obra_numRows = 0
%>
<%
Dim rs_fiscal
Dim rs_fiscal_numRows

Set rs_fiscal = Server.CreateObject("ADODB.Recordset")
rs_fiscal.ActiveConnection = MM_cpf_STRING
rs_fiscal.Source = "SELECT *  FROM c_lista_fiscal"
rs_fiscal.CursorType = 0
rs_fiscal.CursorLocation = 2
rs_fiscal.LockType = 1
rs_fiscal.Open()

rs_fiscal_numRows = 0
%>
<%
Dim rs_lista_acomp__MMColParam
rs_lista_acomp__MMColParam = "1"
If (Request.QueryString("PI") <> "") Then 
  rs_lista_acomp__MMColParam = Request.QueryString("PI")
End If
%>
<%
Dim rs_lista_acomp
Dim rs_lista_acomp_numRows
Dim cod_tipo_registro

Select Case Session("MM_UserAuthorization")
  Case 1,4,5
    cod_tipo_registro = 1
  Case 2
    cod_tipo_registro = 3
  Case 6
    cod_tipo_registro = 0
  Case 7
    cod_tipo_registro = 2
End Select

sql = "SELECT *  FROM c_lista_acompanhamento WHERE PI = '" + Replace(rs_lista_acomp__MMColParam, "'", "''") + "' "

'If cod_tipo_registro <> "" Then
''  sql = sql & " AND cod_tipo_registro="& cod_tipo_registro
'End If

sql = sql & " ORDER BY [Data do Registro] DESC"

Set rs_lista_acomp = Server.CreateObject("ADODB.Recordset")
rs_lista_acomp.ActiveConnection = MM_cpf_STRING
rs_lista_acomp.Source = sql
rs_lista_acomp.CursorType = 0
rs_lista_acomp.CursorLocation = 2
rs_lista_acomp.LockType = 1
rs_lista_acomp.Open()

rs_lista_acomp_numRows = 0
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
<!DOCTYPE html>
<html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <title>Acompanhamento</title>
    <style type="text/css">
      .style5 {font-family: Arial, Helvetica, sans-serif; font-size: 12; }
      .style6 {font-size: 12}
      .style19 {font-family: Arial, Helvetica, sans-serif; font-size: 14px; font-weight: bold; }
      .style22 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
      .style26 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
      .style27 {
        font-family: Arial, Helvetica, sans-serif;
        font-size: 12px;
      }
      .style1 {font-family: Arial, Helvetica, sans-serif;
        font-weight: bold;
        font-size: 10;
      }
      .style28 {font-family: Arial, Helvetica, sans-serif; font-size: 10; }
      .style29 {font-family: Arial, Helvetica, sans-serif}
    </style>

    <script language="JavaScript">
      function FP_preloadImgs() {//v1.0
       var d=document,a=arguments; if(!d.FP_imgs) d.FP_imgs=new Array();
       for(var i=0; i<a.length; i++) { d.FP_imgs[i]=new Image; d.FP_imgs[i].src=a[i]; }
      }

      function FP_swapImg() {//v1.0
       var doc=document,args=arguments,elm,n; doc.$imgSwaps=new Array(); for(n=2; n<args.length;
       n+=2) { elm=FP_getObjectByID(args[n]); if(elm) { doc.$imgSwaps[doc.$imgSwaps.length]=elm;
       elm.$src=elm.src; elm.src=args[n+1]; } }
      }

      function FP_getObjectByID(id,o) {//v1.0
       var c,el,els,f,m,n; if(!o)o=document; if(o.getElementById) el=o.getElementById(id);
       else if(o.layers) c=o.layers; else if(o.all) el=o.all[id]; if(el) return el;
       if(o.id==id || o.name==id) return o; if(o.childNodes) c=o.childNodes; if(c)
       for(n=0; n<c.length; n++) { el=FP_getObjectByID(id,c[n]); if(el) return el; }
       f=o.forms; if(f) for(n=0; n<f.length; n++) { els=f[n].elements;
       for(m=0; m<els.length; m++){ el=FP_getObjectByID(id,els[n]); if(el) return el; } }
       return null;
      }
    </script>

    <script language="JavaScript" type="text/javascript">
      function abre_janela(width, height, nome) {
      var top; var left;
      top = ( (screen.height/2) - (height/2) )
      left = ( (screen.width/2) - (width/2) )
      window.open('',nome,'width='+width+',height='+height+',scrollbars=no,toolbar=no,location=no,status=no,menubar=no,resizable=no,left='+left+',top='+top);
      }
      function recebe_imagem(campo, imagem){
      var foto = 'img_' + campo
      document.form_incluir[campo].value = imagem;
      document.form_incluir[foto].src = imagem;
      }
      function verifica_form(form) {
      var passed = false;
      var ok = false
      var campo
      for (i = 0; i < form.length; i++) {
        campo = form[i].name;
        if (form[i].df_verificar == "sim") {
          if (form[i].type == "text"  | form[i].type == "textarea" | form[i].type == "select-one") {
            if (form[i].value == "" | form[i].value == "http://") {
          form[campo].className='campo_alerta'
              form[campo].focus();
              alert("Preencha corretamente o campo");
              return passed;
              stop;
            }
          }
          else if (form[i].type == "radio") {
            for (x = 0; x < form[campo].length; x++) {
              ok = false;
              if (form[campo][x].checked) {
                ok = true;
                break;
              }
            }
            if (ok == false) {
              form[campo][0].focus();
          form[campo][0].select();
              alert("Informe uma das opcões");
              return passed;
              stop;
            }
          }
          var msg = ""
          if (form[campo].df_validar == "cpf") msg = checa_cpf(form[campo].value);
          if (form[campo].df_validar == "cnpj") msg = checa_cnpj(form[campo].value);
          if (form[campo].df_validar == "cpf_cnpj") {
          msg = checa_cpf(form[campo].value);
          if (msg != "") msg = checa_cnpj(form[campo].value);
        }
          if (form[campo].df_validar == "email") msg = checa_email(form[campo].value);
          if (form[campo].df_validar == "numerico") msg = checa_numerico(form[campo].value);
          if (msg != "") {
          if (form[campo].df_validar == "cpf_cnpj") msg = "informe corretamente o número do CPF ou CNPJ";
          form[campo].className='campo_alerta'
            form[campo].focus();
            form[campo].select();
            alert(msg);
            return passed;
            stop;
          }
        }
      }
      passed = true;
      return passed;
      }
      function desabilita_cor(campo) {
      campo.className='campos_formulario'
      }
      function checa_numerico(String) {
      var mensagem = "Este campo aceita somente números"
      var msg = "";
      if (isNaN(String)) msg = mensagem;
      return msg;
      }
      function checa_email(campo) {
      var mensagem = "Informe corretamente o email"
      var msg = "";
      var email = campo.match(/(\w+)@(.+)\.(\w+)$/);
      if (email == null){
        msg = mensagem;
        }
      return msg;
      }
      function checa_cpf(CPF) {
      var mensagem = "informe corretamente o número do CPF"
      var msg = "";
      if (CPF.length != 11 || CPF == "00000000000" || CPF == "11111111111" ||
        CPF == "22222222222" || CPF == "33333333333" || CPF == "44444444444" ||
        CPF == "55555555555" || CPF == "66666666666" || CPF == "77777777777" ||
        CPF == "88888888888" || CPF == "99999999999")
      msg = mensagem;
      soma = 0;
      for (y=0; y < 9; y ++)
      soma += parseInt(CPF.charAt(y)) * (10 - y);
      resto = 11 - (soma % 11);
      if (resto == 10 || resto == 11)resto = 0;
      if (resto != parseInt(CPF.charAt(9)))
        msg = mensagem; soma = 0;
      for (y = 0; y < 10; y ++)
        soma += parseInt(CPF.charAt(y)) * (11 - y);
      resto = 11 - (soma % 11);
      if (resto == 10 || resto == 11) resto = 0;
      if (resto != parseInt(CPF.charAt(10)))
        msg = mensagem;
      return msg;
      }
      function checa_cnpj(s) {
      var mensagem = "informe corretamente o número do CNPJ"
      var msg = "";
      var y;
      var c = s.substr(0,12);
      var dv = s.substr(12,2);
      var d1 = 0;
      for (y = 0; y < 12; y++)
      {
      d1 += c.charAt(11-y)*(2+(y % 8));
      }
      if (d1 == 0) msg = mensagem;
      d1 = 11 - (d1 % 11);
      if (d1 > 9) d1 = 0;
      if (dv.charAt(0) != d1)msg = mensagem;
      d1 *= 2;
      for (y = 0; y < 12; y++)
      {
      d1 += c.charAt(11-y)*(2+((y+1) % 8));
      }
      d1 = 11 - (d1 % 11);
      if (d1 > 9) d1 = 0;
      if (dv.charAt(1) != d1) msg = mensagem;
      return msg;
      }
      function mascara_data(data){ 
      var mydata = ''; 

      mydata = mydata + data; 
      if (mydata.length == 2){ 
      mydata = mydata + '/'; 
      } 
      if (mydata.length == 5){ 
      mydata = mydata + '/'; 
      } 
      return mydata; 
      } 
      function verifica_data(data) { 
      if (data.value != "") {
      dia = (data.value.substring(0,2));
      mes = (data.value.substring(3,5)); 
      ano = (data.value.substring(6,10)); 
      situacao = ""; 
      if ((dia < 01)||(dia < 01 || dia > 30) && (  mes == 04 || mes == 06 || mes == 09 || mes == 11 ) || dia > 31) { 
      situacao = "falsa"; 
      } 

      if (mes < 01 || mes > 12 ) { 
      situacao = "falsa"; 
      }
      if (mes == 2 && ( dia < 01 || dia > 29 || ( dia > 28 && (parseInt(ano / 4) != ano / 4)))) { 
      situacao = "falsa"; 
      } 
      if (situacao == "falsa") { 
      data.focus();
      data.select();
      alert("Data inválida!"); 
      }
      } 
      }
    </script>

    <style type="text/css">
      label {
          display: block;
      }
      #campoA,
      {
          background-color: #;
          margin-left: 20px;

      }
    </style>

    <script type="text/javascript">
      function mostraEsconde(empresa,LO)
      {
          var check = document.getElementById(empresa);
          var LO = document.getElementById(LO);
          
          if(check.checked==true)
              LO.style.display = 'block';
          else
              LO.style.display = 'none';

      }
    </script>

    <style type="text/css">
      body {
        font-family: Arial, Helvetica, sans-serif !important;
        font-size: 12px !important;
      }

      textarea {
        width: 98%;
      }

      table.dados-empreendimento tr td {
        background-color: #ccc;
      }

      table.lista-acompanhamento tr.title {
        font-weight: bold;
      }

      table.lista-acompanhamento tr.title td {
        text-align: center;
      }

      .text-center {
        text-align: center;
      }
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
    <div align="center">
      <h3>
        <%
          Select Case Session("MM_UserAuthorization")
            Case 1,3,4
              Response.Write "Registro de Informações"
            Case 5
              Response.Write "Registro de RDO"
            Case 2
              Response.Write "Registro de Ocorrência de Projeto"
            Case 6
              Response.Write "Registro de Ocorrência de Contrato, Medição ou Pagamento"
            Case 7
              Response.Write "Registro de Ocorrência de Meio Ambiente"
          End Select
        %>
      </h3>

      <table width="800" border="0" class="dados-empreendimento">
        <tr>
          <td width="100"><strong>Município:</strong></td>
          <td width="170"><span><%=(rs_pi.Fields.Item("nme_municipio").Value)%></span></td>
          <td width="100"><strong>Empreendimento:</strong></td>
          <td width="360"><span><%=(rs_pi.Fields.Item("nome_empreendimento").Value)%></span></td>
          <td width="35"><strong>Autos:</strong></td>
          <td width="80"><span><%=(rs_pi.Fields.Item("PI").Value)%></span></td>
        </tr>
        <tr>
          <td><strong>Descrição:</strong></td>
          <td colspan="5"><span><%=(rs_pi.Fields.Item("dsc_objeto_obra").Value)%></span></td>
        </tr>
        
        <%
          If Session("MM_UserAuthorization") = 1 Or Session("MM_UserAuthorization") = 3 Then
        %>
        
        <tr>
          <td><strong>Última Informação:</strong></td>
          <td colspan="5"><span><%=(rs_pi.Fields.Item("dsc_observacoes_relatorio_mensal").Value)%></span></td>
        </tr>
        
        <%
          End If
        %>

      </table>

      <table width="800" border="0" class="dados-empreendimento">
        <tr>
          <td width="95"><strong>Bacia DAEE:</strong></td>
          <td><span><%=(rs_pi.Fields.Item("bacia_daee").Value)%></span></td>
          <td width="100"><strong>Bacia Secretaria:</strong></td>
          <td><span><%=(rs_pi.Fields.Item("bacia_secretaria").Value)%></span></td>
        </tr>
      </table>

      <table width="800" border="0" class="dados-empreendimento">
        <tr>
          <td width="95"><strong>Programa:</strong></td>
          <td><span><%=(rs_pi.Fields.Item("programa").Value)%></span></td>
          <td width="95"><strong>Situação:</strong></td>
          <td><span><%=(rs_pi.Fields.Item("dsc_situacao_interna").Value)%></span></td>
        </tr>
      </table>
    </div>

    <br/>

    <%
      flg_can_insert = False

      If Session("MM_UserAuthorization") = 5 Then
        cod_eng_obras_consorcio = rs_pi.Fields.Item("cod_fiscal").Value
        cod_fiscal_consorcio    = rs_pi.Fields.Item("cod_fiscal_consorcio").Value

        If cod_fiscal = CInt(Session("MM_UserCodFiscal")) Or cod_eng_obras_consorcio = CInt(Session("MM_UserCodFiscal")) Then
          flg_can_insert = True
        End If
      Else
          flg_can_insert = True
      End If

      If flg_can_insert Then
    %>

    <form method="POST" action="<%=MM_editAction%>" name="form1">
      <table align="center">
        <!-- data -->
        <tr valign="baseline">
          <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC">
            <strong class="style5">Data do Registro:</strong>
          </td>
          <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
            <input name="Data_do_Registro" type="text" class="datepicker" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" value="<% =Date %>" size="15" />
          </td>
        </tr>
        
        <!-- responsável -->
        <tr valign="baseline">
          <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC">
            <strong class="style5">Responsável:</strong>
          </td>
          <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
            <select name="Responsvel" style="width: 100%;">
              <option value=""></option>
              <%
                While (NOT rs_fiscal.EOF)
              %>
              <option value="<%=(rs_fiscal.Fields.Item("cod_fiscal").Value)%>"><%=(rs_fiscal.Fields.Item("nme_interessado").Value)%></option>
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
        
        <%
          If Session("MM_UserAuthorization") <> 3 Then
        %>
        
        <!-- registro -->
        <tr valign="middle">
          <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC">
            <strong class="style5">Registro:</strong>
          </td>
          <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
            <textarea name="Registro" cols="32"></textarea>
          </td>
        </tr>

        <%
          End If

          If Session("MM_UserAuthorization") = 1 Or Session("MM_UserAuthorization") = 5 Then
        %>

        <!-- msst -->
        <tr valign="middle">
          <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC">
            <strong class="style5">MSST:</strong>
          </td>
          <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
            Tipo Situação: <br/>
            <select name="cod_situacao_sso" style="width: 100%;">
              <option value=""></option>
              <%
                While (NOT rs_situacao_sso.EOF)
              %>
              <option value="<%=(rs_situacao_sso.Fields.Item("id").Value)%>"><%=(rs_situacao_sso.Fields.Item("dsc_situacao_sso").Value)%></option>
              <%
                  rs_situacao_sso.MoveNext()
                  Wend
                  If (rs_situacao_sso.CursorType > 0) Then
                    rs_situacao_sso.MoveFirst
                  Else
                    rs_situacao_sso.Requery
                  End If
              %>
            </select>
            <br/>
            Nota/Observações: <br/>
            <textarea name="dsc_nota_sso" cols="32"></textarea>
          </td>
        </tr>

        <!-- condições da obra -->
        <tr valign="middle">
          <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC">
            <strong class="style5">Condições da Obra:</strong>
          </td>
          <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
            <table>
              <tr>
                <td>
                  Clima Manhã:
                </td>
                <td>
                  <select name="cod_situacao_clima_manha" style="width: 100%;">
                    <option value=""></option>
                    <%
                      While (NOT rs_situacao_clima.EOF)
                    %>
                    <option value="<%=(rs_situacao_clima.Fields.Item("id").Value)%>"><%=(rs_situacao_clima.Fields.Item("dsc_situacao_clima").Value)%></option>
                    <%
                        rs_situacao_clima.MoveNext()
                        Wend
                        If (rs_situacao_clima.CursorType > 0) Then
                          rs_situacao_clima.MoveFirst
                        Else
                          rs_situacao_clima.Requery
                        End If
                    %>
                  </select>
                </td>
                <td>
                  <input type="text" name="dsc_nota_clima_manha" placeholder="Notas">
                </td>
              </tr>
              <tr>
                <td>
                  Clima Tarde:
                </td>
                <td>
                  <select name="cod_situacao_clima_tarde" style="width: 100%;">
                    <option value=""></option>
                    <%
                      While (NOT rs_situacao_clima.EOF)
                    %>
                    <option value="<%=(rs_situacao_clima.Fields.Item("id").Value)%>"><%=(rs_situacao_clima.Fields.Item("dsc_situacao_clima").Value)%></option>
                    <%
                        rs_situacao_clima.MoveNext()
                        Wend
                        If (rs_situacao_clima.CursorType > 0) Then
                          rs_situacao_clima.MoveFirst
                        Else
                          rs_situacao_clima.Requery
                        End If
                    %>
                  </select>
                </td>
                <td>
                  <input type="text" name="dsc_nota_clima_tarde" placeholder="Notas">
                </td>
              </tr>
              <tr>
                <td>
                  Clima Noite:
                </td>
                <td>
                  <select name="cod_situacao_clima_noite" style="width: 100%;">
                    <option value=""></option>
                    <%
                      While (NOT rs_situacao_clima.EOF)
                    %>
                    <option value="<%=(rs_situacao_clima.Fields.Item("id").Value)%>"><%=(rs_situacao_clima.Fields.Item("dsc_situacao_clima").Value)%></option>
                    <%
                        rs_situacao_clima.MoveNext()
                        Wend
                        If (rs_situacao_clima.CursorType > 0) Then
                          rs_situacao_clima.MoveFirst
                        Else
                          rs_situacao_clima.Requery
                        End If
                    %>
                  </select>
                </td>
                <td>
                  <input type="text" name="dsc_nota_clima_noite" placeholder="Notas">
                </td>
              </tr>
              <tr>
                <td>
                  Limpeza da Obra:
                </td>
                <td>
                  <select name="cod_situacao_limpeza_obra" style="width: 100%;">
                    <option value=""></option>
                    <%
                      While (NOT rs_situacao_obra.EOF)
                    %>
                    <option value="<%=(rs_situacao_obra.Fields.Item("id").Value)%>"><%=(rs_situacao_obra.Fields.Item("dsc_situacao_obra").Value)%></option>
                    <%
                        rs_situacao_obra.MoveNext()
                        Wend
                        If (rs_situacao_obra.CursorType > 0) Then
                          rs_situacao_obra.MoveFirst
                        Else
                          rs_situacao_obra.Requery
                        End If
                    %>
                  </select>
                </td>
                <td>
                  <input type="text" name="dsc_nota_limpeza_obra" placeholder="Notas">
                </td>
              </tr>
              <tr>
                <td>
                  Organização da Obra:
                </td>
                <td>
                  <select name="cod_situacao_organizacao_obra" style="width: 100%;">
                    <option value=""></option>
                    <%
                      While (NOT rs_situacao_obra.EOF)
                    %>
                    <option value="<%=(rs_situacao_obra.Fields.Item("id").Value)%>"><%=(rs_situacao_obra.Fields.Item("dsc_situacao_obra").Value)%></option>
                    <%
                        rs_situacao_obra.MoveNext()
                        Wend
                        If (rs_situacao_obra.CursorType > 0) Then
                          rs_situacao_obra.MoveFirst
                        Else
                          rs_situacao_obra.Requery
                        End If
                    %>
                  </select>
                </td>
                <td>
                  <input type="text" name="dsc_nota_organizacao_obra" placeholder="Notas">
                </td>
              </tr>
            </table>
          </td>
        </tr>

        <%
          End If

          If Session("MM_UserAuthorization") <> 6 And Session("MM_UserAuthorization") <> 3 Then
        %>

        <!-- vistoria -->
        <tr valign="baseline">
          <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC">
            <strong class="style5">Foi Realizada a Vistoria?</strong>
          </td>
          <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
            <label>
              <input name="vistoria" type="checkbox" id="campo_vistoria" onclick="mostraEsconde('campo_vistoria','dta_vistoria');" value="1" />
            </label>
            <div id="dta_vistoria" style="display:none;">
              Data da Vistoria:
              <br/>
              <input name="dt_vistoria" type="text" id="dt_vistoria" class="datepicker" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" size="15" />
            </div>
          </td>
        </tr>

        <%
          End If
        %>

        <!-- pendencia -->
        <tr valign="baseline">
          <td align="right" nowrap bordercolor="#CCCCCC" bgcolor="#CCCCCC">
            <strong class="style5">É Pendência?</strong>
          </td>
          <td bordercolor="#CCCCCC" bgcolor="#CCCCCC">
            <label>
              <input name="flg_pendencia" type="checkbox" id="campo_pendencia" onclick="mostraEsconde('campo_pendencia','tipo_pendencia'); mostraEsconde('campo_pendencia','desc_pendencia'); mostraEsconde('campo_pendencia','dt_limite_pendencia');" value="1" />
            </label>
            <div id="tipo_pendencia" style="display:none;">
              Tipo de Pendência:
              <br/>
              <select name="cod_tipo_pendencia" style="width: 100%;">
                <option value=""></option>
                <%
                  Set objCon = Server.CreateObject("ADODB.Connection")
                      objCon.Open MM_cpf_STRING

                  strQ = "SELECT * FROM tb_tipo_pendencia "

                  Set rs_combo = Server.CreateObject("ADODB.Recordset")
                    rs_combo.CursorLocation = 3
                    rs_combo.CursorType = 3
                    rs_combo.LockType = 1
                    rs_combo.Open strQ, objCon, , , &H0001

                  If Not rs_combo.EOF Then
                    While Not rs_combo.EOF
                      If Trim(rs_combo.Fields.Item("dsc_tipo_pendencia").Value) <> "" Then
                %>
                <option value="<%=(rs_combo.Fields.Item("id").Value)%>"><%=(rs_combo.Fields.Item("dsc_tipo_pendencia").Value)%></option>
                <%
                      End If
                      rs_combo.MoveNext
                    Wend
                  End If
                %>
              </select>
            </div>
            <div id="desc_pendencia" style="display: none;">
              Descrição:
              <br/>
              <textarea name="dsc_pendencia" cols="32" maxlength="255"></textarea>
            </div>
            <div id="dt_limite_pendencia" style="display:none;">
              Data Limíte da Pendência:
              <br/>
              <input name="dta_limite_pendencia" type="text" id="dta_limite_pendencia" class="datepicker" onblur="verifica_data(this)" onkeypress="desabilita_cor(this)" onkeyup="this.value=mascara_data(this.value)" size="15" />
            </div>
          </td>
        </tr>

        <tr valign="baseline">
          <td colspan="2">
            <input type="submit" value="Salvar" style="float: right;">
          </td>
        </tr>
      </table>
      
      <input type="hidden" name="cod_usuario_lancamento" value="<%=(Session("MM_Userid"))%>">
      <input type="hidden" name="cod_tipo_registro" value="<%=(cod_tipo_registro)%>">
      <input type="hidden" name="PI" value="<%=(rs_pi.Fields.Item("PI").Value)%>">
      <input type="hidden" name="MM_insert" value="form1">
    </form>

    <%
      End If
    %>

    <br/>

    <table border="0" class="lista-acompanhamento">
      <tr bgcolor="#999999" class="title">
        <%
          If Session("MM_UserAuthorization") = 1 Or Session("MM_UserAuthorization") = 3 Or Session("MM_UserAuthorization") = 4 Or Session("MM_UserAuthorization") = 5 Then
        %>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <%
          End If
        %>
        <td>Upload de Arquivos (Máx. 2MB)</td>
        <%
          If Session("MM_UserAuthorization") = 1 or Session("MM_UserAuthorization") = 4 or Session("MM_UserAuthorization") = 5 Then
        %>
        <td>&nbsp;</td>
        <%
          End If
        %>
        <td style="min-width: 200px;">Tipo de Pendência</td>
        <td style="min-width: 300px;">Descrição Pendência</td>
        <td style="min-width: 80px;">Pendência?</td>
        <td style="min-width: 100px;">Data Limíte Pendência</td>
        <td style="min-width: 200px;">Respons&aacute;vel</td>
        <td style="min-width: 100px;">Data do Registro</td>
        <td style="min-width: 300px;">Registro</td>
        <%
          If Session("MM_UserAuthorization") = 1 or Session("MM_UserAuthorization") = 4 or Session("MM_UserAuthorization") = 5 Then
        %>
        <td style="min-width: 200px;">Situação MSST</td>
        <td style="min-width: 200px;">Notas/Observações MSST</td>
        <td style="min-width: 100px;">Clima Manhã</td>
        <td style="min-width: 300px;">Nota Clima Manhã</td>
        <td style="min-width: 100px;">Clima Tarde</td>
        <td style="min-width: 300px;">Nota Clima Tarde</td>
        <td style="min-width: 100px;">Clima Noite</td>
        <td style="min-width: 300px;">Nota Clima Noite</td>
        <td style="min-width: 100px;">Limpeza da Obra</td>
        <td style="min-width: 300px;">Nota Limpeza da Obra</td>
        <td style="min-width: 100px;">Organização da Obra</td>
        <td style="min-width: 300px;">Nota Organização da Obra</td>
        <%
          End If

          If Session("MM_UserAuthorization") <> 6 And Session("MM_UserAuthorization") <> 3 Then
        %>
        <td style="min-width: 80px;">Vistoria Realizada?</td>
        <td style="min-width: 100px;">Data Vistoria </td>
        <%
          End If
        %>
      </tr>
      <%
        While (NOT rs_lista_acomp.EOF)
      %>
      <tr bgcolor="#CCCCCC">
        <%
          If Session("MM_UserAuthorization") = 1 Or Session("MM_UserAuthorization") = 3 Or Session("MM_UserAuthorization") = 4 Or Session("MM_UserAuthorization") = 5 Then
        %>
        <td>
          <a href="altera_acomp.asp?cod_acompanhamento=<%=(rs_lista_acomp.Fields.Item("cod_acompanhamento").Value)%>">
            <img src="depto/imagens/edit.gif" width="16" height="15" border="0" />
          </a>
        </td>
        <td>
          <a href="delete_acomp.asp?pi=<%=(Request.QueryString("pi"))%>&cod_acompanhamento=<%=(rs_lista_acomp.Fields.Item("cod_acompanhamento").Value)%>">
            <img src="depto/imagens/delete.gif" width="16" height="15" border="0" />
          </a>
        </td>
        <%
          End If
        %>
        <td>
            <form id="form-upload" method="post" enctype="multipart/form-data" accept-charset="ISO-8859-1"
              action="novo_upload.asp?id=<%=(rs_lista_acomp.Fields.Item("cod_acompanhamento").Value)%>&folder=ACOMPANHAMENTO&retUrl=<%=(Request.ServerVariables("URL"))%>?<%=(Request.QueryString)%>">
              <input type="file" name="blob">
              <br/>
              <input type="text" name="dsc_observacoes" placeholder="Observações do arquivo">
              <input type="submit" id="btnSubmit" value="Upload">
            </form>

            <%
              cod_acompanhamento = rs_lista_acomp.Fields.Item("cod_acompanhamento").Value

              Set objCon = Server.CreateObject("ADODB.Connection")
                  objCon.Open MM_cpf_STRING

              strF = "SELECT * FROM tb_acompanhamento_arquivo WHERE cod_referencia = " & cod_acompanhamento

              Set rs_files = Server.CreateObject("ADODB.Recordset")
                rs_files.CursorLocation = 3
                rs_files.CursorType = 3
                rs_files.LockType = 1
                rs_files.Open strF, objCon, , , &H0001

              If Not rs_files.EOF Then
                While Not rs_files.EOF
                  fileid = rs_files.Fields.Item("id_arquivo").Value

                  If rs_files.Fields.Item("flg_pmweb_file").Value Then
                    filename = rs_files.Fields.Item("nme_arquivo").Value
                  Else
                    filename = rs_lista_acomp.Fields.Item("cod_acompanhamento").Value &"_"& rs_files.Fields.Item("nme_arquivo").Value
                  End If
            %>
              <ul>
                <li>
                  <a href="delete_file.asp?fileid=<%=(fileid)%>&foldername=ACOMPANHAMENTO&filename=<%=(filename)%>&returnurl=<%=(Request.ServerVariables("URL"))%>?<%=(Request.QueryString)%>">
                    <img src="depto/imagens/delete.gif" width="16" height="15" border="0" />
                  </a>
                  <a href="download.asp?path=/ARQUIVOS/ACOMPANHAMENTO&filename=<%=(filename)%>">
                    <%=(rs_files.Fields.Item("nme_arquivo").Value)%>
                  </a>
                </li>
              </ul>
            <%
                  rs_files.MoveNext
                Wend
              End If
            %>
        </td>
        <%
          If Session("MM_UserAuthorization") = 1 or Session("MM_UserAuthorization") = 4 or Session("MM_UserAuthorization") = 5 Then
        %>
        <td>
          <a href="cad_histograma.asp?cod_acompanhamento=<%=(rs_lista_acomp.Fields.Item("cod_acompanhamento").Value)%>&nome_municipio=<%=(rs_pi.Fields.Item("nme_municipio").Value)%>">
            Histograma
          </a>
        </td>
        <%
          End If
        %>
        <td class="text-center"><%=(rs_lista_acomp.Fields.Item("dsc_tipo_pendencia").Value)%></td>
        <td><%=(rs_lista_acomp.Fields.Item("dsc_pendencia").Value)%></td>
        <td class="text-center"><%=(rs_lista_acomp.Fields.Item("flg_pendencia").Value)%></td>
        <td class="text-center"><%=(rs_lista_acomp.Fields.Item("dta_limite_pendencia").Value)%></td>
        <td class="text-center"><%=(rs_lista_acomp.Fields.Item("nme_interessado").Value)%></td>
        <td class="text-center"><%=(rs_lista_acomp.Fields.Item("Data do Registro").Value)%></td>
        <td>
          <%
            If Not rs_lista_acomp.Fields.Item("Registro").Value = "" Then
              Response.Write Replace( rs_lista_acomp.Fields.Item("Registro").Value, VbCrLf, "<br>")
            End If
          %>
        </td>
        <%
          If Session("MM_UserAuthorization") = 1 or Session("MM_UserAuthorization") = 4 or Session("MM_UserAuthorization") = 5 Then
        %>
        <td><%=(rs_lista_acomp.Fields.Item("dsc_situacao_sso").Value)%></td>
        <td><%=(rs_lista_acomp.Fields.Item("dsc_nota_sso").Value)%></td>
        <td class="text-center"><%=(rs_lista_acomp.Fields.Item("clima_manha").Value)%></td>
        <td><%=(rs_lista_acomp.Fields.Item("dsc_nota_clima_manha").Value)%></td>
        <td class="text-center"><%=(rs_lista_acomp.Fields.Item("clima_tarde").Value)%></td>
        <td><%=(rs_lista_acomp.Fields.Item("dsc_nota_clima_tarde").Value)%></td>
        <td class="text-center"><%=(rs_lista_acomp.Fields.Item("clima_noite").Value)%></td>
        <td><%=(rs_lista_acomp.Fields.Item("dsc_nota_clima_noite").Value)%></td>
        <td class="text-center"><%=(rs_lista_acomp.Fields.Item("limpeza_obra").Value)%></td>
        <td><%=(rs_lista_acomp.Fields.Item("dsc_nota_limpeza_obra").Value)%></td>
        <td class="text-center"><%=(rs_lista_acomp.Fields.Item("organizacao_obra").Value)%></td>
        <td><%=(rs_lista_acomp.Fields.Item("dsc_nota_organizacao_obra").Value)%></td>
        <%
          End If

          If Session("MM_UserAuthorization") <> 6 And Session("MM_UserAuthorization") <> 3 Then
        %>
        <td class="text-center"><%=(rs_lista_acomp.Fields.Item("vistoria").Value)%></td>
        <td class="text-center"><%=(rs_lista_acomp.Fields.Item("dt_vistoria").Value)%></td>
        <%
          End If
        %>        
      </tr>
      <% 
          rs_lista_acomp.MoveNext()
        Wend
      %>
    </table>
  </body>
</html>
<%
rs_pi.Close()
Set rs_pi = Nothing
%>
<%
rs_fiscal.Close()
Set rs_fiscal = Nothing
%>
<%
rs_lista_acomp.Close()
Set rs_lista_acomp = Nothing
%>
<%
rs_situacao.Close()
Set rs_situacao = Nothing
%>