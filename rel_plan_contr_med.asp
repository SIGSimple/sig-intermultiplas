<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_cpf_STRING
Recordset1_cmd.CommandText = "SELECT tb_pi.PI, tb_predio.cod_predio, tb_predio.Nome_Unidade, tb_Municipios.Municipios, tb_pi.Órgão, tb_responsavel.Responsável, tb_pi.[Data da Abertura], tb_situacao_pi.desc_situacao, tb_pi.[Descrição da Intervenção Gerenciadora], Max(tb_Medicao_Construtora.n_medicao) AS ultima_medicao, tb_pi.[data da abertura]+tb_pi.[prazo do contrato] AS [Término Contratual], IIf([tb_pi].[data do trp]>0,100/100,[medicaoatual]) AS [% Avanço Físico Atual] FROM tb_responsavel RIGHT JOIN (((((tb_pi INNER JOIN tb_predio ON tb_pi.cod_predio = tb_predio.cod_predio) INNER JOIN tb_situacao_pi ON tb_pi.cod_situacao = tb_situacao_pi.cod_situacao) INNER JOIN tb_Municipios ON tb_predio.Município = tb_Municipios.Municipios) LEFT JOIN tb_Medicao_Construtora ON tb_pi.PI = tb_Medicao_Construtora.PI) LEFT JOIN c_Semaforico ON tb_pi.PI = c_Semaforico.[PI-item]) ON tb_responsavel.cod_fiscal = tb_pi.cod_fiscal GROUP BY tb_pi.PI, tb_predio.cod_predio, tb_predio.Nome_Unidade, tb_Municipios.Municipios, tb_pi.Órgão, tb_responsavel.Responsável, tb_pi.[Data da Abertura], tb_situacao_pi.desc_situacao, tb_pi.[Descrição da Intervenção Gerenciadora], tb_pi.[data da abertura]+tb_pi.[prazo do contrato], IIf([tb_pi].[data do trp]>0,100/100,[medicaoatual]) ORDER BY tb_pi.PI, tb_predio.cod_predio; " 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 9px;
}
.style11 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10px;
	color: #FFFFFF;
}
.style13 {
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style15 {font-family: Arial, Helvetica, sans-serif; font-size: 8px; }
.style17 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; color: #FFFFFF; }
.style19 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
-->
</style>
<script language="JavaScript">
<!--
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
// -->
</script>
</head>

<body onload="FP_preloadImgs(/*url*/'button199.jpg', /*url*/'button203.jpg')">
<p align="center" class="style13"><U>RELATÓRIO DE PLANEJAMENTO E CONTROLE DAS MEDIÇÕES</U></p>
<p align="center" class="style13"><a href="saida_excel_rel_plan.asp">
<img border="0" id="img1" src="button204.jpg" height="23" width="116" alt="Exportar p/ Excel" fp-style="fp-btn: Embossed Capsule 5" fp-title="Exportar p/ Excel" onmouseover="FP_swapImg(1,0,/*id*/'img1',/*url*/'button199.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img1',/*url*/'button204.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img1',/*url*/'button203.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img1',/*url*/'button199.jpg')"></a></p>
<table border="0">
  <tr bgcolor="#999999">
    <td><span class="style17">PI</span></td>
    <td><span class="style17">cod_predio</span></td>
    <td><span class="style17">Nome_Unidade</span></td>
    <td><span class="style17">Municipio</span></td>
    <td class="style11">Tipo de obra gerenciadora</td>
    <td><span class="style17">Órgão</span></td>
    <td><span class="style17">Fiscal</span></td>
    <td><span class="style17">Data de Abertura</span></td>
    <td><span class="style17">Situação</span></td>
    <td><span class="style17">Última Medição</span></td>
    <td><span class="style17">Término Contratual</span></td>
    <td><span class="style17">% Avanço Físico</span></td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) %>
    <tr bgcolor="#F3F3F3" class="style19">
      <td><span class="style19"><%=(Recordset1.Fields.Item("PI").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("cod_predio").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("Nome_Unidade").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("Municipios").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("Descrição da Intervenção Gerenciadora").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("Órgão").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("Responsável").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("Data da Abertura").Value)%></span></td>
      <td><span class="style19"><%=(Recordset1.Fields.Item("desc_situacao").Value)%></span></td>
      <td><%=(Recordset1.Fields.Item("ultima_medicao").Value)%></td>
      <td><%=(Recordset1.Fields.Item("Término Contratual").Value)%></td>
      <td><%= FormatPercent((Recordset1.Fields.Item("% Avanço Físico Atual").Value), 2, -2, -2, -2) %></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>