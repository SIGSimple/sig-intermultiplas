<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<%
Dim Recordset1
Dim objCon
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open MM_cpf_STRING

strQ = "SELECT tb_predio.cod_predio, tb_predio.Município FROM tb_predio RIGHT JOIN tb_PI ON tb_predio.cod_predio = tb_PI.cod_predio GROUP BY tb_predio.cod_predio, tb_predio.Município ORDER BY tb_predio.Município"

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.CursorLocation = 3
Recordset1.CursorType = 3
Recordset1.LockType = 1
Recordset1.Open strQ, objCon, , , &H0001

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>Untitled Document</title>
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
<style type="text/css">
<!--
.style9 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; }
-->
</style>
<!--mstheme--><link rel="stylesheet" href="spri1011-28591.css">
<meta name="Microsoft Theme" content="spring 1011">
</head>

<body onload="FP_preloadImgs(/*url*/'button52.jpg', /*url*/'button53.jpg')">
<p align="center"><strong><span class="style17">Acompanhamento de Obra</span></strong></p>
<form id="form1" name="form1" method="post" action="filtro_exibir_adm.asp?cod_predio=<%=(Recordset1.Fields.Item("cod_predio").Value)%>">
  <label>
  <div align="center">
    <select name="cod_predio" class="style9" id="cod_predio">
	<option value=""></value>
      <%
While (NOT Recordset1.EOF)
%>
      <option value="<%=(Recordset1.Fields.Item("cod_predio").Value)%>"><%=(Recordset1.Fields.Item("Município").Value)%></option>
      <%
  Recordset1.MoveNext()
Wend
If (Recordset1.CursorType > 0) Then
  Recordset1.MoveFirst
Else
  Recordset1.Requery
End If
%>
    </select>
    <input type="submit" name="Submit" value="Ok" />
  <a href="cadastro_pi.asp"></a></div>
  </label>
  <label>
  <div align="center"></div>
  </label>
</form>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>