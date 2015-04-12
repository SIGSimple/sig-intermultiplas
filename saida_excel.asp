<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
</head>

<body>
<% 
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "filename =  "&date&"-semaforico.xls"

'DIM Conteudo varchar(50)
'SET Conteudo = convert(varchar(50),getdate(),105) + ' somaforico.xls'
'Response.AddHeader "Content-Disposition", "filename = "  + Conteudo
 


set objconn=server.createobject("adodb.connection")

connpath= "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\inetpub\wwwroot\original\ARQUIVOS\DADOS\bd_fde.mdb"
objconn.Open connpath 

set objrs=objconn.execute("SELECT c_Semaforico1.*, tb_situacao_pi.desc_situacao " & _
	"FROM c_Semaforico1 INNER JOIN (tb_situacao_pi INNER JOIN tb_pi ON tb_situacao_pi.cod_situacao = tb_pi.cod_situacao) ON c_Semaforico1.[PI-item] = tb_pi.PI ")




%>
<TABLE BORDER=1>
<TR>
<% 
'Percorre cada campo e imprime o nome dos campos da tabela
For i = 0 to objrs.fields.count - 1 
%>
<TD><% = objrs(i).name %></TD>
<% next %>
</TR>
<% 

'Percorre cada linha e exibe cada campo da tabela

while not objrs.eof
%>
<TR>
<% For i = 0 to objrs.fields.count - 1
%>
<TD VALIGN=TOP><% = objrs(i) %></TD>
<% Next %>
</TR>
<%
objrs.MoveNext

wend

objrs.Close
objconn.close
%>
</TABLE> 

</body>
</html>
