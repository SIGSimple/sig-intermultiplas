﻿<%

Session.LCID = 1046

' FileName="Connection_ado_conn_string.htm"
' Type="ADO"
' DesigntimeType="ADO"
' HTTP="false"
' Catalog=""
' Schema=""
Dim isInDevelopment
	isInDevelopment = False

Dim userFriendlyMessage
	userFriendlyMessage = "<strong>Caro usuário, o sistema está temporáriamente fora do ar, devido atualizações de dados.</strong><br/><br/>Tente acessar novamente mais tarde!"

Dim MM_cpf_STRING
	'MM_cpf_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=e:\home\programaagualimpa\web\ARQUIVOS\DADOS\bd_fde.mdb"
	MM_cpf_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\10.0.75.124\intermultiplas.net\public\ARQUIVOS\DADOS\bd_fde.mdb"
	'MM_cpf_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\inetpub\wwwroot\inter\ARQUIVOS\DADOS\bd_fde.mdb"

If isInDevelopment Then
	MM_cpf_STRING = ""
End If

%>
