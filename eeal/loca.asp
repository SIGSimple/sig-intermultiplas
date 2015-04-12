<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
<script> 
function abre_painel(){ 
            if (document.painel_webmail.form_login.value==''){ 
                alert ("Por favor, informe o Usuário !"); 
                return false; 
            } 
            if (document.painel_webmail.form_passwd.value==''){ 
                alert ("Por favor, informe a Senha !"); 
                return false; 
            }

document.painel_webmail.action = 'https://locamail.locaweb.com.br/locamail/'; 
return true; 
} 
</script> 

<html> 
<body> 
Formulário para acesso ao Locamail Admin <br> <br> 
<form name="painel_webmail" method="post" action="" target="_self" onSubmit="return abre_painel()"> 
Usuário: <input name="form_login" type="text"> <br> 
<input type="hidden" name="A" value="checkin"> 
Senha: <input name="form_passwd" type="password"> <br> 
<input name="ok" type="submit" value="OK"> 
</form> 
</body> 
</html> 

</body>
</html>
