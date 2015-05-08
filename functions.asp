<%
	Function IIf(bClause, sTrue, sFalse)
		If CBool(bClause) Then
			IIf = sTrue
		Else
			IIf = sFalse
		End If
	End Function

	Function CaptalizeText(Text)
		Temp = Split(Text, " ")
		
		For i = LBound(Temp) to UBound(Temp)
			TextTemp = TextTemp & " " & UCase(Left(Temp(i), 1)) & Mid(Temp(i), 2)
		Next
		
		CaptalizeText = TextTemp
	End Function

	Function HTMLEspeciais(sString)
		If (sString <> "") Then
			sString = Replace(sString, "á", "&aacute;")
			sString = Replace(sString, "â", "&acirc;")
			sString = Replace(sString, "à", "&agrave;")
			sString = Replace(sString, "ã", "&atilde;")

			sString = Replace(sString, "ç", "&ccedil;")

			sString = Replace(sString, "é", "&eacute;")
			sString = Replace(sString, "ê", "&ecirc;")

			sString = Replace(sString, "í", "&iacute;")

			sString = Replace(sString, "ó", "&oacute;")
			sString = Replace(sString, "ô", "&ocirc;")
			sString = Replace(sString, "õ", "&otilde;")

			sString = Replace(sString, "ú", "&uacute;")
			sString = Replace(sString, "ü", "&uuml;")

			sString = Replace(sString, "Á", "&Aacute;")
			sString = Replace(sString, "Â", "&Acirc;")
			sString = Replace(sString, "À", "&Agrave;")
			sString = Replace(sString, "Ã", "&Atilde;")

			sString = Replace(sString, "Ç", "&Ccedil;")

			sString = Replace(sString, "É", "&Eacute;")
			sString = Replace(sString, "Ê", "&Ecirc;")

			sString = Replace(sString, "Í", "&Iacute;")

			sString = Replace(sString, "Ó", "&Oacute;")
			sString = Replace(sString, "Ô", "&Ocirc;")
			sString = Replace(sString, "Õ", "&Otilde;")

			sString = Replace(sString, "Ú", "&Uacute;")
			sString = Replace(sString, "Ü", "&Uuml;")

			sString = Replace(sString, """", "&quot;") '"
			sString = Replace(sString, "<", "&lt;") '<
			sString = Replace(sString, ">", "&gt;") '>
		End If

		HTMLEspeciais = sString
	End Function

	Function URLDecode(sConvert)
		Dim aSplit
		Dim sOutput
		Dim I
		If IsNull(sConvert) Then
		   URLDecode = ""
		   Exit Function
		End If

		' convert all pluses to spaces
		sOutput = REPLACE(sConvert, "+", " ")

		' next convert %hexdigits to the character
		aSplit = Split(sOutput, "%")

		If IsArray(aSplit) Then
		  sOutput = aSplit(0)
		  For I = 0 to UBound(aSplit) - 1
		    sOutput = sOutput & _
		      Chr("&H" & Left(aSplit(i + 1), 2)) &_
		      Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
		  Next
		End If

		URLDecode = sOutput
	End Function

	Function EnviaEmail(ToEmail, Subject, HTMLTextBody)
		' Cria o objeto CDOSYS
		Set objCDOSYSMail = Server.CreateObject("CDO.Message")

		'Cria o objeto para configuração do SMTP
		Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")

		'Servidor SMTP que será utilizado para enviar o e-mail
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"

		'Porta do SMTP
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")= 25

		'Porta do CDOSYS
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

		'Timeout de conexão com o Servidor SMTP
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
		objCDOSYSCon.Fields.update

		'Atualiza a configuração do CDOSYS para envio do e-mail
		Set objCDOSYSMail.Configuration = objCDOSYSCon

		'Configura o remetente do e-mail
		objCDOSYSMail.From = "sender@intermultiplas.net"

		'Configura o destinatário(TO)
		objCDOSYSMail.To = ToEmail

		'Configura o Reply-To(Responder Para) 
		objCDOSYSMail.ReplyTo = "no-reply@intermultiplas.net"

		'Configura o assunto(SUBJECT)
		objCDOSYSMail.Subject = Subject

		'Para definir o charset da mensagem
		objCDOSYSMail.BodyPart.Charset = "utf-8"

		'Para enviar mensagens no formato HTML
		objCDOSYSMail.HtmlBody = HtmlBody

		' ### ENVIA O E-MAIL ###
		objCDOSYSMail.Send

		Set objCDOSYSMail = Nothing
		Set objCDOSYSCon = Nothing
	End Function
%>