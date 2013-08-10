<%
function sendMail(pFromAddress, pFromName, pRecipient, pRecipientName, pSubject, pBody)
	on error resume next
	' Geocel DevMailer 1.51
	' VBScript Usage Example
	' (c) 1999, Geocel International, Inc.

	' Creamos el Objeto DevMailer
	 Set Mailer = CreateObject("Geocel.Mailer")

	 ' Añadir el primer servidor SMTP
	 Mailer.AddServer "217.76.145.56",25

	 ' Tipo de contenido
	 Mailer.ContentType = "text/html"

	' Set Sender Information
	Mailer.FromAddress = pFromAddress
	Mailer.FromName = pFromName

	' Add a recipient to the message
	Mailer.AddRecipient pRecipient, pRecipientName

	' Set the Subject and Body
	Mailer.Subject = pSubject
	Mailer.Body =	pBody

	' Send Email - Perform Error Checking
	sendMail = Mailer.Send()

end function


%>