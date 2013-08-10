<%
function sendMail(pFromAddress, pFromName, pRecipient, pRecipientName, pSubject, pBody)
	On Error Resume Next
	Set Mail = Server.CreateObject("Persits.MailSender")
'	Mail.Host = "82.98.139.38" ' IP Dinahosting
	Mail.Host = "smtp.brmusic.net"
	Mail.IsHTML = true
	Mail.From = pFromAddress
	Mail.FromName = pFromName
	Mail.AddAddress pRecipient, pRecipientName
	Mail.Username = "brmusic01"
	Mail.Password = "musica"
	Mail.Subject = pSubject

	Mail.Body = pBody
	Mail.Send
	If Err <> 0 Then
		sendMail = Err.Description
	else
		sendMail = ""
	End If
	Set Mail = nothing
	on error goto 0
end function


%>