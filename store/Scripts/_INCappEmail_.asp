<%
'*************************************************************************
' DO NOT MODIFY THIS SCRIPT IF YOU WANT UPDATES TO WORK!
' Function : Sends email
' Product  : CandyPress eCommerce Storefront
' Version  : 2.2
' Modified : November 2002
' Copyright: Copyright (C) 2002 CandyPress.Com 
'            See "license.txt" for this product for details regarding 
'            licensing, usage, disclaimers, distribution and general 
'            copyright requirements. If you don't have a copy of this 
'            file, you may request one at webmaster@candypress.com
'*************************************************************************

'*************************************************************************
'Forward email send request to appropriate component
'*************************************************************************
Function sendMail(fromName, fromEmail, toEmail, subject, body, contType)

	on error resume next
	
	select case mailComp
	
		case "1" 'JMail
			call JMail(fromName, fromEmail, toEmail, subject, body, contType)
			
		case "2" 'CDONTS
			call CDONTS(fromName, fromEmail, toEmail, subject, body, contType)
			
		case "3" 'Persits ASPEmail
			call PASPEmail(fromName, fromEmail, toEmail, subject, body, contType)
			
		case "4" 'ServerObjects ASPMail
			call SOASPMail(fromName, fromEmail, toEmail, subject, body, contType)
			
		case "5" 'Bamboo SMTP
			call bamboo(fromName, fromEmail, toEmail, subject, body, contType)
			
		case "6" 'CDOSYS
			call CDOSYS(fromName, fromEmail, toEmail, subject, body, contType)
			
	end select
	
	on error goto 0
	
end Function

'JMail
Function JMail(fromName, fromEmail, toEmail, subject, body, contType)
	dim mail,I
	
	'Version 4.3
	'------------------------------------------------------------
	set mail			= server.CreateObject("JMail.Message")
	mail.From			= fromEmail
	mail.FromName		= fromName
	'------------------------------------------------------------
	
	'Version 3.7
	'------------------------------------------------------------
	'set mail			= server.CreateObject("JMail.SMTPMail")
	'mail.ServerAddress	= pSMTPServer
	'mail.Sender		= fromEmail
	'mail.SenderName	= fromName
	'------------------------------------------------------------

	mail.silent   = true
	mail.Subject  = subject
	mail.Body     = body
	if contType = 1 then			'Send HTML Email
		mail.ContentType = "text/html"
	end if
	if isArray(toEmail) then		'Send Multiple Emails
		for I = 0 to Ubound(toEmail)
			if len(toEmail(I)) > 0 then
				mail.ClearRecipients()
				mail.AddRecipient toEmail(I)
				mail.Send(pSmtpServer)	'V4.3
				'mail.Execute			'V3.7
			end if
		next
	else							'Send Single Email
		mail.AddRecipient toEmail
		mail.Send(pSmtpServer)			'V4.3
		'mail.Execute					'V3.7
	end if
	set mail = nothing
end Function

'CDONTS
'Note : After the "Send" method, the "NewMail" object becomes invalid.
'       We therefore have to create the "NewMail" object for each email.
Function CDONTS(fromName, fromEmail, toEmail, subject, body, contType)
	dim mail,I
	if isArray(toEmail) then		 'Send Multiple Emails
		for I = 0 to Ubound(toEmail)
			if len(toEmail(I)) > 0 then
				Set mail = Server.CreateObject ("CDONTS.NewMail")
				if contType = 1 then 'Send HTML Email
					mail.BodyFormat = 0
					mail.MailFormat = 0
				end if
				mail.Send fromEmail & " (" & fromName & ")", toEmail(I), subject, body
				set mail = nothing
			end if
		next
	else							 'Send Single Email
		Set mail = Server.CreateObject ("CDONTS.NewMail")
		if contType = 1 then		 'Send HTML Email
			mail.BodyFormat = 0
			mail.MailFormat = 0
		end if
		mail.Send fromEmail & " (" & fromName & ")", toEmail, subject, body
		set mail = nothing
	end if
end Function

'Persits ASP Email
Function PASPEmail(fromName, fromEmail, toEmail, subject, body, contType)
	dim mail,I
	set mail 	  = server.CreateObject("Persits.MailSender")
	mail.Host 	  = pSmtpServer 
	mail.From 	  = fromEmail
	mail.FromName = fromName
	mail.Subject  = subject
	mail.Body 	  = body
	if contType = 1 then			'Send HTML Email
		mail.IsHTML = True 
	end if
	if isArray(toEmail) then		'Send Multiple Emails
		for I = 0 to Ubound(toEmail)
			if len(toEmail(I)) > 0 then
				mail.Reset
				mail.AddAddress toEmail(I)
				mail.Send
			end if
		next
	else							'Send Single Email
		mail.AddAddress toEmail
		mail.Send
	end if
	set mail = nothing
end Function

'ServerObjects ASPMail
Function SOASPMail(fromName, fromEmail, toEmail, subject, body, contType)
	dim mail,I
	set mail		 = server.CreateObject("SMTPsvg.Mailer")
	mail.RemoteHost	 = pSmtpServer
	mail.FromAddress = fromEmail
	mail.FromName	 = fromName
	mail.Subject	 = subject
	mail.BodyText	 = body
	if contType = 1 then			'Send HTML Email
		mail.ContentType = "text/html"
	end if
	if isArray(toEmail) then		'Send Multiple Emails
		for I = 0 to Ubound(toEmail)
			if len(toEmail(I)) > 0 then
				mail.ClearRecipients
				mail.AddRecipient "", toEmail(I)
				mail.SendMail
			end if
		next
	else							'Send Single Email
		mail.AddRecipient "", toEmail
		mail.SendMail
	end if
	set mail = nothing
end Function

'Bamboo SMTP
Function bamboo(fromName, fromEmail, toEmail, subject, body, contType)
	dim mail,I
	set mail      = Server.CreateObject("Bamboo.SMTP") 
	mail.Server   = pSmtpServer
	mail.From     = fromEmail
	mail.FromName = fromName
	mail.Subject  = subject
	mail.Message  = body
	if contType = 1 then			'Send HTML Email
		mail.ContentType = "text/html"
	end if
	if isArray(toEmail) then		'Send Multiple Emails
		for I = 0 to Ubound(toEmail)
			if len(toEmail(I)) > 0 then
				mail.Rcpt = toEmail(I)
				mail.Send 
			end if
		next
	else							'Send Single Email
		mail.Rcpt = toEmail
		mail.Send 
	end if
	set mail      = nothing
end Function

'CDOSYS
Function CDOSYS(fromName, fromEmail, toEmail, subject, body, contType)
	dim mail,conf,flds,I
	
	Set mail = Server.CreateObject("CDO.Message")
	Set conf = Server.CreateObject("CDO.Configuration")
	Set flds = conf.Fields

	'Configure the server
	flds("http://schemas.microsoft.com/cdo/configuration/smtpserver") = pSmtpServer
	flds("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	flds("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	flds("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	flds.Update
	mail.Configuration = conf
	
	'Message
	mail.From = fromName & " <" & fromEmail & ">"
	mail.Subject = subject
	if contType = 1 then			'Send HTML Email
		mail.HTMLBody = body
	else
		mail.TextBody = body
	end if
	if isArray(toEmail) then		'Send Multiple Emails
		for I = 0 to Ubound(toEmail)
			if len(toEmail(I)) > 0 then
				mail.To = toEmail(I)
				mail.Send
			end if
		next
	else							'Send Single Email
		mail.To = toEmail
		mail.Send
	end if
	
	'Clean up
	set mail = nothing
	set flds = nothing
	set conf = nothing

end Function
%>