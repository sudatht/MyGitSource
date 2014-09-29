<%

'Sub that builds and send mail that informs the customer of
'their user acount info - Userid, password etc.
sub SendmailPWDIUD(IFirmID, strpassword, strType)
	dim strSubject
	dim strBody
	dim strRemotehost
	dim strFromname
	dim strFromAdress
	dim strToAdress
	dim StrToName
	dim StrNavn
	dim emailText

	if strType = "KUNDE" then
		StrNavn = Request.form("tbxFirstName") & " " & Request.form("tbxLastName")
	elseif strType = "ANSATT" then
		StrNavn = Request("tbxFornavn") & " " & Request("tbxEtternavn")
	End if

	emailText = getsetting("Mail_PWDUID")

	emailText = replace(emailText, "%NAME%", strNavn)
	emailText = replace(emailText, "%PASSWORD%", strPassword)
	emailText = replace(emailText, "%USERNAME%", trim(Request("Username")) )
	emailText = replace(emailText, "%MAILBODYSIGN%", Application("MailBodySign") )

	strSubject = Application("MailSubjectNewUser")
	strBody = emailText
	strRemotehost = Application("XtraMailServer")
	strFromname  = Application("XtraSenderName")
	strFromAdress = Application("XtraSenderMail")
	strToAdress = Request.form("tbxEPost")

	call SendCustomerMail(IFirmID, strFromAdress, StrFromName, strSubject, strBody, strRemoteHost, strToAdress, strNavn)
end sub

sub SendmailUPDATECV(IFirmID, strpassword)
	dim strSubject
	dim strBody
	dim strRemotehost
	dim strFromname
	dim strFromAdress
	dim strToAdress
	dim StrToName
	dim StrNavn
	dim emailText

	StrNavn = Request("tbxFornavn") & " " & Request("tbxEtternavn")

	emailText = getsetting("Mail_UIDPWD_EDITCV")
	emailText = replace(emailText, "%NAME%", strNavn)
	emailText = replace(emailText, "%PASSWORD%", strPassword)
	emailText = replace(emailText, "%USERNAME%", trim(Request("Username")) )
	emailText = replace(emailText, "%MAILBODYSIGN%", Application("MailBodySign") )

	strSubject = Application("MailSubjectNewUser")
	strBody = emailText
	strRemotehost = Application("XtraMailServer")
	strFromname  = Application("XtraSenderName")
	strFromAdress = Application("XtraSenderMail")
	strToAdress = Request.form("tbxEPost")

	call SendCustomerMail(IFirmID, strFromAdress, StrFromName, strSubject, strBody, strRemoteHost, strToAdress, StrToName)
end sub

sub SendUpdatedUID(IVikarID, strType)
	dim strSubject
	dim strBody
	dim strRemotehost
	dim strFromname
	dim strFromAdress
	dim strToAdress
	dim StrToName
	dim StrNavn
	dim emailText

	if strType = "KUNDE" then
		StrNavn = Request.form("tbxFirstName") & " " & Request.form("tbxLastName")
	elseif strType = "ANSATT" then
		StrNavn = Request("tbxFornavn") & " " & Request("tbxEtternavn")
	End if

	emailText = getsetting("Mail_UID")
	emailText = replace(emailText, "%NAME%", strNavn)
	emailText = replace(emailText, "%USERNAME%", trim(Request("Username")) )
	emailText = replace(emailText, "%MAILBODYSIGN%", Application("MailBodySign") )

	strSubject = Application("MailSubjectNewUID")
	strBody = emailText
	strRemotehost = Application("XtraMailServer")
	strFromname  = Application("XtraSenderName")
	strFromAdress = Application("XtraSenderMail")
	strToAdress = Request.form("tbxEPost")

	call SendCustomerMail(IVikarID,strFromAdress,StrFromName,strSubject,strBody,strRemoteHost,strToAdress,StrToName)
end sub

'Check for duplicate usernames
function VerifyUserName(strUserName, XtraID, XtraType)
	'Variables used by IM Publish (IMP) user routines.
	dim rsUser 'Recordset
	dim StrSQL 'as string
	dim ConIMP 'Connection

	'Initialize ADO objects
	Set rsUser = CreateObject("ADODB.recordset")
	Set conIMP = CreateObject("ADODB.Connection")
	conIMP.CursorLocation = 3
	conIMP.Open application("IMPconnection")

	VerifyUserName = true

	StrSQL = "select t1.type, t1.Xtra_id from imp_xtra_users t1, users t2 " &_
	"where lower(t2.username) = '" & StrUserName & "' " &_
	"and t1.userid=t2.userid "


	'response.write strsql & "<br>"

	'Check if username allready exists..
	set RsUser = conIMP.execute(strSQL)
	set RsUser.Activeconnection = nothing
	set conIMP = nothing
	'New user
	if Xtraid = 0 then
		if not RsUser.EOF then
			VerifyUserName = false
		end if
	else
		if not RsUser.EOF then
			if ((RsUser.fields("type").value <> XtraType) and (RsUser.fields("xtra_id").value = XtraID)) then
				VerifyUserName = false
			elseif ((RsUser.fields("type").value = XtraType) and (RsUser.fields("xtra_id").value <> XtraID)) then
				VerifyUserName = false
			elseif ((RsUser.fields("type").value <> XtraType) and (RsUser.fields("xtra_id").value <> XtraID)) then
				VerifyUserName = false
			end if

		end if
	end if
	set rsUser = nothing
	set conIMP = nothing
end function

'---------------------------------------------------------------------------------------------
'Fred 03.08.2000
sub SendCustomerMail(IFirmID,strFromAdress,StrFromName,strSubject,strBody,strRemoteHost,strToAdress,StrToName)
	dim ConIMP 'Connection
	dim rsUser 'Recordset

	'Initialize ADO objects
	'Set conIMP = CreateObject("ADODB.Connection")
	'Set rsUser = CreateObject("ADODB.recordset")

	'Open connection to IMP base
	'conIMP.Open application("IMPconnection")

	'Updates sentmail coloumn with date and time if sucessfully send.

	if sendmail(strFromAdress,StrFromName,strSubject,strBody,strRemoteHost,strToAdress,StrToName) then
		'ConIMP.Execute "update imp_xtra_users set sentmail=getdate() where Type='KUNDE' and xtra_id=" & IFirmID
	else
		'ConIMP.Execute "update imp_xtra_users set sentmail=NULL where Type='KUNDE' and xtra_id=" & IFirmID
	end if
	'Clean up
	set rsUser = nothing
	set conIMP = nothing
end sub

'---------------------------------------------------------------------------------------------
'Fred 03.08.2000
'Generic mailer returns TRUE if mail is succesfullt sent, FALSE otherwise.
'If the strToadress is invalid then the mail rebounds to strFromAdress @ strRemoteHost.

function sendMail(strFromAdress, StrFromName, strSubject, strBody, strRemoteHost, strToAdress, StrToName)
	'Initiate Serverobjects' ASPMail object
	Set ObjMailer = Server.CreateObject("SMTPsvg.Mailer")
	'ObjMailer.SMTPLog = "D:\W3Server\Logs\mail.txt"
	'response.end
	if IsObject(ObjMailer) then
		ObjMailer.FromName   = StrFromName
		ObjMailer.FromAddress = strFromAdress
		ObjMailer.RemoteHost  = strRemoteHost
		ObjMailer.AddRecipient StrToName, strToAdress
		ObjMailer.Subject    = strSubject
		ObjMailer.BodyText   = strBody
		ObjMailer.CharSet 	 = 2
		if ObjMailer.SendMail then
		  sendMail = true
		else
		  sendMail = false
		end if
	else
		sendMail = false
	end if
	set ObjMailer = nothing
end function

%>
