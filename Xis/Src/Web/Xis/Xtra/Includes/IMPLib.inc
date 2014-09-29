<%

'Sub that builds and send mail that informs the customer of
'their user acount info - Userid, password etc.
sub SendmailPWDIUD(IFirmID,strpassword,strType)
	dim strSubject  'as string
	dim strBody  'as string
	dim strRemotehost 'as string
	dim strFromname  'as string
	dim strFromAdress 'as string
	dim strToAdress 'as string
	dim StrToName 'as string
	dim StrNavn 'as string

	if strType = "KUNDE" then
		StrNavn = Request.form("tbxFirstName") & " " & Request.form("tbxLastName")	
	elseif strType = "ANSATT" then
		StrNavn = Request("tbxFornavn") & " " & Request("tbxEtternavn")
	End if

	strSubject = Application("MailSubjectNewUser")
	strBody = "Til " & strNavn & vbCrLf & vbCrLf & "Du er nå opprettet som bruker på Xtra's web-sider." & vbCrLf & vbCrLf
	strBody = strBody & Application("MailBodyNewUser")
	strBody = strBody & vbCrLf & "Ditt brukernavn er '" & trim(Request("Username"))& "'."
	strBody = strBody & vbCrLf & "Ditt passord er '" & strPassword & "'."
	strBody = strBody & vbCrLf & vbCrLf & "Dersom du ikke er " & strNavn & " ber vi deg svare på denne e-posten snarest og si ifra om dette."
	strBody = strBody & vbCrLf & vbCrLf & Application("MailBodySign")
	strRemotehost = Application("XtraMailServer")
	strFromname  = Application("XtraSenderName")
	strFromAdress = Application("XtraSenderMail")
	strToAdress = Request.form("tbxEPost")

	call SendCustomerMail(IFirmID,strFromAdress,StrFromName,strSubject,strBody,strRemoteHost,strToAdress,StrToName)
end sub


sub SendUpdatedUID(IVikarID,strType)
	dim strSubject  'as string
	dim strBody  'as string
	dim strRemotehost 'as string
	dim strFromname  'as string
	dim strFromAdress 'as string
	dim strToAdress 'as string
	dim StrToName 'as string
	dim StrNavn 'as string
	
	if strType = "KUNDE" then
		StrNavn = Request.form("tbxFirstName") & " " & Request.form("tbxLastName")	
	elseif strType = "ANSATT" then
		StrNavn = Request("tbxFornavn") & " " & Request("tbxEtternavn")
	End if
	
	strSubject = Application("MailSubjectNewUID")
	strBody = "Til " & strNavn & vbCrLf & vbCrLf & "Du har fått et nytt brukernavn til Xtra's web-sider. Passordet ditt er uendret. " & vbCrLf & vbCrLf
	strBody = strBody & Application("MailBodyNewUID") & "'" & trim(Request("Username")) & "'."
	strBody = strBody & vbCrLf & vbCrLf & "Dersom du ikke er " & strNavn & " ber vi deg svare på denne e-posten snarest og si ifra om dette."
	strBody = strBody & vbCrLf & vbCrLf & Application("MailBodySign")
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

%>