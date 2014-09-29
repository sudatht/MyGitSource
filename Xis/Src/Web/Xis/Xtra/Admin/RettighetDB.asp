<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<%
dim strProfil
'DB specific variables
dim Conn
'Selected User specific variables
dim strID
dim strNavn
dim strBrukerID
dim HowMany
dim bSkattekort
'dim bKundegrupper
dim bFAgreementHandler
dim bDeletePerson
dim bQuestback

If  HasUserRight(ACCESS_ADMIN, RIGHT_SUPER) Then 
	' Connect to database
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
	Conn.CommandTimeOut = Session("xtra_CommandTimeout")
	Conn.Open Session("xtra_ConnectionString"),Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

	' check parameters
	strID = Request.Querystring("ID")
	strNavn = Request.Querystring("Navn")
	strBrukerID = Request.Querystring("BrukerID")
	HowMany = Request.Form("HowMany")

	If (Request.Form("skattekort") <> "") Then
		bSkattekort = 1
	Else
		bSkattekort = 0
	End If

	'If (Request.Form("kundegruppe") <> "") Then
		'bKundegrupper = 1
	'Else
		'bKundegrupper = 0
	'End If
	
	If (Request.Form("fagreement") <> "") Then
		bFAgreementHandler = 1
	Else
		bFAgreementHandler = 0
	End If

	If (Request.Form("deleteperson") <> "") Then
		bDeletePerson = 1
	Else
		bDeletePerson = 0
	End If
	
	If (Request.Form("questback") <> "") Then
		bQuestback = 1
	Else
		bQuestback = 0
	End If
	
	strProfil = ""

	For i = 1 To HowMany
		strProfil = strProfil & Request.Form("HovedMeny" & i)
	Next

	' sql for update of profile
	strSQL = "UPDATE [Bruker] " & _
			"SET [Profil] = '" & strProfil & "', " & _
			"[SkattekortEndringer] = '" & bSkattekort & "', " & _
			"[FAgreementHandler] = '" & bFAgreementHandler & "', " & _
			"[DeletePersonInfo] = '" & bDeletePerson & "', " & _
			"[Questback] = '" & bQuestback & "', " & _
			"[KundeGruppeEndringer] = '0' " & _			
			"WHERE [ID] = '" & strID & "' "

	Conn.Execute(strSQL)
	Conn.Close
	set Conn = nothing
	
	'--------------------------------------------------------------------------------------------------
	' display rights updated rights
	'--------------------------------------------------------------------------------------------------
		Response.Redirect "rettigheter.asp?ID=" & strID & "&BrukerID=" & strBrukerID
End If 
%>
