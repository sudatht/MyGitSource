<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\MailLib.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.Settings.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="includes\Xis.Security.Utils.inc"-->
<!--#INCLUDE FILE="includes\DNN.Users.inc"-->
<%

	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	Dim aAvdelingskontor	'As ()
	Dim aTjenesteomrader	'As ()
	Dim aCategory			'As ()
	Dim iSkattekort			'As Integer
	Dim sSkattekortDisabled 'As String
	dim strForerKort
	dim strDispbil
	dim strOppsigelse
	dim strWorkType
	dim strBankkontoNr
	
	'Her sjekkes og lagres data for vikar
	Function Modulus_11_Check( aNr )
	' This function checks a string after MOD 11
	'(If CheckSum = 10 erstattes 10 med et minustegn "-")

		lNrStr = Mid( aNr, 1, Len(aNr) - 1 )
		lLenNr = Len(lNrStr)

		lProdukt = 0
		lKontrol = 0

		For I = 0 To lLenNr - 1
			lProdukt = lProdukt + (Mid(lNrStr, lLenNr - I, 1) * ((I Mod 6) + 2))
		Next

		lKontrol = 11 - ( lProdukt Mod 11)

		If lKontrol = 10 Then
			Modulus_11_ControlNr = 0
		Else
			Modulus_11_ControlNr = lKontrol
		End If

		' check if correct
		If Modulus_11_ControlNr = CInt(Right(aNr,1)) Then
		Modulus_11_Check = True
		Else
		Modulus_11_Check = False
		End If

	End Function

	Sub OnlyDigits( strString )
		' Remove all non-nummeric signs from string
		For idx = 1 To Len( strString) Step 1
			Digit = Asc( Mid( strString, idx, 1 ) )
			If (( Digit > 47 ) And ( Digit < 58 )) Then
				strNewstring = strNewString & Mid(strString, idx,1)
			End If
		Next
		strString = StrNewString
	End Sub

	' Check parameters..
	If Request("tbxEtternavn") = "" Then
		 AddErrorMessage("Etternavn ikke utfyllt!")
	End If

	If Request("tbxAdresse") = "" Then
		AddErrorMessage("Adresse ikke utfyllt!")
	End If

	If Request("tbxFoedselsdato") <> "" Then
		If Not IsDate( Request("tbxFoedselsdato") ) Then
			AddErrorMessage("Fødselsdato: Ugyldig dato!")
		End If
	End If

	If Request("lsttjenesteomrader") <> "" Then
		aTjenesteomrader = split(Request("lsttjenesteomrader"),",")
	else
		AddErrorMessage("Du må velge minst ett tjenesteomr&aring;de!")
	end if

	If Request("VikarCategory") <> "" Then
		aCategory = split(Request("VikarCategory"),",")
	else
		AddErrorMessage("Vennligst velg en kategori for vikaren!")
	end if

	If Request("lstavdelingskontor") <> "" Then
		AAvdelingskontor = split(Request("lstavdelingskontor"),",")
	else
		AddErrorMessage("Du må velge minst ett avdelingskontor!")
	end if

	if (Request("cboForerkort"))="" then
		strForerKort = "null"
	else
		strForerKort = "'" & Request("cboForerkort")  & "'"
	end if

	if (Request("cboDispBil"))="" then
		strDispbil = "null"
	else
		strDispbil = "'" & Request("cboDispBil") & "'"
	end if

	if (Request("cboOppsigelse"))="" then
		strOppsigelse = "null"
	else
		strOppsigelse = "'" & Request("cboOppsigelse") & "'"
	end if

	' Work Type
	strWorkType = "'" & Request("cboWorkType") & "'"

	' Get Fødselnr
	strDate = Request("tbxFoedselsdato")

	If Request("tbxPersonnummer") <> "" and Request("tbxPersonnummer") <> "0" Then
		' Remove separator
		Call OnlyDigits( strDate )

		' Create Fødselsnummer
		strNumber =  strDate & Request("tbxPersonnummer")

		' Check Fødselnummer
		If Not Modulus_11_Check( strNumber ) Then
			AddErrorMessage("Ugyldig fodselsnummer:" & strNumber)
		End If
	End If

	' Henter verdi for skattekort(Default: -1 -> ikke levert)
	If Request("skattekort") = "" Then
		iSkattekort = -1
	Else
		iSkattekort = Request("skattekort")
	End If

	sSkattekortDisabled = Request("tbxSkattekortDisabled")

	' Timeloen must have value
	If Request("tbxTimeloenn") = "" Then
		lTimeloenn = 0
	Else
		lTimeloenn = Request("tbxTimeloenn")
	End IF

	' Personnummer must have value
	If Request("tbxPersonnummer") = "" Then
		lPersonnummer = 0
	Else
		lPersonnummer = Request("tbxPersonnummer")
	End IF

	'Kjønn må være fylt ut
	If Request("rbnKjoenn") = "" Then
		AddErrorMessage("Du må angi kjønn på vikaren!")
	End If

	if len(request("tbxBankkontonr"))>0 then
		strBankkontoNr =  trim(request("tbxBankkontonr"))
		sReg	= "^\d{4}\.\d{2}\.\d{5}$"
		set oRegExp = New regexp
		oRegExp.Pattern = sReg
		if (oRegExp.Test(strBankkontoNr)=false) then
			AddErrorMessage("Ugyldig bankkontonummer!")
		end if
		set oRegExp = nothing
	end if

	If Request("dbxStatus") = "" Then
		AddErrorMessage("Du må angi status på vikaren!")
	End If

	if len(trim(Request("tbxPostNr"))) > 0 then
		dim regEx
		Set regEx = New RegExp
		regEx.Pattern = "^\d*$"
		if (regEx.test(Request("tbxPostNr")) = false) then
			AddErrorMessage("Postnr må være nummerisk!")
		end if
	end if

	' Get a database connection
	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))	

	if(HasError() = true) then
		call RenderErrorMessage()
	end if

' Action against database depending on Button pressed
'Ny vikar
If (Request("tbxVikarID") = "0") Then
	'TODO need to double check what exactly this code does
	'if not VerifyUserName(trim(Request("Username")),0,"ANSATT") then
	'	AddErrorMessage("Dette brukernavnet er allerede i bruk!")
	'	call RenderErrorMessage()
	'end if

   '  Status ansatt => setter godkjent av og godkjentDato
   If CInt( Request("dbxStatus") ) = 3 Then
      GodkjentAv = Session("Brukernavn")
      GodkjentDato = Date()
   End If

   'Setter registrert dato
   Regdato = Date()

   ' Kurskode
   If Request("rbnKurskode") = "" Then
      Kurskode = 0
   Else
      Kurskode = Request("rbnKurskode")
   End If

   strEtternavn = Request("tbxEtternavn")
   strFornavn = Request("tbxFornavn")
   strLink1URL = Request("Link1URL")
   strLink2URL = Request("Link2URL")
   strLink3URL = Request("Link3URL")

  ' create new Vikar in database
  strSQL = "Insert into Vikar( Fornavn, Etternavn, Foedselsdato," & _
		"StatusID, TypeID, AnsMedID, Notat,InterestedJobs, Oppsummering_intervju, oppsummering_ref_sjekk, loenn1," & _
		"Intervjudato, Telefon, MobilTlf, fax, Epost, "&_
		"kontraktsendt, kontraktmottatt, endring, kurskode, kjoenn," & _
		"GodkjentAv, GodkjentDato, Regdato, KundePresentasjon, " & _
		"Foererkort, hasCar, Oppsigelsestid,WorkType, Link1URL,Link2URL,Link3URL,Country,MottattSkattekort, Bankkontonr )" &_
		"Values('" & Request("tbxFornavn") & "','" &_
		Request("tbxEtternavn") & "'," & _
		DbDate( Request("tbxFoedselsdato") ) & "," & _
		Request("dbxStatus") & "," & _
		Request("dbxType") & "," &_
		Request("dbxMedarbeider") & ",'" & _
		PadQuotes(Request("tbxNotat")) & "', '" &_
		PadQuotes(Request("tbxJobs")) & "', '" &_
		PadQuotes(Request("tbxOppsumInt")) & "','" &_
		PadQuotes(Request("tbxOppsumRef")) & "'," &_
		lTimeloenn & "," & _
		DbDate( Request("tbxIntervjudato") ) & "," & _
		"'" & Request("tbxTelefon") & "'," & _
		"'" & Request("tbxMobilTlf") & "'," & _
		"'" & Request("tbxFax") & "'," & _
		"'" & Request("tbxEPost") & "'," & _		
		DbDate( Request("tbxKontraktSendt") ) & "," &_
		DbDate( Request("tbxKontraktMottatt") ) & "," &_
		Request("endring") & "," &_
		Kurskode & "," &_
		Quote(Request("rbnKjoenn")) & "," &_
		Quote( GodkjentAv ) & " ," &_
		DbDate( GodkjentDato ) & " ," &_
		DbDate(Regdato) & " ," & _
		Quote(Request("tbxPresentasjon")) & "," & _
		strForerKort & "," & _
		strDispbil & "," & _
		strOppsigelse & "," & _
		strWorkType & "," & _
		Quote(strLink1URL) & "," & _
		Quote(strLink2URL) & "," & _
		Quote(strLink3URL) & "," & _
		Request("dbxCountry") & "," &_
		iSkattekort & "," & _
		Quote(strBankkontoNr) & ")"

	' Start transaction
	Conn.Begintrans

   If ExecuteCRUDSQL(strSQL, Conn) = false then
      ' Rollback transaction
		Conn.RollBackTrans
		AddErrorMessage("En eller flere feil oppstod under oppretting av vikar.")
		call RenderErrorMessage()
   End if

   ' Get new VikarID
   Set rsVikar  = Conn.Execute("Select NewVikarID=max( VikarID ) from Vikar")

   NewVikarID = rsVikar("NewVikarID")

   ' Close and release recordset
   rsVikar.Close
   Set rsVikar = Nothing

   ' Create new main adress
   ' AdresseRelasjon: 1 = Kunde / 2 => Vikar
   strSQL = "INSERT INTO ADRESSE(AdresseRelasjon, AdresseRelID, Adresse, Postnr, Poststed, AdresseType ) " & _
           "VALUES( 2," & NewVikarID & "," & _
		   "'" & Request("tbxadresse") & "'," & _
		   Quote( Request("tbxPostNr") )& "," & _
		   Quote( Request("tbxPoststed") )& "," & _
		   Request("tbxAdrType") & ")"

   If ExecuteCRUDSQL(strSQL, Conn) = false then
      ' Rollback transaction
		Conn.RollBackTrans
		AddErrorMessage("En eller flere feil oppstod under oppretting av vikar.")
		call RenderErrorMessage()
   End if

   '  Status Intern ansatt ?
   If CInt( Request("dbxStatus") ) = 4 Then

		strSQL  = "INSERT INTO MEDARBEIDER( Fornavn, Etternavn, VikarID ) " &_
                "Values( " & Quote( Request("tbxFornavn") ) & "," &_
				Quote( Request("tbxEtternavn") ) & "," & _
				NewVikarID & ")"

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			' Rollback transaction
			Conn.RollBackTrans
			AddErrorMessage("En eller flere feil oppstod under oppretting av vikar.")
			call RenderErrorMessage()
		End if
   End If

   ' Insert into VIKAR_AVDELING
   For I = 0 To ubound(aAvdelingskontor)
		strSQL = "INSERT INTO VIKAR_ARBEIDSSTED( VikarID, AvdelingskontorID ) Values(" & NewVikarID & "," & aAvdelingskontor(I) & ")"

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			' Rollback transaction
			Conn.RollBackTrans
			AddErrorMessage("En eller flere feil oppstod under oppretting av vikar.")
			call RenderErrorMessage()
		End if
   Next

   ' Insert into VIKAR_TJENESTEOMRADE
	For I = 0 To ubound(aTjenesteomrader)
		strSQL = "INSERT INTO VIKAR_TJENESTEOMRADE( VikarID, tomID ) Values(" & NewVikarID & "," & aTjenesteomrader(I) & ")"

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			' Rollback transaction
			Conn.RollBackTrans
			AddErrorMessage("En eller flere feil oppstod under oppretting av vikar.")
			call RenderErrorMessage()
		End if
	Next
	
	' Insert into VIKAR_CATEGORY
	For I = 0 To ubound(aCategory)
		strSQL = "INSERT INTO VIKAR_CATEGORY( VikarID, CategoryID ) Values(" & NewVikarID & "," & aCategory(I) & ")"

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			' Rollback transaction
			Conn.RollBackTrans
			AddErrorMessage("En eller flere feil oppstod under oppretting av vikar.")
			call RenderErrorMessage()
		End if
	Next
	
	
	'Log activity for create new vikar in xis
	Dim strActivity
	Dim rsActivityType
	Dim strComment
	Dim rsVikarStatus
	strActivity = "Søker reg. som vikar"
	'Get vikar status
	set rsVikarStatus = GetFirehoseRS("SELECT Vikarstatus FROM H_VIKAR_STATUS WHERE VikarstatusID = " & Request("dbxStatus"), Conn)
	strComment = "Registrert som vikar med status " & rsVikarStatus("Vikarstatus")
	' Close and release recordset
      	rsVikarStatus.Close
      	Set rsVikarStatus = Nothing
	 
	'Get activity id
	set rsActivityType = GetFirehoseRS("SELECT AktivitetTypeID FROM H_AKTIVITET_TYPE WHERE AktivitetType = '" & strActivity & "'", Conn)
	nActivityTypeID = CInt(rsActivityType("AktivitetTypeID"))
	' Close and release recordset
      	rsActivityType.Close
      	Set rsActivityType = Nothing
	sDate = GetDateNowString()
	strSql = "INSERT INTO AKTIVITET( AktivitetTypeID, AktivitetDato, VikarID, Notat, Registrertav, RegistrertAvID, AutoRegistered ) Values(" & nActivityTypeID & ",'" & sDate & "'," & NewVikarID & ",'" & strComment & "','" & Session("Brukernavn") & "'," & Request("dbxMedarbeider") & ", 1)"
	
	If ExecuteCRUDSQL(strSql, Conn) = false then
		' Rollback transaction
		Conn.RollBackTrans
		AddErrorMessage("Aktivitetsregistrering for ny vikar i xis feilet.")
		call RenderErrorMessage()
	End if
	

  ' End transaction
  Conn.CommitTrans

   ' set return value
   strVikarID = NewVikarID
   call WebBruker()

ElseIf (Request("tbxVikarID") <> "0" )Then
	'TODO need to figure out what this code does
	'if not VerifyUserName(trim(Request("Username")), clng(Request.form("tbxVikarID")), "ANSATT") then
	'	AddErrorMessage("Dette brukernavner er allerede i bruk!")
	'	call RenderErrorMessage()
	'end if

	'Auto register Activity for Vikar
	Dim nOldVikarStatus
	Dim nOldResponsiblePer
	Dim rsPrevStatus
	
	set rsPrevStatus = GetFirehoseRS("SELECT StatusID, AnsMedID FROM VIKAR WHERE VikarID = " & Request("tbxVikarID"), Conn)
	nOldVikarStatus = CInt( rsPrevStatus("StatusID"))
	nOldResponsiblePer = CInt( rsPrevStatus("AnsMedID"))
	
	' Close and release recordset
      	rsPrevStatus.Close
      	Set rsPrevStatus = Nothing
	'End register activity
	
   '  Status vikar ?
   If CInt( Request("dbxStatus") ) = 3 Then

		' Get previous status
		set rsVikarStatus = GetFirehoseRS("SELECT statusid, godkjentav, godkjentdato FROM VIKAR WHERE VikarID = " & Request("tbxVikarID"), Conn)

      ' Previous status PROSPECT ?
      If CInt( rsVikarStatus("Statusid") ) = 1 or CInt( rsVikarStatus("Statusid") ) = 2 Then
         ' Set values for accepted customer
         GodkjentAv = Session("Brukernavn")
         GodkjentDato = Date()
      Else
         ' Set values like previous values
         GodkjentAv =  rsVikarStatus("Godkjentav")
         GodkjentDato = rsVikarStatus("GodkjentDato")
      End If

      ' Close and release recordset
      rsVikarStatus.Close
      Set rsVikarStatus = Nothing
   End If

   ' Kurskode
	If Request("rbnKurskode") = "" Then
		Kurskode = 0
	Else
		Kurskode = Request("rbnKurskode")
	End If

	' Update Vikar in database
	strSQL = "UPDATE Vikar SET " &_
		" Fornavn      = " & "'" & Request("tbxFornavn") & "'" & _
		", Etternavn    = " & "'" & Request("tbxEtternavn") & "'" & _
		", Foedselsdato = " & DbDate( Request("tbxFoedselsdato") ) & _
		", Notat        = " & "'" & PadQuotes(Request("tbxNotat")) & "'" & _
		", InterestedJobs  = " & "'" & PadQuotes(Request("tbxJobs")) & "'" & _   
		", Oppsummering_intervju = ' " &PadQuotes(Request("tbxOppsumInt")) & "'" &_
		", Oppsummering_ref_sjekk = '" &PadQuotes(Request("tbxOppsumRef")) & "'" &_
		", StatusID     = " & Request("dbxStatus") & _
		", TypeID       = " & Request("dbxType") & _
		", AnsMedID     = " & Request("dbxMedarbeider") & _
		", loenn1    = " & lTimeloenn & _
		", IntervjuDato = " & DbDate( Request("tbxIntervjudato") ) & _
		", Telefon      = " & "'" & Request("tbxTelefon") & "'" & _
		", MobilTlf     = " & "'" & Request("tbxMobilTlf") & "'" & _
		", Fax          = " & "'" & Request("tbxFax") & "'" & _
		", EPost        = " & "'" & Request("tbxEPost") & "'" & _		
		", GodkjentAv    = " & Quote( GodkjentAv ) &_
		", GodkjentDato = " & DbDate( GodkjentDato ) &_
		", Kontraktsendt = " & DbDate( Request("tbxKontraktsendt") ) & _
		", Kontraktmottatt = " & DbDate( Request("tbxKontraktmottatt") ) & _
		", Endring = 1"  & _
		", Kurskode = " & Kurskode & _
		", KundePresentasjon = '" & PadQuotes(Request("tbxPresentasjon"))& "'" &_
		", Foererkort = " & strForerKort & _
		", hasCar = " & strDispbil & _
		", Oppsigelsestid = " & strOppsigelse & _
		", WorkType = " & strWorkType & _
		", Link1URL    = " & "'" & Request("Link1URL") & "'" & _		
		", Link2URL    = " & "'" & Request("Link2URL") & "'" & _		
		", Link3URL    = " & "'" & Request("Link3URL") & "'" & _		
		", Country       = " & Request("dbxCountry") & _
		", Kjoenn = " & Quote(Request("rbnKjoenn")) & _
		", Bankkontonr= " & Quote(strBankkontoNr)

	If (LCase(sSkattekortDisabled) = "false") Then
		strSQL = strSQL & ", MottattSkattekort = " & iSkattekort
	End If

	strSQL = strSQL & " WHERE Vikarid = " & Request("tbxVikarID")

	NewVikarID = Request("tbxVikarID")

	' Start transaction
	Conn.Begintrans

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		' Rollback transaction
		Conn.RollBackTrans
		AddErrorMessage("En eller flere feil oppstod under oppdatering av vikar.")
		call RenderErrorMessage()
	End if


	' Update Adresse in database
	strSQL = "Update Adresse set " &_
		" Adresse      = " & "'" & Request("tbxAdresse") & "'," & _
		" Postnr = " & Quote( Request("tbxPostNr") )& "," & _
		" Poststed = " & Quote( Request("tbxPoststed") ) & _
		" where Adrid =" & Request("tbxAdrID")

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		' Rollback transaction
		Conn.RollBackTrans
		AddErrorMessage("En eller flere feil oppstod under oppdatering av vikar.")
		call RenderErrorMessage()
	End if

   '  Status Intern ansatt ?
   If CInt( Request("dbxStatus") ) = 4 Then

		' VikarId exist in medarbeider ?
		Set rsMedarbeider  = Conn.Execute("Select antall = count( VikarID ) from MEDARBEIDER where VikarID = " & Request("tbxVikarID") )

		Antall = rsMedarbeider( "Antall" )

		' Close and release recordset
		rsMedarbeider.Close
		Set rsMedarbeider = Nothing

		If CInt( Antall ) = 0 Then
			strSql  = "INSERT INTO MEDARBEIDER( Fornavn, Etternavn, VikarID ) " &_
					"Values( " & Quote( Request("tbxFornavn") ) & "," &_
					Quote( Request("tbxEtternavn") ) & "," & _
					Request("tbxVikarID") & ")"
		Else
			strSql  = " UPDATE MEDARBEIDER  Set " &_
					" Fornavn = " & Quote( Request("tbxFornavn") ) & "," &_
					" Etternavn = " & Quote( Request("tbxEtternavn") ) &_
					" where Vikarid =" & Request("tbxVikarID")
		End If

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			' Rollback transaction
			Conn.RollBackTrans
			AddErrorMessage("En eller flere feil oppstod under oppdatering av vikar.")
			call RenderErrorMessage()
		End if
	End If


	' Delete all connected ARBEIDSSTEDER
	strSql = "DELETE FROM VIKAR_ARBEIDSSTED where VikarID = " & Request("tbxVikarID")

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		' Rollback transaction
		Conn.RollBackTrans
		AddErrorMessage("En eller flere feil oppstod under oppdatering av vikar.")
		call RenderErrorMessage()
	End if

	' Update VIKAR_ARBEIDSSTED
	For I = 0 To ubound(aAvdelingskontor)
		strSQL = "INSERT INTO VIKAR_ARBEIDSSTED( VikarID, AvdelingskontorID ) Values(" & NewVikarID & "," & aAvdelingskontor(I) & ")"

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			' Rollback transaction
			Conn.RollBackTrans
			AddErrorMessage("En eller flere feil oppstod under oppdatering av vikar.")
			call RenderErrorMessage()
		End if
	Next

   ' Delete all connected TJENESTEOMRADER
	strSQL = "delete from VIKAR_TJENESTEOMRADE where VikarID = " & Request("tbxVikarID")

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		' Rollback transaction
		Conn.RollBackTrans
		AddErrorMessage("En eller flere feil oppstod under oppdatering av vikar.")
		call RenderErrorMessage()
	End if

	' Insert into VIKAR_TJENESTEOMRADE
	For I = 0 To ubound(aTjenesteomrader)
		strSql = "INSERT INTO VIKAR_TJENESTEOMRADE( VikarID, tomID ) Values(" & NewVikarID & "," & aTjenesteomrader(I) & ")"

		If ExecuteCRUDSQL(strSql, Conn) = false then
			' Rollback transaction
			Conn.RollBackTrans
			AddErrorMessage("En eller flere feil oppstod under oppdatering av vikar.")
			call RenderErrorMessage()
		End if
	Next

	' Delete all connected CATEGORIES
	strSQL = "delete from VIKAR_CATEGORY where VikarID = " & Request("tbxVikarID")

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		' Rollback transaction
		Conn.RollBackTrans
		AddErrorMessage("En eller flere feil oppstod under oppdatering av vikar.")
		call RenderErrorMessage()
	End if
	
	' Insert into VIKAR_CATEGORY
	For I = 0 To ubound(aCategory)
		strSQL = "INSERT INTO VIKAR_CATEGORY( VikarID, CategoryID ) Values(" & NewVikarID & "," & aCategory(I) & ")"

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			' Rollback transaction
			Conn.RollBackTrans
			AddErrorMessage("En eller flere feil oppstod under oppretting av vikar.")
			call RenderErrorMessage()
		End if
	Next

	'Write Activity to DB if status or responsible is changed in Vikar
	
	Dim nActivityTypeID
	Dim nResponsiblePerID
	Dim strNewStatus
		
	Dim rsTempRecordset
	nResponsiblePerID = CInt(Session("medarbID"))
	Dim sDate
	if CInt(Request("dbxStatus"))<>nOldVikarStatus then
	
		strActivity = "V status endret"
		
		'Get new vikar status
	      	set rsActivityType = GetFirehoseRS("SELECT Vikarstatus FROM H_VIKAR_STATUS WHERE VikarstatusID = " & CInt(Request("dbxStatus")), Conn)
		strNewStatus = rsActivityType("Vikarstatus")
		' Close and release recordset
	      	rsActivityType.Close
	      	Set rsActivityType = Nothing
	      	
	      	
		strComment = "Vikar status endret til " & strNewStatus
		
	      	'Temp status is changed first time to "Employed".
	      	if strNewStatus="Ansatt" then
	      		strActivity = "V ansatt"
	      		strComment = "Vikar godkjent og ansatt. Ansattnummer 0"
	      		
	      		'Retrieve the employee number from VIKAR_ANSATTNUMMER
	      		set rsTempRecordset = GetFirehoseRS("SELECT ansattnummer FROM VIKAR_ANSATTNUMMER WHERE vikarId = " & NewVikarID, Conn)
			
			if not rsTempRecordset.EOF then
				strComment = "Vikar godkjent og ansatt. Ansattnummer " & CInt(rsTempRecordset("ansattnummer"))
				' Close and release recordset
			      	rsTempRecordset.Close
			      	Set rsTempRecordset = Nothing
			end if
				      		
	      	end if
	      	
		set rsActivityType = GetFirehoseRS("SELECT AktivitetTypeID FROM H_AKTIVITET_TYPE WHERE AktivitetType = '" & strActivity & "'", Conn)
		nActivityTypeID = CInt(rsActivityType("AktivitetTypeID"))
		' Close and release recordset
	      	rsActivityType.Close
	      	Set rsActivityType = Nothing
	      	
	      	sDate = GetDateNowString()
		strSql = "INSERT INTO AKTIVITET( AktivitetTypeID, AktivitetDato, VikarID, Notat, Registrertav, RegistrertAvID, AutoRegistered ) Values(" & nActivityTypeID & ",'" & sDate & "'," & NewVikarID & ",'" & strComment & "','" & Session("Brukernavn") & "'," & nResponsiblePerID & ", 1)"
		
		If ExecuteCRUDSQL(strSql, Conn) = false then
			' Rollback transaction
			Conn.RollBackTrans
			AddErrorMessage("Aktivitetsregistrering på vikar feilet.")
			call RenderErrorMessage()
		End if
		
		
	end if
	
	Dim rsResponsiblePerson
	Dim strRPName
	
	if nOldResponsiblePer<>CInt(Request("dbxMedarbeider")) then
		strActivity = "Endret ansv."
		set rsActivityType = GetFirehoseRS("SELECT AktivitetTypeID FROM H_AKTIVITET_TYPE WHERE AktivitetType = '" & strActivity & "'", Conn)
		nActivityTypeID = CInt(rsActivityType("AktivitetTypeID"))
		' Close and release recordset
	      	rsActivityType.Close
	      	Set rsActivityType = Nothing
	      	
	      	set rsResponsiblePerson = GetFirehoseRS("SELECT Etternavn,Fornavn FROM MEDARBEIDER WHERE MedID = " & CInt(Request("dbxMedarbeider")) , Conn)
		strRPName = rsResponsiblePerson("Fornavn") & " " & rsResponsiblePerson("Etternavn")
		' Close and release recordset
	      	rsResponsiblePerson.Close
	      	Set rsResponsiblePerson = Nothing
	      	
	      	strComment = "Ansvarlig endret til " & strRPName
	      	sDate = Year(Now()) & "-" & Month(Now()) & "-" & Day(Now()) & " " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now())		
		strSql = "INSERT INTO AKTIVITET( AktivitetTypeID, AktivitetDato, VikarID, Notat, Registrertav, RegistrertAvID, AutoRegistered ) Values(" & nActivityTypeID & ",'" & sDate & "'," & NewVikarID & ",'" & strComment & "','" & Session("Brukernavn") & "'," & nResponsiblePerID & ", 1)"
		'response.write(strSql)
		If ExecuteCRUDSQL(strSql, Conn) = false then
			' Rollback transaction
			
			Conn.RollBackTrans
			AddErrorMessage("Aktivitetsregistrering ved bytte av ansvarlig feilet.")
			call RenderErrorMessage()
		End if
	end if
	'End write activity
	
  ' Commit transaction
  Conn.CommitTrans

   ' Set return value
   strVikarID = Request("tbxVikarID")
   call WebBruker()

ElseIf Trim( Request("pbnDataAction") ) = "Slette" Then
   ' Delete Vikar in database
   strSQL = "DELETE Vikar WHERE Vikarid = " & Request("tbxVikarID")

	If ExecuteCRUDSQL(strSql, Conn) = false then
		' Rollback transaction
		Conn.RollBackTrans
		AddErrorMessage("En eller flere feil oppstod under sletting av vikar.")
		call RenderErrorMessage()
	End if

   ' Set return value
   strVikarID = ""
Else
	AddErrorMessage("Feil i parameter.")
	call RenderErrorMessage()
End If

CloseConnection(Conn)
set Conn = nothing

Response.redirect "vikarvis.asp?VikarID=" & strVikarID

sub WebBruker()
	'IMP Brukerhåndtering 
	'cmu@EC changed the function to create users in DNN

	dim cnXis		' Xis connection
	dim rsUser 			'Recordset containing user data
	dim strUsername 	'Consultant's userid
	dim strSubstitute 	'string
	dim IVikarID 		'Consultant's id
	dim strPassword 	'contains new Password
	dim status
	dim cons
	dim cv
	dim CVisLocked
	Dim objUserProxy  'Web service proxy for the DNN user service
	Dim sUserServiceURL 'Url of the user web service
	Dim objUserDom
	Dim iApp
	Dim sFirstName
	Dim sLastName
	Dim sEmail
	Dim sUserXml

	IVikarID = Request.form("tbxVikarID")
	sFirstName = Trim(Request("tbxFornavn"))
	sLastName = Trim(Request("tbxEtternavn"))
	sEmail = Trim(Request("tbxEPost"))

	if clng(IVikarID) = 0 then
		IVikarID = NewVikarID
	end if

	iApp = Cint(Application("Application"))
	sUserServiceURL = Application("DNNUserServiceURL")
	
	'Initialize ADO objects
	Set cnXis = GetClientConnection(GetConnectionstring(XIS, ""))	
	Set objUserProxy = Server.CreateObject("ECDNNUserProxy.UserServiceProxy")
	objUserProxy.Url = sUserServiceURL

	
	status = Request.form("dbxStatus")
	strUsername = trim(Request("Username"))

	'Cv edit - Handles locking and unlocking of CV
	set cons = Server.CreateObject("XtraWeb.Consultant")

	cons.XtraConString = Application("XtraWebConnection")
	cons.GetConsultant(IVikarID)

	set cv	= cons.CV
	cv.XtraConString = Application("Xtra_intern_ConnectionString")
	cv.XtraDataShapeConString = Application("ConXtraShape")
	cv.Refresh

	if cons.CV.DataValues.Count = 0 then
		cons.CV.Save
	end if

	if ( (len(strUsername)>0 ) and (CInt(status) = 3 or CInt(status) = 2 or CInt(status) = 1 OR CInt(status) = 8)  ) then
	'Creates or updates Web user

		'Get existing user details from xis
		sUserXml = objUserProxy.GetUser(iApp, IVikarID,"V")

		if sUserXml <> "" then
			Set objUserDom = Server.CreateObject("Microsoft.XMLDOM")
			objUserDom.LoadXml sUserXml
			
			If Trim(strUsername) <> objUserDom.selectSingleNode("/user/userName").Text Then
				if objUserProxy.IsDNNUserExist(strUsername) then	' if dnn user exists, cannot create the user again
					AddErrorMessage("User already exists with the same username") // 1
					Call RenderErrorMessage()
				end if
				strPassword = GeneratePassword(6)
				Set objUserDom = CreateUserDom(iApp, IVikarID, strUsername, strPassword, sFirstName, sLastName, sEmail,"V")
				Set objUserDom =AppendRoles(objUserDom, aTjenesteomrader)
				If Not objUserProxy.SaveUser(objUserDom.xml) Then
                                        
					AddErrorMessage("Failed to save web user") // 1
					Call RenderErrorMessage()
				End If
				Call SendmailPWDIUD(IVikarID, strPassword, "ANSATT")
			Else ' User name is the same just update the user info and roles
				Set objUserDom = CreateUserDom(iApp, IVikarID, "", "", sFirstName, sLastName, sEmail,"V")
				Set objUserDom = AppendRoles(objUserDom, aTjenesteomrader)
                               
				If Not objUserProxy.SaveUser(objUserDom.xml) Then

					AddErrorMessage("Failed to save web user") //2
					Call RenderErrorMessage()
				End If
			End If

			if (Cv.isLocked = true and cint(Request("chkEditCV")) = 1 and cint(status) = 1) then 'søker med redigere cv-rettighet skal få egen epost
				call SendmailUPDATECV(IVikarID, strPassword)
			end if

		Else
		
		   if( cint(Request("chkCreateuser")) = 1 )  Then
			
				if objUserProxy.IsDNNUserExist(strUsername) then	' if dnn user exists, cannot create the user again
					AddErrorMessage("User already exists with the same username") // 1
					Call RenderErrorMessage()
				end if
				
			'Insert substitute as new user
			strPassword = GeneratePassword(6)
			Set objUserDom = CreateUserDom(iApp, IVikarID, strUsername, strPassword, sFirstName, sLastName, sEmail,"V")			
			Set objUserDom = AppendRoles(objUserDom, aTjenesteomrader)

			If Not objUserProxy.SaveUser(objUserDom.xml) Then
                         
				AddErrorMessage("Failed to save web user") //3
				Call RenderErrorMessage()
			End If			


			if (cint(Request("chkEditCV")) = 1 and clng(status) = 1) then 'søker med redigere cv-rettighet skal få egen epost
				call SendmailUPDATECV(IVikarID, strpassword)
			else
				call SendmailPWDIUD(IVikarID, strpassword, "ANSATT")
			end if
		    end if
		end if

	elseif ( ( cint(Request("chkCreateuser"))= 0 ) and (len(strUsername)= 0 ) ) then
	'Remove consultant from DNN and Mapping table altogether
    'With new recruiment process existing users does not remove from the system
		'sUserXml = objUserProxy.GetUser(iApp, IVikarID,"V")
	'	If sUserXml <> "" Then		
			'If objUserProxy.DeleteUser(iApp, IVikarID,"V") = False Then
			'	'AddErrorMessage("Feil under oppdatering av webbruker")
				'Call RenderErrorMessage()
			'End if
		'End If
	end if


	if (Cv.isLocked = true and cint(Request("chkEditCV")) = 1) then
		cv.unlockCV
	elseif (Cv.isLocked = false and cint(Request("chkEditCV")) = 0) then
		cv.lockCV
	end if

	set cv = nothing
	cons.CV.cleanup
	cons.cleanup
	set cons = nothing

	'Clean up all objects after use
	CloseConnection(cnXis)
	set cnXis = nothing
	'/ Brukerhåndtering
end sub
 %>
