<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	Dim RsProductKnowledge 	'as adodb.recordset
	'Dim ObjKompCon 		'as adodb.connection
	'Dim aKompetanse 		'as string()
	'Dim aFagomrade 		'as string()
	'Dim iKompChoosen 		'as integer
	Dim iFagomChoosen 		'as integer
	Dim StrSQL				'as string
	Dim rsOverfort			'as adodb.recordset
	dim Conn

	'Aktuell suspect
	suspectID = request.querystring("suspectID")
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

	strSQL = "SELECT overfort = count(suspectid) FROM v_suspect WHERE suspectid = " & suspectID & " and overfort = 1"
	set rsOverfort = GetFirehoseRS(strSQL, Conn)	
	if rsOverfort.fields("Overfort") = 1 then
		rsOverfort.close
		set rsOverfort = nothing
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("Søkeren er allerede overført!")
		call RenderErrorMessage()
	end if
	rsOverfort.close
	set rsOverfort = nothing 
	
	'Sett ansvarlig konsulent for å behandle søknaden
	if Request.querystring("oppdat") = "ja" Then
		ansMed = request("dbxMedarbeider")
		If ExecuteCRUDSQL("UPDATE V_SUSPECT set AnsMedID=" & ansMed & " WHERE suspectID = " & suspectID, Conn) = false then
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Feil oppstod under overføring til ansvarlig.")
			call RenderErrorMessage()
		End if		
		response.redirect "SuspectList.asp"
	end if

	' Check datavalues
	If NOT len(Request.Form("tbxEtternavn")) > 0Then
		AddErrorMessage("Etternavn ikke utfyllt!")		
	End If

	If Trim(Request.Form("tbxAdresse")) = "" Then
		AddErrorMessage("Adresse ikke utfyllt!")		
	End If

	If  Trim(Request.Form("tbxFoedselsdato")) <> "" Then
		If Not IsDate( Request.Form("tbxFoedselsdato") ) Then
   			AddErrorMessage("Fødselsdato: ugyldig dato!")
		End If
	End If

	' Get Fødselnr
	strDate = Request.Form("tbxFoedselsdato")

	If Request.Form("tbxPersonnummer") <> "" and Request.Form("tbxPersonnummer") <> "0" Then
		' Remove seperator
		Call OnlyDigits( strDate )

		' Create Fødselsnummer
		strNumber =  strDate & Request.Form("tbxPersonnummer")

		' Check Fødselnummer
		If Not Modulus_11_Check( strNumber ) Then
			AddErrorMessage("FødselsNummer: ugyldig fødselsdato! " & strNumber)
		End If
	End If

	'Kjønn må være fylt ut
	If Request("rbnKjoenn") = "" Then
		AddErrorMessage("Kjønn må fylles ut før lagring!")	
	End If

	If Request.Form("tbxPersonnummer") = "" Then
		lPersonnummer = 0
	Else
		lPersonnummer = Request.Form("tbxPersonnummer")
	End if

	If Request("lstavdelingskontor") <> "" Then
		AAvdelingskontor = split(Request("lstavdelingskontor"),",")
	else
		AddErrorMessage("Du m&aring; velge minst ett avdelingskontor!")	
	end if

	if HasError() then
		call RenderErrorMessage()
	end if

	' Action against database depending on Button pressed
	If Request("overfoer") = "ja" Then

	'  Status kunde ?
	If CInt( Request.form("dbxStatus") ) = 3 Then
		GodkjentAv = Session("Brukernavn")
		GodkjentDato = Date()
	End If
	
	Regdato = Date()
	
	' Kurskode
	If Request.form("rbnKurskode") = "" Then
		Kurskode = 0
	Else
		Kurskode = Request.form("rbnKurskode")
	End If

	strSQL = "INSERT INTO [Vikar]( Fornavn, Etternavn, Foedselsdato," & _
			"StatusID, TypeID, AnsMedID, Notat," & _
			"Telefon,MobilTlf,fax,Epost,Hjemmeside, "&_
			"endring, kurskode, kjoenn, " & _
			"GodkjentAv, GodkjentDato, Regdato )" &_
			" VALUES('" & Request.form("tbxFornavn") & "','" &_
			Request.form("tbxEtternavn") & "'," & _
			DbDate( Request.form("tbxFoedselsdato") ) & "," & _
			Request.form("dbxStatus") & "," & _
			Request.form("dbxType") & "," &_
			Request.form("dbxMedarbeider") & ",'" & _
			Request.form("tbxNotat") & "'," &_
			"'" & Request.form("tbxTelefon") & "'," & _
			"'" & Request.form("tbxMobilTlf") & "'," & _
			"'" & Request.form("tbxFax") & "'," & _
			"'" & Request.form("tbxEPost") & "'," & _
			"'" & Request.form("tbxHjemmeside") & "','1'," & _
			Kurskode & "," &_
			Quote(Request("rbnKjoenn")) & "," &_
			Quote( GodkjentAv ) & " ," &_
			DbDate( GodkjentDato ) & " ," &_
			DbDate(Regdato) & ")"

	' Start transaction
	Conn.Begintrans

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		AddErrorMessage("Feil under overføring av suspect.")
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing		
		call RenderErrorMessage()
	End if	
	
	' Get new VikarID
	set rsVikar = GetFirehoseRS("SELECT NewVikarID = MAX( VikarID ) from Vikar", Conn)	
	NewVikarID = rsVikar("NewVikarID")

	' Close and release recordset
	rsVikar.Close
	Set rsVikar = Nothing

	' Create new main adress
	' AdresseRelasjon: 1 = Kunde / 2 => Vikar
	strSQL = "Insert into ADRESSE(AdresseRelasjon, AdresseRelID, Adresse, Postnr, Poststed, AdresseType ) " & _
			"Values( 2," & NewVikarID & "," & _
			"'" & Request.form("tbxadresse") & "'," & _
			Quote( Request.form("tbxPostNr") )& "," & _
			Quote( Request.form("tbxPoststed") )& "," & _
			Request.form("tbxAdrType") & ")"

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		AddErrorMessage("Feil under overføring av suspects adresse.")
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing		
		call RenderErrorMessage()
	End if	

	'  Status Intern ansatt ?
	If CInt( Request.form("dbxStatus") ) = 4 Then

		strSQL  = "INSERT INTO MEDARBEIDER( Fornavn, Etternavn, VikarID ) " &_
				"Values( " & Quote( Request.form("tbxFornavn") ) & "," &_
				Quote( Request.form("tbxEtternavn") ) & "," & _
				NewVikarID & ")"
		
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			AddErrorMessage("Feil under overføring av medarbeider info.")
			Conn.RollBackTrans
			CloseConnection(Conn)
			set Conn = nothing			
			call RenderErrorMessage()
		End if	
	End If

	'   ' Insert into VIKAR_AVDELING
	For I = 0 To ubound(aAvdelingskontor)
		strSQL = "INSERT INTO [VIKAR_ARBEIDSSTED]( VikarID, AvdelingskontorID ) Values(" & NewVikarID & "," & aAvdelingskontor(I) & ")"
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			AddErrorMessage("Feil under overføring av avdelingskontor.")
			Conn.RollBackTrans
			CloseConnection(Conn)
			set Conn = nothing			
			call RenderErrorMessage()
		End if
	Next

	'End transaction
	Conn.CommitTrans
	
	    

	
	%>
	<!--#INCLUDE FILE="Suspect_overfor_cv.asp"-->
	<%
	End If
		'DNN user updateion goes here
	Call UserMapUpdate(suspectID,NewVikarID )
	call ExecuteCRUDSQL("UPDATE [v_suspect] SET [overfort] = 1 WHERE [suspectID] = " & suspectID, Conn)

	CloseConnection(Conn)
	set Conn = nothing

	
	
	Response.redirect "SuspectVis.asp?suspectID=" & suspectID & "&VikarID=" & NewVikarID


'Functions used 

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
				strNewstring = strNewString & Mid(strString,idx,1)
			End If
		Next
		strString = StrNewString
	End Sub
	
	
	sub UserMapUpdate(suspectid,vikarid)

	dim iSuspectID 	'string
	dim IVikarID 		'Consultant's id
	Dim objUserProxy  'Web service proxy for the DNN user service
	Dim sUserServiceURL 'Url of the user web service
	Dim iApp
	
	iSuspectID  = Cstr(suspectid)
     IVikarID   = Cstr(vikarid)
	iApp = Cstr(Application("Application"))
	sUserServiceURL = Application("DNNUserServiceURL")
	
	Set objUserProxy = Server.CreateObject("ECDNNUserProxy.UserServiceProxy")
	objUserProxy.Url = sUserServiceURL

		If Not objUserProxy.ConvertSuspectUserToVikar(IVikarID,iSuspectID ,iApp ) Then
            AddErrorMessage("Failed to save web user") // 1
			Call RenderErrorMessage()
	   End If
		
end sub

%>
