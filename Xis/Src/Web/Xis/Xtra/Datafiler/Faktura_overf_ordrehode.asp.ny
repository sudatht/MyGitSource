<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.utils.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Integration.Contact.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<%
	'--------------------------------
	' Endret 20051017 av MichalK:
	'   Om fakturaadresse (addressType=3) er definert for kontakten, brukes denne
	'   i stedet for postadresse.
	'--------------------------------
	
	Dim Avdeling 'avdelingsID
	dim Conn
	dim strSQL
	dim rsBnr
	Dim strWhereAvd 'Where-clause for utvalg p� avdeling
	dim kontakt
	dim kundenavn
	dim deresRef
	dim vaarRef	
	
	dim klientNr
	Dim OrdreDato
	dim DDato
	dim EkspGebyr
	dim MVAKode
	dim Prosjekt
	dim valgt_avd
	Dim cts 

	If (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	Function OnlyDigits( strString )
		' Remove all non-nummeric signs FROM string
		If Not IsNull(strString) Then
		For idx = 1 To Len( strString) Step 1
			Digit = Asc( Mid( strString, idx, 1 ) )
			If (( Digit > 47 ) And ( Digit < 58 )) Then
				strNewstring = strNewString & Mid(strString, idx,1)
			End If
		Next
		End If
		Onlydigits = strNewString
	End Function

	' Connect to database
	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))

	' Get data
	Klientnr = 1
	Ordredato = OnlyDigits(dbDate(DateValue(Date)))
	DDato = Left(Ordredato,4) + Right(Ordredato,2)
	Ordredato = DDato
	EkspGebyr = 0
	MVAKode = 0
	Prosjekt = 1

	if Request.QueryString("avd") <> "" then
		valgt_avd = CInt(Request.QueryString("avd"))
	else
		valgt_avd = 0
	end if

	'Lag where-clause for avdeling avh. av om det er valgt en avdeling eller ikke
	if valgt_avd > 0 then
		strWhereAvd = " AND [avdeling] = " & valgt_Avd & " "
	else
		strWhereAvd = ""
	end if

	'	Hent Buntnr
	strSQL = "SELECT NesteBnr = (MAX(Buntnr)+1) FROM [FAKTURAGRUNNLAG] WHERE [Status] = 2"
	Set rsBnr = GetFirehoseRS(strSQL, Conn)
	If IsNull(rsBnr("NesteBnr")) Then
		Buntnr = 1
	else
		Buntnr = rsBnr("NesteBnr")
	End If
	rsBnr.Close
	Set rsBnr = Nothing

	strSQL = "UPDATE [FAKTURAGRUNNLAG] SET [Buntnr] = " & Buntnr & " WHERE [Status] = 2" & strWhereAvd

	if (ExecuteCRUDSQL(strSQL, Conn) = false) then
		CloseConnection(Conn)
		set Conn = nothing	   
		AddErrorMessage("Feil under oppdatering av buntnr p� fakturagrunnlag.")
		call RenderErrorMessage()		
	end if

	'	Hent Ordrenr
	Ordrenr = Request.Form("ONR")

	'	Hent id-er fra FAKTURALINJER
	strSQL = "SELECT DISTINCT 	[F].[FirmaID], [K].[SOCuID], [K].[Firma],  " &_
		"[F].[Kontakt], DeresRef = ([Ko].[Fornavn] + ' ' + [Ko].[Etternavn]), [F].[SOKontakt],  " &_
		"[F].[Fakturanr], [F].[OppdragID], [F].[avdeling], vaaresRef = ([M].[Fornavn] + ' ' + [M].[Etternavn]), [O].[AnsmedID], [O].[SoPeID] " &_
		" FROM [FAKTURAGRUNNLAG] AS [F], [Firma] AS [K], [Kontakt] AS [Ko], [Oppdrag] AS [O], [MEDARBEIDER] AS [M] " &_
		" WHERE [F].[Status] = 2 " &_
		" AND [F].[FirmaID] = [K].[FirmaID] " & _
		" AND [F].[OppdragID] = [O].[OppdragID] " & _
		" AND [F].[Kontakt] *= [Ko].[KontaktID] " & _
		" AND [O].[AnsmedID] *= [M].[MedID] " & _
		strWhereAvd &_
		" ORDER BY [Fakturanr]"

	set rsID = GetFirehoseRS(strSQL, Conn)

	Set cts = server.CreateObject("Integration.SuperOffice")
	Conn.BeginTrans
	'	Loop for � overf�re til ordrehode
	do while NOT rsID.EOF

		KundeNR = rsID("FirmaID")
		HovedbokKtoNr = 0
		Avdeling = rsID("Avdeling")

		'	Hent Ordrenr
		If Request.Form("ONR") <> "" Then
			Ordrenr = Ordrenr + 1
			Ordrenummer = "'" & CStr(Ordrenr) & "'"
		Else
			Ordrenummer = "NULL"
		End If
	
		if(not isnull(rsID("SOCuID"))) then
			dim rsAddress
			set rsAddress = cts.GetAddressByContactId(clng(rsID("SOCuID")), 1)
			if (hasRows(rsAddress)) then
				if(IsNull(rsAddress("address1"))) then
					Adresse1 = "NULL"
				else
					Adresse1 = rsAddress("address1")
				end if
				if(IsNull(rsAddress("zipcode"))) then
					Postnummer = "NULL"
				else
					Postnummer = rsAddress("zipcode")
				end if
				if(IsNull(rsAddress("city"))) then
					Poststed = "NULL"
				else
					Poststed = rsAddress("city")
				end if	
				rsAddress.close							
			else
				Adresse1 = "NULL"
				Postnummer = "NULL"
				Poststed = "NULL"
			end if		
			set rsAddress = nothing
			
			' Sjekk om vi har en fakturaadresse som skal overstyre kontaktens adresse
			dim rsInvoiceAddress
			set rsInvoiceAddress = cts.GetAddressByContactId(clng(rsID("SOCuID")), 3)
			if (hasRows(rsInvoiceAddress)) then
				if( not IsNull(rsInvoiceAddress("zipcode")) and (rsInvoiceAddress("zipcode")<>"") ) then
				    if( not IsNull(rsInvoiceAddress("city")) and (rsInvoiceAddress("city")<>"") ) then
					    Postnummer = rsInvoiceAddress("zipcode")
					    Poststed = rsInvoiceAddress("city")
				    end if
				end if
				if( not IsNull(rsInvoiceAddress("address1")) and (rsInvoiceAddress("address1")<>"") ) then
				    Adresse1 = rsInvoiceAddress("address1")
				end if
			end if
			set rsInvoiceAddress = nothing
		end if
	
		'Customer name
		kundenavn = rsID("Firma")	
		'Name of contact at customer who ordered the task
		if isnull(rsID("Kontakt")) then
			kontakt = "NULL"		
			deresRef = GetSOPersonName(rsID("SOKontakt"))
			SOKontakt = rsID("SOKontakt")
		else
			deresRef = rsID("DeresRef")
			Kontakt = rsID("Kontakt")
			SOKontakt = "NULL"

		end if
		
		'Customers contact at Xtra
		vaarRef = Left(rsID("vaaresRef"), 30)
		
		'	Delete existing invoice export data
		strSQL = "DELETE FROM [EKSPORT_RUB_ORDRE] WHERE [Fakturanr] = " & rsID("Fakturanr")
	   if (ExecuteCRUDSQL(strSQL, Conn) = false) then
			Conn.RollbackTrans
			CloseConnection(Conn)
			set Conn = nothing	   
			AddErrorMessage("Feil under sletting av Faktura Ordrehode.")
			call RenderErrorMessage()		
	   end if
		
		strSQL = "INSERT INTO [EKSPORT_RUB_ORDRE](" &_
			"[Fakturanr], [Formatnummer], [Klientnummer], [Registernummer], " &_
			"[Status], [Ordretype], [Ordrenummer], [KundeNummer], " &_
			"[KundeNavn], [AdresseI], [Postnummer], [Poststed], [DeresRef], " &_
			"[VaarRef], [EkspGebyr], [MVAKode], [Prosjekt], [HovedbokKtoNr]," &_
			"[Eksportert_Dato], [Avdeling], [Kontakt], [SOKontakt], [FaktDatoType], [Buntnr]" &_
			") VALUES(" &_
			rsID("Fakturanr") &_
			",6000, 1, 6, 1, 1," &_
			Ordrenummer & ", " &_
			Kundenr & ",'" &_
			Left(Kundenavn, 30) & "','" &_
			Left(Adresse1, 30) & "','" &_
			Postnummer & "','" &_
			Left(Poststed, 25) & "','" &_
			Left(DeresRef, 30) & "','" &_
			Left(VaarRef, 30) & "'," &_
			EkspGebyr & "," &_
			MVAKode & "," &_
			Prosjekt & "," &_
			HovedbokKtoNr & "," &_
			"NULL," &_
			Avdeling & "," &_
			Kontakt & "," & _
			SOKontakt & "," & _
			"'O'," &_
 			Buntnr & ")"

		if (ExecuteCRUDSQL(strSQL, Conn) = false) then
			Conn.RollbackTrans
			CloseConnection(Conn)
			set Conn = nothing	   
			AddErrorMessage("Feil under oppdretting av faktura ordrehode.")
			call RenderErrorMessage()		
		end if
		rsID.MoveNext
	loop
	rsID.Close
	Set rsID = Nothing
	set cts = nothing
	Conn.CommitTrans
	CloseConnection(Conn)
	set Conn = nothing

	Response.Redirect "Faktura_overf_ordrelinje.asp?avd=" & valgt_avd
%>