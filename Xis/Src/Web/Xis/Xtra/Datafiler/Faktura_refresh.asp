<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<%
	If (HasUserRight(ACCESS_ADMIN, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim FRate
	dim art
	dim art2
	dim Conn
	dim strVikarID
	dim strOppdragID
	dim strFirmaID
	dim strStartDato
	dim strBestilltAv
	dim SOBestilltAv
	dim strKontakt
	dim SOKontakt
	dim strSQL
	dim kontaktSQL
	dim ErXisContact
	dim rsFaktgr
	dim rsFaktgr2

	'Lagring av timelisten
	Sub fjernKomma(strString )
		pos = Instr(strString, ",")
		If pos <> 0 Then
			mellom = Left(strString, pos-1) & "." & Left(Mid(strString, pos+1), 2)
 			strString = mellom
		End If
		'Response.Write strString
	End Sub

	' Connect to database
	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))

	' parameters
	' Check parameters AND put into variables
	strVikarID = Request.QueryString("VikarID")
	strOppdragID = Request.QueryString("OppdragID")
	strFirmaID = Request.QueryString("FirmaID")
	strStartDato = Request.QueryString("startDato")
	strBestilltAv = Request.QueryString("BestilltAv")
	SOBestilltAv = Request.QueryString("SOBestilltAv")
	
	'FAKTURAGRUNNLAG
	if (len(strBestilltAv) > 0) then	
		strSQL = "DELETE FROM FAKTURAGRUNNLAG " &_
			"WHERE Kontakt = " & strBestilltAv &_
			" AND VikarID = " & strVikarID &_
			" AND Status < 3"
	else
		strSQL = "DELETE FROM FAKTURAGRUNNLAG " &_
			"WHERE SOKontakt = " & SOBestilltAv &_
			" AND VikarID = " & strVikarID &_
			" AND Status < 3"	
	end if
	'Start transaction
	Conn.Begintrans
	
	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under sletting av fakturagrunnlag.")
		call RenderErrorMessage()	
	End if

	'Prosesser tidligere dags/uke lister
	dim strFaktTekst 'Variabel for å teste lengde av streng til fakturalinje
	redim oppdrag(20) 'Inneholder tekst med oppdragsnummer

	if (len(strBestilltAv) > 0) then	
		kontaktSQL = " U.BestilltAv = " & strBestilltAv
	else
		kontaktSQL = " U.SOBestilltAv = " & SOBestilltAv
	end if

	'Get invoice data
	strSQL = "SELECT U.VikarID, OppdragID, Ant=sum(Fakturatimer), " &_
		"Fakturapris, fakturaType, LinjeSum=Sum(FakturaBeloep), " &_
		"BestilltAv, SOBestilltAv, Navn=(Fornavn +' ' + Etternavn), Dato, FraDato " &_
		"FROM VIKAR_UKELISTE U, VIKAR " &_
		" WHERE" &_
		kontaktSQL &_
		" AND U.VikarID = VIKAR.VikarID" &_
		" AND U.VikarID = " & strVikarID &_
		" AND Overfort_fakt_status < 3" &_
		" AND U.StatusId > 4" &_
		" Group by fakturaType, Fakturapris, U.VikarID, U.OppdragID, BestilltAv, SOBestilltAv, ForNavn, Etternavn, Dato, Fradato " &_
		" Order by U.VikarID desc, fakturaType desc"

		'" AND U.Fakturatimer > 0.0" &_
		'" AND U.FakturaBeloep > 0.0" &_

	Set rsFaktgr = GetFireHoseRS(strSQL, Conn)

	if not HasRows(rsFaktgr) Then
		Conn.RollBackTrans
		AddErrorMessage("Ingen registrerte ordre.")
		call RenderErrorMessage()
	end if

	' Deleting from FAKTURAGRUNNLAG If exist rows on this id
	strKontakt = rsFaktgr("BestilltAv")
	SOKontakt = rsFaktgr("SOBestilltAv")
	If IsNull(strKontakt) Then
		strKontakt = "NULL"
		ErXisContact = false
	else
		SOKontakt = "NULL"
		ErXisContact = true
	end if

	' Insert into FAKTURAGRUNNLAG

	redim 	oppdrag(20) 	'Inneholder tekst med oppdragsnummer
	dim 	strArtNr		'Holder artikkelnummer og tekst
	dim 	rsTom			'Holder tjenesteområdeid for oppdrag / kurs
	dim 	IntLinjeNr		'Tellevariabel som brukes for å sette sortering på fakturalinjer = Linjenummer
	dim 	IntAntOppdrag	'Tellevariabel som teller antall oppdrag som blir lagt inn på en fakturalinje. Tekst må være < 40 tegn
	dim 	IntOppdLinjer	'Holder orden paa alle oppdragene for linjene
	dim 	strOppdrag
	dim 	strAllOppdrag

	IntLinjeNr = 0
	IntAntOppdrag = 1
	IntOppdLinjer = 1

	strOppdrag = ""
	strAllOppdrag = ""


	VID = rsFaktgr("VikarID")
	OID = rsFaktgr("OppdragID")
	oppdrag(1) = OID

	do while not rsFaktgr.EOF

		'Sjekker hvilken avdeling dette er (kurs har eget opplegg).
		strSQL = "SELECT AvdelingID FROM OPPDRAG WHERE OppdragID = " & rsFaktgr("OppdragID")
		Set rsID = GetFirehoseRS(strSQL, Conn)

		If (Not HasRows(rsID)) Then
			Conn.RollBackTrans
			set rsid = nothing
			rsFaktgr.close
			set rsFaktgr = nothing
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Oppdraget har ikke tilknyttet noen avdeling!")
			call RenderErrorMessage()
		end if

		VType  = rsID("AvdelingID")
		Avdeling = rsID("AvdelingID")
		Art = rsFaktgr("fakturaType")
		IntLinjeNr = IntLinjeNr + 1

		'slår inn hvis det er en kursvikar
		If ((VType = 1) AND (instr(strOppdrag, rsFaktgr("OppdragID")) = 0)) Then

			if (ErXisContact) then	
				kontaktSQL = " U.BestilltAv = " & strKontakt
			else
				kontaktSQL = " U.SOBestilltAv = " & SOKontakt
			end if		
		
			strSQL = "SELECT U.VikarID, Ant=sum(Fakturatimer), " &_
				"Fakturapris, fakturaType, LinjeSum=Sum(FakturaBeloep), " &_
				"BestilltAv, SOBestilltAv, Navn=(Fornavn +' ' + Etternavn), Dato, FraDato " &_
				" FROM VIKAR_UKELISTE U, VIKAR " &_
				" WHERE" &_
				kontaktSQL &_
				" AND U.VikarID = VIKAR.VikarID" &_
				" AND U.Fakturatimer > 0" &_				
				" AND U.VikarID = " & strVikarID &_
				" AND Overfort_fakt_status = 1" &_
				" AND U.oppdragid = " & rsFaktgr("OppdragID") &_
				" AND U.StatusId > 4" &_
				" AND U.FakturaType IS NOT NULL" &_									
				" Group by fakturaType, Fakturapris, U.VikarID, BestilltAv, SOBestilltAv, ForNavn, Etternavn, Dato, Fradato " & _
				"Order by U.VikarID desc, fakturaType desc"

				'" AND U.Fakturatimer > 0.0" &_
				'" AND U.FakturaBeloep > 0.0" &_

			Set rsFaktgr2 = GetFirehoseRS(strSQL, Conn)

			'Dette er kurs - tjenesteomraade = 1
			strSQL = "SELECT DISTINCT o.tomid " & _
				" FROM VIKAR_UKELISTE U, VIKAR, oppdrag O " & _
				" WHERE O.oppdragid = U.oppdragid " & _
				" AND U.oppdragid = " & rsFaktgr("OppdragID") &_
				" AND" & kontaktSQL &_
				" AND U.VikarID = VIKAR.VikarID" &_
				" AND U.VikarID = " & strVikarID &_
				" AND Overfort_fakt_status = 1" &_
				" AND U.StatusId > 4"

			Set rsTom = GetFirehoseRS(strSQL, Conn)

			if rsTom.recordcount > 1 then
				rsTom.close
				set rsTom = nothing
				CloseConnection(Conn)
				set Conn = nothing						
				AddErrorMessage("Ugyldig tjenesteomraade for kurs!")
				AddErrorMessage("Kunne ikke generere faktura!")
				call RenderErrorMessage()				
			end if

			do while Not rsFaktgr2.EOF

				Art2 = rsFaktgr2("fakturaType")

				'Hardcoded to "andre" for the following toms: 7,9,10,11,12,13,14,15 and 16
				if (clng(rsTom("tomid")) = 7 OR clng(rsTom("tomid")) = 9 OR clng(rsTom("tomid")) = 10 OR clng(rsTom("tomid")) = 11 OR clng(rsTom("tomid")) = 12 OR clng(rsTom("tomid")) = 13 OR clng(rsTom("tomid")) = 14 OR clng(rsTom("tomid")) = 15 OR clng(rsTom("tomid")) = 16) then
					strArtNr = "301008" 
				else
					strArtNr = cstr( ( 301000 + clng(rsTom("tomid")) ) )
				end if

				strSQL = "SELECT [Fakturabeskrivelse], [FakturaSats] FROM [H_FAKTURA_TYPE] WHERE [FakturaType] = " & Art2

				set rsFakturaSats = GetFirehoseRS(strSQL, Conn)
				
				FRate = rsFakturaSats("FakturaSats")
				if (FRate = 1.0) then 'ordinære timer
					strTekst = rsFaktgr("Navn") & " - " & rsFakturaSats("Fakturabeskrivelse")
				elseif (FRate > 1.0) then 'overtid
					strTekst = rsFaktgr("Navn") & " - " & rsFakturaSats("Fakturabeskrivelse")
					strArtNr = strArtNr & "-" & cstr(((CDbl(FRate)-1) * 100))
				end if
				rsFakturaSats.close
				set rsFakturaSats = nothing

				strLinjeSum = rsFaktgr2("LinjeSum")
				If IsNull(rsFaktgr2("Ant")) Then
					strAntall = 0
				Else
					strAntall = rsFaktgr2("Ant")
				End IF
				strFaktgr = rsFaktgr2("Fakturapris")
				Call fjernKomma(strLinjeSum)
				Call fjernKomma(strAntall)
				Call fjernKomma(strFaktgr)

				strSQL = "INSERT INTO FAKTURAGRUNNLAG " &_
					"(VikarID, OppdragID, FirmaID, Kontakt, SOKontakt, Artikkelnr, Tekst, Antall, Enhetspris, LinjeSum, Status, LinjeNr, Avdeling) " &_
					"VALUES (" &_
					rsFaktgr("VikarID") & ", " &_
					rsFaktgr("OppdragID") & ", " &_
					strFirmaID & ", "  &_
					strKontakt & "," &_
					SOKontakt & ",'" &_
					strArtNr & "','" &_
					strTekst & "'," &_
					strAntall & "," &_
					strFaktgr & "," &_
					strLinjeSum & ", 1 , " &_
					IntLinjeNr & "," &_
					Avdeling & ")"

				If ExecuteCRUDSQL(strSQL, Conn) = false then
					Conn.RollBackTrans
					CloseConnection(Conn)
					set Conn = nothing
					AddErrorMessage("En feil oppstod under oppretting av fakturagrunnlag.")
					call RenderErrorMessage()	
				End if

				fraDato = rsFaktgr("FraDato")
				tilDato = rsFaktgr("Dato")
				rsFaktgr2.MoveNext
				IntLinjeNr = IntLinjeNr + 1

				strSQL = "SELECT AvdelingID FROM OPPDRAG WHERE OppdragID = " & rsFaktgr("OppdragID")

				Set rsID = GetFirehoseRS(strSQL, Conn)
				Avdeling = rsID("AvdelingID")
				rsID.Close
				Set rsID = Nothing
			loop
			strOppdrag = strOppdrag & "," & rsFaktgr("OppdragID")
		End If 'VType = 1
		'slutt kursvikar

		If VType > 1 Then

			strSQL ="SELECT [tomid] FROM [oppdrag] WHERE [oppdragid] =" & rsFaktgr("OppdragID")
			Set rsTom = GetFirehoseRS(strSQL, Conn)

			'Hardcoded to "andre" for the following tjenesteområde ids: 7,9,10,11,12,13,14,15 and 16
			if (clng(rsTom("tomid")) = 7 OR clng(rsTom("tomid")) = 9 OR clng(rsTom("tomid")) = 10 OR clng(rsTom("tomid")) = 11 OR clng(rsTom("tomid")) = 12 OR clng(rsTom("tomid")) = 13 OR clng(rsTom("tomid")) = 14 OR clng(rsTom("tomid")) = 15 OR clng(rsTom("tomid")) = 16) then
				strArtNr = "301008" 
			else
				strArtNr = cstr( ( 301000 + clng(rsTom("tomid")) ) )
			end if

			strSQL = "SELECT [Fakturabeskrivelse], [FakturaSats] FROM [H_FAKTURA_TYPE] WHERE [FakturaType] = " & Art
			set rsFakturaSats = GetFirehoseRS(strSQL, Conn)
			
			FRate = rsFakturaSats("FakturaSats")
			if (FRate = 1.0) then
				strTekst = rsFaktgr("Navn") & " - " & rsFakturaSats("Fakturabeskrivelse")
			elseif (FRate > 1.0) then
				strTekst = rsFaktgr("Navn") & " - " & rsFakturaSats("Fakturabeskrivelse")
				strArtNr = strArtNr & "-" & cstr(((CDbl(FRate)-1) * 100))
			end if
			rsFakturaSats.close
			set rsFakturaSats = nothing

			strLinjeSum = rsFaktgr("LinjeSum")
			If IsNull(rsFaktgr("Ant")) Then
				strAntall = 0
			Else
				strAntall = rsFaktgr("Ant")
			End IF
			strFaktgr = rsFaktgr("Fakturapris")
			Call fjernKomma(strLinjeSum)
			Call fjernKomma(strAntall)
			Call fjernKomma(strFaktgr)

			strSQL = "INSERT INTO FAKTURAGRUNNLAG " &_
			"(VikarID, OppdragID, FirmaID, Kontakt, SOKontakt, Artikkelnr, Tekst, Antall, Enhetspris, LinjeSum, Status, LinjeNr, Avdeling) " &_
			"values (" &_
			rsFaktgr("VikarID") & ", " &_
			rsFaktgr("OppdragID") & ", " &_
			strFirmaID & ", "  &_
			strKontakt & "," &_
			SOKontakt & ",'" &_
			strArtNr & "','" &_
			strTekst & "'," &_
			strAntall & "," &_
			strFaktgr & "," &_
			strLinjeSum & ", 1 , " &_
			IntLinjeNr & "," &_
			Avdeling & ")"
			
			If ExecuteCRUDSQL(strSQL, Conn) = false then
				Conn.RollBackTrans
				CloseConnection(Conn)
				set Conn = nothing
				AddErrorMessage("En feil oppstod under oppretting av fakturagrunnlag.")
				call RenderErrorMessage()	
			End if
			
			fraDato = rsFaktgr("FraDato")
			tilDato = rsFaktgr("Dato")

		End If 'VType > 1

		If ((rsFaktgr("OppdragID") <> OID) AND (Instr(strAllOppdrag, trim(rsFaktgr("OppdragID")))=0)) Then

			IntAntOppdrag = IntAntOppdrag + 1

			'Sjekker om teksten blir for lang
			strFaktTekst = oppdrag(IntOppdLinjer)& "," & rsFaktgr("OppdragID")
			

			If (IntOppdLinjer = 1 AND len(strFaktTekst) > 32) or (len(strFaktTekst) > 40) then 'Max 40 ('Avtale  ' kommer i tillegg)
				IntOppdLinjer = IntOppdLinjer + 1
				IntAntOppdrag = 1
			End if

			If IntOppdLinjer = 1 Then
				oppdrag(IntOppdLinjer) = oppdrag(IntOppdLinjer) & "," & rsFaktgr("OppdragID")
			Else
				If IntAntOppdrag = 1 Then
					oppdrag(IntOppdLinjer) = rsFaktgr("OppdragID")
				Else
					oppdrag(IntOppdLinjer) = oppdrag(IntOppdLinjer) & "," & rsFaktgr("OppdragID")
				End If
			End If

			OID = rsFaktgr("OppdragID")
		End If
		strAllOppdrag = strAllOppdrag & "," & OID
 		rsFaktgr.MoveNext
	loop

	IntLinjeNr = IntLinjeNr + 1

	strSQL = "SELECT strTekst = (CASE tom.rub_Mvakode WHEN '00' THEN 'Opplæring ' WHEN '01' then 'Vikar ' ELSE 'Opplæring ' END), AvdelingID " & _
		"FROM tjenesteomrade tom, oppdrag o " & _
		"WHERE tom.tomid = o.tomid " & _
		"AND oppdragid = " & OID

	Set rsID = GetFirehoseRS(strSQL, Conn)
	strTekst = rsID("strTekst") & fraDato & " - " & tilDato
	Avdeling = rsID("AvdelingID")
	rsID.Close
	Set rsID = Nothing

	strSQL = "INSERT INTO FAKTURAGRUNNLAG " &_
		"(VikarID, OppdragID, FirmaID, Kontakt, SoKontakt, Tekst, Status, LinjeNr, Avdeling) " &_
		"VALUES (" &_
		VID & ", " &_
		OID & ", " &_
		strFirmaID & ", "  &_
		strKontakt & "," &_
		SOKontakt & ",'" &_
		strTekst & "'," &_
		" 1 , " &_
		IntLinjeNr & "," &_
		Avdeling & ")"

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under oppretting av fakturagrunnlag.")
		call RenderErrorMessage()	
	End if

	IntLinjeNr = IntLinjeNr + 1

	For i = 1 To IntOppdLinjer

		If i = 1 Then
			strTekst = "Oppdrag " & oppdrag(i)
		End If

		strSQL = "INSERT INTO FAKTURAGRUNNLAG " &_
		"(VikarID, OppdragID, FirmaID, Kontakt, SOKontakt, Tekst, Status, LinjeNr, Avdeling) " &_
		"VALUES (" &_
		VID & ", " &_
		OID & ", " &_
		strFirmaID & ", "  &_
		strKontakt & "," &_
		SOKontakt & ",'" &_
		strTekst & "'," &_
		" 1 , " &_
		IntLinjeNr & "," &_
		Avdeling & ")"

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollBackTrans
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("En feil oppstod under oppretting av fakturagrunnlag.")
			call RenderErrorMessage()	
		End if
	Next
	Conn.CommitTrans
	CloseConnection(Conn)
	set Conn = nothing

	Redir = "Vikar_timeliste_vis3.asp?VikarID=" & strVikarID & "&OppdragID=" & strOppdragID & "&FirmaID=" & strFirmaID & "&startdato=" & strStartdato & "&splittopt=" & session("splittopt") & "&splittuke=" & session("splittuke")
	Response.Redirect Redir
%>