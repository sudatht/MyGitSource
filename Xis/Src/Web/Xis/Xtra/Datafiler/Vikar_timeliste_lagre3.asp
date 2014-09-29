<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<!--#INCLUDE FILE="../includes/Xis.Economics.Constants.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

session("debug") = true

	'Lagring av timelisten
	dim strProjektNr
	dim strLoennsartNr
	dim strFaktAntSum
	dim strSats
	dim fakturaTypeSQL
	dim rsFakturaType
	dim fakturaType
	dim Conn
	dim strFaktAntall : strFaktAntall = 0
	dim rsFakturaSats
	dim FRate
	dim VikarTStatus
	dim startdato
	dim sluttdato
	dim splittopt
	dim strSQL
	dim contactSQL
	dim strSOBestilltAv
	dim strBestilltAv
	dim ErXisContact
	dim strMiddleOpp
    dim strNote
    
	function FixNumericQueryString(queryString)
		dim value
		value = request(queryString)
		if (lenB(trim(value)) = 0) then
			FixNumericQueryString = "0"
		else
			FixNumericQueryString = value		
		end if
	end function

	' Connect to database
	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))

	' Check parameters AND put into variables
	strVikarID = Request("VikarID")
	strFirmaID = Request("FirmaID")
	strOppdragID = Request("OppdragID")
	strTimeLonn = Request("TimeLonn")
	strUkeNr = Request("UkeNr")
	strDato = DbDate(Request("Dato"))
	strStartDato = Request("startDato")
	lagringskode = Request("lagringskode")


	strBestilltAv = Request("BestilltAv")
	If len(strBestilltAv) = 0 Then
		strBestilltAv = "NULL"
		ErXisContact = false
	end if

'Response.Write "strBestilltAv:" & strBestilltAv & "<br>"
'Response.Write "ErXisContact:" & ErXisContact & "<br>"

	
	strSOBestilltAv = Request("SOBestilltAv")
	If len(strSOBestilltAv) = 0 Then
		strSOBestilltAv = "NULL"
		ErXisContact = true
	end if
'Response.Write "strSOBestilltAv:" & strSOBestilltAv & "<br>"
'Response.Write "ErXisContact:" & ErXisContact & "<br>"

	strFakturapris = Request("Fakturapris")
	strLoennsartNr = Request("LoennsartNr")
	strTilStatus= Request("tilStatus")
	strAntSum = Request("Summ")
	strAntallLoennedeSteg1 = Request("Pr50")
	strAntallLoennedeSteg2 = Request("Pr100")
	if len(Request("FAKTSUMM")) > 0 then
		strFaktAntall = Request("FAKTSUMM")
	end if
	strAntallFakturerteSteg1 = Request("PRFAKT50")
	strAntallFakturerteSteg2 = Request("PRFAKT100")
	strLoennsArt = Request("SYK")

	'Sjekker hvor brukeren kommer fra
	'Hvis konsulentledere godkjenner, skal listen få status 4, hvis økonomi godkjenner skal listen
	'få status 5
	If CInt(session("frakode")) = 1 Or CInt(session("frakode")) = -1 Then
		VikarTStatus = 4
	Else 
		VikarTStatus = 5
	End If
Response.Write "VikarTStatus:" & VikarTStatus & "<br>"
	'Er uken splittet?
	splittopt = Request("splittopt")
	If splittopt = "Null" then
		splittopt = Null
	End If

	startdato = session("startdato")
	sluttdato = session("sluttdato")

'Nedgradering av timeliste (til status 1)
If lagringskode = "nedgrad" Then

	'Start transaction
	Conn.Begintrans
	
	' Update status in DAGSLISTE_VIKAR    (timelisten)
	if (strTilStatus) <> 2 then
		strTilStatus = 1
	end if

	strSQL = "UPDATE DAGSLISTE_VIKAR SET" &_
		" TimelisteVikarStatus =" & strTilStatus & "," &_
		" Loennstatus = 1," &_
		" Fakturastatus = 1," &_
		" Loenndato = NULL," &_
		" Fakturadato = NULL" &_
		" WHERE VikarID = " & strVikarID &_
		" AND OppdragID = " & strOppdragID &_
		" AND TimelisteVikarStatus < 6" &_
		" AND Fakturastatus < 3" &_
		" AND Loennstatus < 3" &_
		" AND Dato >= " & DbDate(startdato) &_
		" AND Dato <= " & Dbdate(sluttdato)
		
	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under oppdatering av dagsliste for vikar.")
		call RenderErrorMessage()	
	End if

	strSQL = "DELETE FROM VIKAR_UKELISTE " &_
		" WHERE OppdragID = " & strOppdragID &_
		" AND StatusID < 6" &_
		" AND VikarID = " & strVikarID &_
		" AND Ukenr = " & strUkeNr &_
		" AND Notat = '" & splittopt & "'"

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under sletting fra ukeliste for vikar.")
		call RenderErrorMessage()	
	End if
	
'ikke nedgradering men vanlige delete, update eller insert
Else
	'Start transaction
	Conn.Begintrans

	' Update status in DAGSLISTE_VIKAR    (timelisten)
	strSQL = "UPDATE DAGSLISTE_VIKAR SET " &_
		"TimelisteVikarStatus = " & VikarTStatus &_
		" WHERE VikarID = " & strVikarID &_
		" AND OppdragID = " & strOppdragID &_
		" AND TimelisteVikarStatus < 6" &_
		" AND Dato >= " & DbDate(startdato) &_
		" AND Dato <= " & Dbdate(sluttdato)

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under oppdatering av dagsliste for vikar.")
		call RenderErrorMessage()	
	End if

	' Delete row in VIKAR_UKELISTE (if they exist)
	If IsNull(splittopt) Then 'Uken er ikke splittet

		strSQL = "DELETE FROM VIKAR_UKELISTE " &_
			"WHERE OppdragID = " & strOppdragID &_
			 " AND VikarID = " & strVikarID &_
			 " AND Ukenr = " & strUkeNr &_
			 " AND Overfort_Loenn_status < 3" &_
			 " AND StatusID < 6"

	Else 'Uken er splittet
		strSQL = "DELETE FROM VIKAR_UKELISTE " &_
			"WHERE OppdragID = " & strOppdragID &_
			 " AND VikarID = " & strVikarID &_
			 " AND Ukenr = " & strUkeNr &_
			 " AND (Notat = '" & splittopt &_
			 "' or Notat IS NULL or Notat = ' ' or Notat LIKE 'NULL')" &_
			 " AND Overfort_Loenn_status < 3" &_
			 " AND StatusID < 6"
	End IF

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under sletting fra ukeliste for vikar.")
		call RenderErrorMessage()	
	End if

	if(ErXisContact = true) then
		contactSQL = " BestilltAV = " & strBestilltAV & " " 'old xis contact
	else
		contactSQL = " SOBestilltAV = " & strSOBestilltAv & " " 'superoffice contact
	end if

	' Update VIKAR_UKELISTE (enddate on previous rows with same kontakt to get right date on invoice)
	strSQL = "SELECT ddd = Min(Dato) From DAGSLISTE_VIKAR" &_
		" WHERE " &_
		contactSQL &_
		" AND VikarID = " & strVikarID &_
		" AND Fakturastatus < 3"

	Set rsDD = GetFirehoseRS(strSQL, Conn)

	SDato = rsDD("ddd")
	SDato = dbDate(SDato)

	oldEndDate = session("Enddato")
	oldEndDate = dbDate(oldenddate)

	strSQL = "UPDATE VIKAR_UKELISTE SET " &_
		"Dato = " & oldEndDate &_
		", FraDato = " & SDato &_
		" WHERE " &_ 
		contactSQL &_
		" AND VikarID = " & strVikarID &_
		" AND Overfort_fakt_status < 3"

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under oppdatering av ukeliste for vikar.")
		call RenderErrorMessage()	
	End if

	' Insert into VIKAR_UKELISTE
	strSQL = "SELECT VikarID, TimelisteVikarStatus, FirmaID, TimeLonn, " &_
		"Fakturapris, Bestilltav, SOBestilltav, " &_
		"Fakturastatus, Notat, Loennstatus, " &_
		"Ltim=Sum(AntTimer), Faktt=Sum(Fakturatimer), Loennsart " &_
		"FROM DAGSLISTE_VIKAR " &_
		"WHERE TimelisteVikarStatus < 6"  &_
		" AND Loennstatus < 3" &_
		" AND VikarID = " & strVikarID &_
		" AND OppdragID = " & strOppdragID &_
		" AND Dato >= " & DbDate(startdato) &_
		" AND Dato <= " & Dbdate(sluttdato) &_
		" GROUP BY VikarID, TimelisteVikarStatus, FirmaID, TimeLonn, Fakturapris, " &_
		" Bestilltav, SOBestilltav, Fakturastatus, LoennsArt, Fakturastatus, Notat, Loennstatus" &_
		" ORDER BY Loennsart"
Response.Write "Timeliste grunnlag:<br>"
	Set rsVikar = GetFirehoseRS(strSQL, Conn)
	
	teller = 0

	Do while Not rsVikar.EOF
		strBegDato = SDato
		strEndDato = dbDate(session("EndDato"))

		' Insert Sum into VIKAR_UKELISTE
		' Endret 11.01.2000 (LWS) Overtid trekkes fra normaltid sålangt det er mulig, og deretter fra
		' andre lønnsarter. Rettet for å fikse feil der en uke med kun etterfakturering og overtid ikke
		' ga riktig tall i variabel lønn.

		strAntSum = rsVikar("LTim")
		strFaktAntSum = rsVikar("Faktt")
		
		'check Notat for bonus update
		if ( rsVikar("LoennsArt")="EF" ) then
		    if ""<>rsVikar("Notat") then
		        strNote = rsVikar("Notat")
		    else
		        strNote = "TILLEGG"
		    end if
		        
		end if

		strFaktSort = 1
		teller = teller + 1

		dim loennartSQL 
		dim loennsartKode
		dim rsLoennartInfo
		dim loennartRate
		dim skalFaktureres
		
		If IsNull(rsVikar("LoennsArt")) Then
			loennsartKode = ""
		else
			loennsartKode = rsVikar("LoennsArt")
		end if

		loennartSQL = "SELECT [LoennsartNr], [LoennRate], [Faktureres] FROM H_LOENNSART WHERE [LoennsArtKode] = '" & loennsartKode & "'"

		Set rsLoennartInfo = GetFirehoseRS(loennartSQL, Conn)

		If not rsLoennartInfo.EOF then
			
			strLoennsartNr = rsLoennartInfo("LoennsartNr") 'Løennsartsnr
			loennartRate = rsLoennartInfo("LoennRate") 'Løennsartrate 1 = 100%, 1.5 = 150% osv
			skalFaktureres = rsLoennartInfo("Faktureres") 'Hvorvidt denne lønnarten kan faktureres kunden
			
			if (skalFaktureres = false) then 'Denne lønnarten skal ikke faktureres
				strFaktAntall = 0 
				fakturaType = "NULL"
			else
				fakturaTypeSQL = "SELECT [FakturaType] FROM [H_FAKTURA_TYPE] WHERE [FakturaSats] = '" & loennartRate & "'"
				set rsFakturaType = GetFirehoseRS(fakturaTypeSQL, Conn)
				fakturaType = rsFakturaType("FakturaType")
				rsFakturaType.close
				set rsFakturaType = nothing			
			end if
	
			if strLoennsartnr = 180 then 'Stygg hardkoding for korreksjon, skal ha egen faktura linje
				If strFaktAntSum <> 0 Then
					strFaktSort = 4
					fakturaType = 8
				End If
			end if
		else
			Conn.RollBackTrans
			CloseConnection(Conn)
			set Conn = nothing		
			AddErrorMessage("Ugyldig lønnsartkode!")
			call RenderErrorMessage()				
		end if

		If strAntallLoennedeSteg1 <> "" Then 'Antall lønnede timer for overtid steg 1
			IF (strAntSum - strAntallLoennedeSteg1) > 0 then
				strAntSum = strAntSum - strAntallLoennedeSteg1
				strAntallLoennedeSteg1 = 0
			Else
				strAntallLoennedeSteg1 = strAntallLoennedeSteg1 - strAntSum
				strAntSum = 0
			End if
		End If

		If strAntallLoennedeSteg2 <> "" Then 'Antall lønnede timer for overtid steg 2
			IF (strAntSum - strAntallLoennedeSteg2) > 0 then
				strAntSum = strAntSum - strAntallLoennedeSteg2
				strAntallLoennedeSteg2 = 0 'for next loop
			Else
				strAntallLoennedeSteg2 = strAntallLoennedeSteg2 - strAntSum
				strAntSum = 0
			End if
		End If

		If strAntallFakturerteSteg1 <> "" Then 'Antall fakturerte timer for overtid steg 1
			If cdbl(strAntallFakturerteSteg1) > 0 Then 
				strFaktAntSum = strFaktAntSum - strAntallFakturerteSteg1
				strAntallFakturerteSteg1 = 0 'for next loop
			end if
		End If

		If strAntallFakturerteSteg2 <> "" Then 'Antall fakturerte timer for overtid steg 2
			If cdbl(strAntallFakturerteSteg2) > 0 Then 
				strFaktAntSum = strFaktAntSum - strAntallFakturerteSteg2
				strAntallFakturerteSteg2 = 0 'for next loop
			end if
		End If

		pos1 = InStr(1, strAntSum, ".")
		If pos1 > 0 Then
			strAntall = Left(strAntSum, pos1 - 1) & "," & Mid(strAntSum, pos1 + 1)
		Else
			strAntall = strAntSum
		End If

		strSats = rsVikar("Timelonn")
		strBeloep = strSats * strAntall

		pos1 = InStr(1, strFaktAntSum, ".")
		If pos1 > 0 Then
			strFaktAntall = Left(strFaktAntSum, pos1 - 1) & "," & Mid(strFaktantSum, pos1 + 1)
		Else
			strFaktAntall = strFaktAntSum
		End If

		FakturaStatus = rsVikar("Fakturastatus")
		LoennStatus = rsVikar("Loennstatus")
		strFaktsats = rsVikar("Fakturapris")
		strFaktBeloep = strFaktsats * strFaktAntall
		strFaktSort = 1
		Call fjernKomma(strAntall)
		Call fjernKomma(strSats)
		Call fjernKomma(strBeloep)
		Call fjernKomma(strFaktAntall)
		Call fjernKomma(strFaktsats)
		Call fjernKomma(strFaktBeloep)

		if not VikarTStatus = 4 then

			strSQL = "INSERT INTO VIKAR_UKELISTE " &_
				"(VikarID, OppdragID, FirmaID, UkeNr, Loennsartnr, Antall, Sats, Belop, Notat, Dato, StatusID, " &_
				"Overfort_loenn_status, " &_
				"Overfort_fakt_status, BestilltAv, SOBestilltav, Fakturapris, Fakturatimer, Fakturabeloep, FaktSort, FraDato, FakturaType) " &_
				"values (" &_
				strVikarID & "," &_
				strOppdragID & "," &_
				strFirmaID & "," &_
				strUkeNr & "," &_
				strLoennsartnr & "," &_
				strAntall & "," &_
				strSats & "," &_
				strBeloep & ",'" &_
				splittopt & "'," &_
				strEndDato & "," &_
				VikarTStatus & "," &_
				LoennStatus & "," &_
				FakturaStatus & "," &_
				strBestilltAv & "," &_
				strSOBestilltAv & "," &_
				strFaktsats & "," &_
				strFaktAntall & "," &_
				strFaktBeloep & "," &_
				strFaktSort & "," &_
				strBegDato & "," &_
				fakturaType & ")"
Response.Write "INSERT INTO VIKAR_UKELISTE 1:<br>"
			If ExecuteCRUDSQL(strSQL, Conn) = false then
				Conn.RollBackTrans
				CloseConnection(Conn)
				set Conn = nothing
				AddErrorMessage("En feil oppstod under oppretting av ukeliste for vikar.")
				call RenderErrorMessage()	
			End if
			
		end if 'if not VikarTStatus=4 then
		rsVikar.MoveNext
	loop

	rsVikar.Close

	' Sett inn overtid for steg 1 i ukeliste_vikar
	strLoennsartnr = "160" 'Denne er hardkodet for overtid steg 1 	
	strAntall = FixNumericQueryString("PR50")
	strFaktAntall = FixNumericQueryString("PRFAKT50")
Response.Write "strAntall:" & strAntall & "<br>"
Response.Write "strFaktAntall:" & strFaktAntall & "<br>"
	If cdbl(strAntall) > 0 OR cdbl(strFaktAntall) > 0 Then

		'Hent rate for denne lønnsarten
		loennartSQL = "SELECT [LoennRate] FROM [H_LOENNSART] WHERE [LoennsartNr] = '" & strLoennsartnr & "'"

		Set rsLoennartInfo = GetFirehoseRS(loennartSQL, Conn)	
		strSats = strTimelonn * rsLoennartInfo("LoennRate")
		rsLoennartInfo.close
		set rsLoennartInfo = nothing
		
		If strAntall > 0 Then
			strBeloep = strSats * strAntall
		Else
			strAntall = 0
		End If

		strFaktsats = strFakturapris * Request("lstRateSteg1") 

		strFaktBeloep = strFaktsats * strFaktAntall
				
		fakturaTypeSQL = "SELECT [FakturaType] FROM [H_FAKTURA_TYPE] WHERE [FakturaSats] = '" & Replace(Request("lstRateSteg1"),",",".") & "'"

		set rsFakturaType = GetFirehoseRS(fakturaTypeSQL, Conn)
		fakturaType = rsFakturaType("FakturaType")
		rsFakturaType.close
		set rsFakturaType = nothing
		
		Call fjernKomma(strAntall)
		Call fjernKomma(strSats)
		Call fjernKomma(strBeloep)
		Call fjernKomma(strFaktAntall)
		Call fjernKomma(strFaktsats)
		Call fjernKomma(strFaktBeloep)

		strFaktSort = 2

		strSQL = "INSERT INTO VIKAR_UKELISTE " &_
			"(VikarID, OppdragID, FirmaID, UkeNr, Loennsartnr, Antall, Sats, Belop, Notat, Dato, StatusID, Overfort_loenn_status, " &_
			"Overfort_fakt_status, BestilltAv, SOBestilltav, Fakturapris, Fakturatimer, Fakturabeloep, FaktSort, FraDato, FakturaType) " &_
			"values (" &_
			strVikarID & "," &_
			strOppdragID & "," &_
			strFirmaID & "," &_
			strUkeNr & "," &_
			strLoennsartnr & "," &_
			strAntall & "," &_
			strSats & "," &_
			strBeloep & ",'" &_
			splittopt & "'," &_
			strEndDato & "," &_
			VikarTStatus & "," &_
			" 1,1, " &_
			strBestilltAv & "," &_
			strSOBestilltAv & "," &_	
			strFaktsats & "," &_
			strFaktAntall & "," &_
			strFaktBeloep & "," &_
			strFaktSort & "," &_
			strBegDato & "," &_
			fakturaType & ")"
Response.Write "INSERT INTO VIKAR_UKELISTE 2:<br>"
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollBackTrans
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("En feil oppstod under oppretting av ukeliste for vikar.")
			call RenderErrorMessage()	
		End if
	End if 'strAntall <> "" Or strFaktAntall <> ""

	' Sett inn overtid for steg 2 i ukeliste_vikar
	strLoennsartnr = "163" 'Denne er hardkodet for overtid steg 2
	strAntall = FixNumericQueryString("PR100")
	strFaktAntall = FixNumericQueryString("PRFAKT100")

	If CDbl(strAntall) > 0 OR CDbl(strFaktAntall) > 0  Then

		'Hent rate for denne lønnsarten
		loennartSQL = "SELECT [LoennRate] FROM [H_LOENNSART] WHERE [LoennsartNr] = '" & strLoennsartnr & "'"
		Set rsLoennartInfo = GetFirehoseRS(loennartSQL, Conn)		
		strSats = strTimelonn * rsLoennartInfo("LoennRate")
		rsLoennartInfo.close
		set rsLoennartInfo = nothing

		If strAntall > 0 Then
			strBeloep = strSats * strAntall
		Else
			strAntall = 0
		End If

		strFaktsats = strFakturapris * Request("lstRateSteg2") 
		strFaktBeloep = strFaktsats * strFaktAntall
				
		fakturaTypeSQL = "SELECT [FakturaType] FROM [H_FAKTURA_TYPE] WHERE [FakturaSats] = '" & Replace(Request("lstRateSteg2"), ",", ".") & "'"
		set rsFakturaType = GetFirehoseRS(fakturaTypeSQL, Conn)
		fakturaType = rsFakturaType("FakturaType")
		rsFakturaType.close
		set rsFakturaType = nothing

		Call fjernKomma(strSats)
		Call fjernKomma(strAntall)
		Call fjernKomma(strBeloep)
		Call fjernKomma(strFaktsats)
		Call fjernKomma(strFaktAntall)
		Call fjernKomma(strFaktBeloep)

		strFaktSort = 3

		strSQL = "INSERT INTO VIKAR_UKELISTE " &_
			"(VikarID, OppdragID, FirmaID, UkeNr, Loennsartnr, Antall, Sats, Belop, Notat, Dato, StatusID, Overfort_loenn_status, " &_
			"Overfort_fakt_status, BestilltAv, SOBestilltav, Fakturapris, Fakturatimer, Fakturabeloep, FaktSort, FraDato, FakturaType) " &_
			"values (" &_
			strVikarID & "," &_
			strOppdragID & "," &_
			strFirmaID & "," &_
			strUkeNr & "," &_
			strLoennsartnr & "," &_
			strAntall & "," &_
			strSats & "," &_
			strBeloep & ",'" &_
			splittopt & "'," &_
			strEndDato & "," &_
			VikarTStatus & "," &_
			" 1, 1, " &_
			strBestilltAv & "," &_
			strSOBestilltAv & "," &_			
			strFaktsats & "," &_
			strAntall & "," &_
			strFaktBeloep & "," &_
			strFaktSort & "," &_
			strBegDato & "," &_
			fakturaType & ")"
Response.Write "INSERT INTO VIKAR_UKELISTE 3:<br><br>"
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollBackTrans
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("En feil oppstod under oppretting av ukeliste for vikar.")
			call RenderErrorMessage()	
		End if
		
	End IF 'strAntall = ""

End If    'nedgradering else vanlig behandling

'  VIKAR_LOEN_VARIABLE 

	' Check if selvstendig
	strSQL = "SELECT [TypeID] FROM [VIKAR] WHERE [vikarID] = " & strVikarID

	set rsSelv = GetFirehoseRS(strSQL, Conn)
	TypeID = rsSelv("TypeID")
	rsSelv.Close
	Set rsSelv = Nothing

If Not TypeID = 3 Then
	' SQL for AVDELING
	strSQL = "SELECT AvdelingID From OPPDRAG WHERE OppdragID = " & strOppdragID

	Set rsOppdrag = GetFirehoseRS(strSQL, Conn)

	If Not rsOppdrag.EOF Then
		strAvdeling = rsOppdrag("AvdelingID")
	Else
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing		
		AddErrorMessage("Finner ikke avdelingsID for oppdraget. Generering av timeliste er stanset. Kontakt administrator, og sjekk opplyninger for oppdraget.")
		call RenderErrorMessage()
	End If

	rsOppdrag.Close
	set rsOppdrag = nothing

	' SQL for displaying data
	strSQL = "SELECT Loennsartnr, Ant = SUM(Antall), Sats, Bel=Sum(Belop), VIKAR_UKELISTE.StatusID, Oppdrag.AvdelingID, Oppdrag.TomID,VIKAR_UKELISTE.oppdragid " &_
		"FROM VIKAR_UKELISTE, OPPDRAG " &_
		"WHERE VikarID = " & strVikarID &_
		" AND VIKAR_UKELISTE.OppdragID = OPPDRAG.OppdragID" &_
		" AND VIKAR_UKELISTE.StatusID = 5 " &_
		" AND Overfort_loenn_status < 3 " &_
		"GROUP BY Loennsartnr, Sats, VIKAR_UKELISTE.StatusID, OPPDRAG.AvdelingID, Oppdrag.TomID,VIKAR_UKELISTE.oppdragid"
Response.Write "Select VIKAR_UKELISTE:<br>"
	Set rsVikar = GetFirehoseRS(strSQL, Conn)

	' Deleting from VIKAR_LOEN_VARIABLE If exist rows on this id
	strSQL = "DELETE FROM VIKAR_LOEN_VARIABLE " &_
	      "WHERE VikarID = " & strVikarId &_
		" AND Overfor_loenn_status < 3"

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under sletting av vikar lønn variable.")
		call RenderErrorMessage()	
	End if

	' Insert into database VIKAR_LOEN_VARIABLE
	status = 1
	strLoennstakernr = strVikarID

	strLV = 150
	str50 = 160
	str100 = 163

	do while not rsVikar.EOF

		strAvdeling = rsVikar("AvdelingID")

		If TypeID = 1 Then

			strLArtNr = CStr(rsVikar("Loennsartnr"))

			SELECT Case strLArtNr
				Case 150
					strLoennsArtNr = strLV
					strLV = strLV + 1
				Case 160
					strLoennsArtNr = str50
					str50 = str50 + 1
				Case 163
					strLoennsArtNr = str100
					str100 = str100 + 1
				Case Else
					strLoennsArtNr = strLArtNr
			End Select

		End If 'TypeId = 1

		If TypeID = 2 Then
			strLoennsartnr = 50
		End If

		strAnt = rsvikar("Ant")
		strSats = rsvikar("Sats")
		strBelop = rsVikar("Bel")
		strTStatus = rsVikar("StatusID")
		strProjektNr = rsVikar("TomID")
		strMiddleOpp = rsVikar("OppdragID")

		Call fjernKomma(strAnt)
		Call fjernKomma(strSats)
		Call fjernKomma(strBelop)


		strSQL = "INSERT INTO VIKAR_LOEN_VARIABLE (VikarID, Avdeling, Prosjektnr, Loennstakernr, " &_
			"Loennsartnr, Dato, Antall, Sats, Beloep, FirmaID, "&_
			"OppdragID, Overfor_Loenn_Status, TimelisteStatus) " &_
			"values (" &_
			strVikarID & "," &_
			strAvdeling & "," &_
			strProjektNr & "," &_
			strLoennstakernr & ",'" &_
			strLoennsArtNr & "'," &_
			"NULL," &_
			strAnt & "," &_
			strSats & "," &_
			strBelop & "," &_
			strFirmaID & "," &_
			strMiddleOpp & "," &_
			status & "," &_
			strTStatus &  ")"
Response.Write "VIKAR_LOEN_VARIABLE:<br>"
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollBackTrans
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("En feil oppstod under oppretting av vikar lønn variable.")
			call RenderErrorMessage()	
		End if

		rsVikar.MoveNext
	loop
 	rsVikar.Close
	Set rsVikar = Nothing

End If 'selvstendig

'                         ******  FAKTURAGRUNNLAG *****
'korriger fra-/tildato for faktura

strSQL = "UPDATE VIKAR_UKELISTE Set "&_
	" FraDato = (SELECT MIN(dato) FROM DAGSLISTE_VIKAR "&_
			" WHERE vikarID = " & strVikarID &_
			" AND oppdragID = " & strOppdragID &_
			" AND TimelisteVikarStatus = 5 "&_
			" AND Fakturastatus = 1 "&_
			" AND Loennstatus = 1 )"&_
	", Dato = (SELECT MAX(dato) FROM DAGSLISTE_VIKAR "&_
			" WHERE vikarID = " & strVikarID &_
			" AND oppdragID = " & strOppdragID &_
			" AND TimelisteVikarStatus = 5 "&_
			" AND Fakturastatus = 1 "&_
			" AND Loennstatus = 1 )"&_
	" WHERE OppdragID = " & strOppdragID &_
	" AND VikarID = " & strVikarID &_
	" AND Overfort_fakt_status = 1 "&_
	" AND Overfort_loenn_status = 1 "

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under oppretting av ukeliste for vikar.")
		call RenderErrorMessage()	
	End if

	' parameters
	strVikarID 		= Request("VikarID")
	strOppdragID 	= Request("OppdragID")
	strFirmaID 		= Request("FirmaID")

	if(ErXisContact = true) then
		contactSQL = " Kontakt = " & strBestilltAv
	else
		contactSQL = " SOKontakt = " & strSOBestilltAv
	end if

    'save deleted record comments to later add
    			
	' Deleting from FAKTURAGRUNNLAG If exist rows on this id
	strSQL = "DELETE FROM FAKTURAGRUNNLAG " &_
			"WHERE " &_
			contactSQL &_
			" AND VikarID = " & strVikarID &_
			" AND Status = 1"

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under oppretting av vikar lønn variable.")
		call RenderErrorMessage()	
	End if

	' SQL for displaying data from VIKAR_UKELISTE

	'Prosesser tidligere dags/uke lister
	dim strFaktTekst 'Variabel for å teste lengde av streng til fakturalinje
	redim oppdrag(20) 'Inneholder tekst med oppdragsnummer

	if(ErXisContact = true) then
		contactSQL = " U.BestilltAv = " & strBestilltAv
	else
		contactSQL = " U.SOBestilltAv = " & strSOBestilltAv
	end if
	
	'Henter ut fakturagrunnlag
	strSQL = "SELECT U.VikarID, OppdragID, Ant = SUM(Fakturatimer), " &_
		"Fakturapris, fakturaType, LinjeSum = SUM(FakturaBeloep), " &_
		"BestilltAv, SOBestilltAv, Navn = (Fornavn +' ' + Etternavn), Dato, FraDato " &_
		"FROM VIKAR_UKELISTE U, VIKAR " &_
		" WHERE " &_
		contactSQL &_
		" AND U.VikarID = VIKAR.VikarID" &_
		" AND U.VikarID = " & strVikarID &_
		" AND Overfort_fakt_status < 3" &_
		" AND U.StatusId > 4" &_
		" AND U.FakturaType IS NOT NULL" &_		
		" GROUP BY fakturaType, Fakturapris, U.VikarID, U.OppdragID, BestilltAv, SOBestilltAv, ForNavn, Etternavn, Dato, Fradato " &_
		" ORDER BY U.VikarID desc, fakturaType desc"

		'" AND U.Fakturatimer > 0.0" &_
		'" AND U.FakturaBeloep > 0.0" &_

	Set rsFaktgr = GetFirehoseRS(strSQL, Conn)

' Insert into FAKTURAGRUNNLAG
If not rsFaktgr.EOF Then	
	dim IntLinjeNr		'Tellevariabel som brukes for å sette sortering på fakturalinjer = Linjenummer
	redim oppdrag(20) 	'Inneholder tekst med oppdragsnummer
	dim rsTom			'Holder tjenesteomraade for aktuelt oppdrag i loop
	dim IntAntOppdrag	'Tellevariabel som teller antall oppdrag som blir lagt inn på en fakturalinje. Tekst må være < 40 tegn
	dim IntOppdLinjer	'Holder orden paa alle oppdragene for linjene
	dim strOppdrag
	dim strAllOppdrag

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
		strSQL = "SELECT AvdelingID from OPPDRAG WHERE OppdragID = " & rsFaktgr("OppdragID")
		Set rsID = GetFirehoseRS(strSQL, Conn)

		If (rsID.EOF) Then
			Conn.RollBackTrans
			CloseConnection(Conn)
			set Conn = nothing		
			AddErrorMessage("Oppdraget har ikke tilknyttet noen avdeling!")
			set rsid = nothing
			rsFaktgr.close
			set rsFaktgr = nothing
			call RenderErrorMessage()
		end if

		VType  = rsID("AvdelingID")

		Avdeling = rsID("AvdelingID")
				
		Art = rsFaktgr("fakturaType")
		IntLinjeNr = IntLinjeNr + 1

		If ((VType = 1) AND (instr(strOppdrag, rsFaktgr("OppdragID")) = 0)) Then

			if(ErXisContact = true) then
				contactSQL = " U.BestilltAv = " & strBestilltAv
			else
				contactSQL = " U.SOBestilltAv = " & strSOBestilltAv
			end if
			
			strSQL = "SELECT U.VikarID, Ant = SUM(Fakturatimer), " &_
			"Fakturapris, fakturaType, LinjeSum = SUM(FakturaBeloep), " &_
			"BestilltAv, SOBestilltAv, Navn=(Fornavn +' ' + Etternavn), Dato, FraDato " &_
			 "FROM VIKAR_UKELISTE U, VIKAR " &_
			" WHERE " &_
			contactSQL &_
			" AND U.VikarID = VIKAR.VikarID" &_
			" AND U.VikarID = " & strVikarID &_
			" AND Overfort_fakt_status = 1" &_
			" AND U.oppdragid = " & rsFaktgr("OppdragID") &_
			" AND U.StatusId > 4" &_
			" AND U.FakturaType IS NOT NULL" &_
			" GROUP BY fakturaType, Fakturapris, U.VikarID, BestilltAv, SOBestilltAv, ForNavn, Etternavn, Dato, Fradato " & _
			"Order by U.VikarID desc, fakturaType desc"

'			" AND U.Fakturatimer > 0.0" &_
			'" AND U.FakturaBeloep > 0.0" &_			

			Set rsFaktgr2 = GetFirehoseRS(strSQL, Conn)

			'Dette er kurs - tjenesteomraade = 1
			strSQL = "SELECT DISTINCT o.tomid " & _
				" FROM VIKAR_UKELISTE U, VIKAR, oppdrag O " & _
				" WHERE O.oppdragid = U.oppdragid " & _
				" AND" & contactSQL &_
				" AND U.VikarID = VIKAR.VikarID" &_
				" AND U.VikarID = " & strVikarID &_
				" AND U.oppdragid = " & rsFaktgr("OppdragID") &_
				" AND Overfort_fakt_status = 1" &_
				" AND U.StatusId > 4"

			Set rsTom = GetFirehoseRS(strSQL, Conn)

			if rsTom.recordcount > 1 then			
				set rsTom = nothing
				set rsFaktgr2 = nothing
				Conn.RollBackTrans
				CloseConnection(Conn)
				set Conn = nothing		
				AddErrorMessage("Ugyldig tjenesteområde for kurs!<br>Kunne ikke generere faktura!")
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
				if (FRate = 1.0) then 'Vanlige timer
					strTekst = rsFaktgr("Navn") & " - " & rsFakturaSats("Fakturabeskrivelse")
					
					'for bonus updates add the comment entered by resposible person
					if (rsFakturaSats("Fakturabeskrivelse")="Korreksjon") then
					   
						strTekst = rsFaktgr("Navn") & " - " & strNote
						
					end if 
				
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
				"values (" &_
				rsFaktgr("VikarID") & ", " &_
				rsFaktgr("OppdragID") & ", " &_
				strFirmaID & ", "  &_
				strBestilltAv & ", " &_
				strSOBestilltAv & ",'" &_
				strArtNr & "','" &_
				strTekst & "'," &_
				strAntall & "," &_
				strFaktgr & "," &_
				strLinjeSum & ", 1 , " &_
				IntLinjeNr & "," &_
				Avdeling & ")"

Response.Write "INSERT INTO FAKTURAGRUNNLAG:<br>"
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

				strSQL = "SELECT [AvdelingID] FROM [OPPDRAG] WHERE [OppdragID] = " & rsFaktgr("OppdragID")
				Set rsID = GetFirehoseRS(strSQL, Conn)
				Avdeling = rsID("AvdelingID")
				rsID.Close
				Set rsID = Nothing

			loop
			strOppdrag = strOppdrag & "," & rsFaktgr("OppdragID")
		end if
		'slutt kursvikar
		If VType > 1 Then
		
			'Dersom før 1.7.2001 artikkelnummer -> vikarid ellers artikkelnummer -> tjenesteomradeid
			strSQL ="SELECT tomid FROM oppdrag WHERE oppdragid=" & rsFaktgr("OppdragID")
			Set rsTom = GetFirehoseRS(strSQL, Conn)
			'Hardcoded to "andre" for the following toms: 7,9,10,11,12,13,14,15 and 16
			if (clng(rsTom("tomid")) = 7 OR clng(rsTom("tomid")) = 9 OR clng(rsTom("tomid")) = 10 OR clng(rsTom("tomid")) = 11 OR clng(rsTom("tomid")) = 12 OR clng(rsTom("tomid")) = 13 OR clng(rsTom("tomid")) = 14 OR clng(rsTom("tomid")) = 15 OR clng(rsTom("tomid")) = 16) then
				strArtNr = "301008" 
			else
				strArtNr = cstr( ( 301000 + clng(rsTom("tomid")) ) )
			end if
			rsTom.close
			set rsTom = nothing
					
			strSQL = "SELECT [Fakturabeskrivelse], [FakturaSats] FROM [H_FAKTURA_TYPE] WHERE [FakturaType] = " & Art
Response.Write "rsFakturaSats:<br>"
			set rsFakturaSats = GetFirehoseRS(strSQL, Conn)
			
			FRate = rsFakturaSats("FakturaSats")
			if (FRate = 1.0) then
				strTekst = rsFaktgr("Navn") & " - " & rsFakturaSats("Fakturabeskrivelse")
				
				'for bonus updates add the comment entered by resposible person
				if (rsFakturaSats("Fakturabeskrivelse")="Korreksjon") then
				   
					strTekst = rsFaktgr("Navn") & " - " & strNote
					
				end if 
				
			elseif (FRate > 1.0) then
				strTekst = rsFaktgr("Navn") & " - " & rsFakturaSats("Fakturabeskrivelse")
				strArtNr = strArtNr & "-" & cstr(((CDbl(FRate) - 1) * 100))
			end if
			'this is test
			'strTekst = "Test2 - bonus!!!"
			
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
			strBestilltAv & ", " &_
			strSOBestilltAv & ",'" &_
			strArtNr & "','" &_
			strTekst & "'," &_
			strAntall & "," &_
			strFaktgr & "," &_
			strLinjeSum & ", 1 , " &_
			IntLinjeNr & "," &_
			Avdeling & ")"
			
Response.Write "INSERT INTO FAKTURAGRUNNLAG:<br><br>"
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
			strFaktTekst = oppdrag(IntOppdLinjer) & "," & rsFaktgr("OppdragID")

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
		strAllOppdrag = strAllOppdrag & "," & rsFaktgr("OppdragID")
		rsFaktgr.MoveNext
	loop

	IntLinjeNr = IntLinjeNr + 1

	strSQL = "SELECT strTekst = (CASE tom.rub_Mvakode WHEN '00' THEN 'Opplæring ' WHEN '01' THEN 'Vikar ' ELSE 'Opplæring ' END), AvdelingID " & _
	"FROM tjenesteomrade tom, oppdrag o " & _
	"WHERE tom.tomid = o.tomid " & _
	"AND oppdragid = " & OID

	Set rsID = GetFirehoseRS(strSQL, Conn)

	strTekst = rsID("strTekst") & fraDato & " - " & tilDato
	
	Avdeling = rsID("AvdelingID")

	rsID.Close
	Set rsID = Nothing

	strSQL = "INSERT INTO FAKTURAGRUNNLAG " &_
	"(VikarID, OppdragID, FirmaID, Kontakt, SOKontakt, Tekst, Status, LinjeNr, Avdeling) " &_
	"VALUES (" &_
		VID & ", " &_
		OID & ", " &_
		strFirmaID & ", "  &_
		strBestilltAv & "," &_
		strSOBestilltAv & ",'" &_
		strTekst & "'," &_
		" 1 , " &_
		IntLinjeNr & "," &_
		Avdeling & ")"

'Response.Write "INSERT INTO FAKTURAGRUNNLAG 3:" & strSQL & "<br><br>"
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
			strTekst = "Oppdrag  " & oppdrag(i)
		Else
			strTekst = oppdrag(i)
		End If
		
		strSQL = "INSERT INTO FAKTURAGRUNNLAG " &_
		"(VikarID, OppdragID, FirmaID, Kontakt, SOKontakt, Tekst,  Status, LinjeNr, Avdeling) " &_
		"values (" &_
		VID & ", " &_
		OID & ", " &_
		strFirmaID & ", "  &_
		strBestilltAv & "," &_
		strSOBestilltAv & ",'" &_
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
End If 'det finnes linjer med fakt status 1 i ukelista

	'Response.End
	Conn.CommitTrans
	'Conn.RollBackTrans
	CloseConnection(Conn)
	set Conn = nothing
session("debug") = false
	'  videresending
	Redir = "Vikar_timeliste_vis3.asp?VikarID=" & strVikarID & "&OppdragID=" & strOppdragID & "&FirmaID=" & strFirmaID & "&startdato=" & strStartdato & "&splittopt=" & session("splittopt") & "&splittuke=" & session("splittuke") &"&splittopt2=."& session("splittopt")
'Response.End
	Response.Redirect Redir
%>