<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.utils.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_ADMIN, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	Dim Avdeling 'avdelingsID
	Dim strWhereAvd 'Where-clause for utvalg på avdeling
	dim Conn
	dim strSQL

	' Modifications
	' Date - Who - What
	' 23012000 - Arne leithe - Used GetAvdeling to retrieve the correct dept number

	Function OnlyDigits( strString )
	' Remove all non-nummeric signs from string
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

	Function LeftFill(aStr, aLen)
		If IsNull(aStr) Then
			aStr = ""
		End If
		ss = Trim(aStr)
		If Len(aStr) = 0 Then
			ss = ""
		Else
			For i=Len(ss) To aLen - 1
				ss = "0" & ss
			Next
		End If
		LeftFill = ss
	End Function

	' Connect to database
	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))	

	' Get data
	Klientnr = 1
	'ArtikkelNummer =	(Artikkelnr)
	'ArtikkelNavn =		(Tekst)
	MVAKode = 0
	'IOrdre = 	(Antall)
	'Levert =	(Antall)
	'Utpris1 = 	(Enhetspris)

	if Request.QueryString("avd") <> "" then
		valgt_avd = CInt(Request.QueryString("avd"))
	else
		valgt_avd = 0
	end if

	'Lag where-clause for avdeling avh. av om det er valgt en avdeling eller ikke
	if valgt_avd > 0 then
		strWhereAvd = " and avdeling = " & valgt_Avd & " "
	else
		strWhereAvd = ""
	end if

	'	Hent id-er fra FAKTURALINJER
	strSQL = "SELECT [Artikkelnr], [Tekst], [Antall], [Enhetspris], [Fakturanr], [Buntnr], [Kontakt], [SOKontakt], [Avdeling] " &_
		"FROM FAKTURAGRUNNLAG " &_
		"WHERE [Status] = 2 " & strWhereAvd &_
		" ORDER BY [Fakturanr], [Linjenr]"

	set rsID = GetFirehoseRS(strSQL, Conn)
	if (HasRows(rsID)) then
		Conn.Begintrans
		fnr = rsID("Fakturanr")
		strSQL = "DELETE FROM [EKSPORT_RUB_ORDRELINJE] WHERE [Fakturanr] = " & rsID("Fakturanr")
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollBackTrans
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("En feil oppstod under oppretting av rubikon ordre linjer.")
			call RenderErrorMessage()
		End if

		DO WHILE NOT rsID.EOF
			If fnr <> rsID("Fakturanr") Then
				strSQL = "DELETE FROM [EKSPORT_RUB_ORDRELINJE] WHERE [Fakturanr] = " & rsID("Fakturanr")
				If ExecuteCRUDSQL(strSQL, Conn) = false then
					Conn.RollBackTrans
					CloseConnection(Conn)
					set Conn = nothing
					AddErrorMessage("En feil oppstod under sletting av gammelt datagrunnlag fra rubikon ordre linjer.")
					call RenderErrorMessage()
				End if
				fnr = rsID("Fakturanr")
			End If

			Artikkelnummer = LeftFill(rsID("Artikkelnr"), 6)
			Artikkelnavn = rsID("Tekst")
			If IsNull(rsID("Antall")) Then
				IOrdre = "NULL"
				Levert = "NULL"
			Else
				IOrdre = rsID("Antall")
				Call fjernKomma(IOrdre)
	 			Levert = IOrdre
			End If
			
			If IsNull(rsID("Enhetspris")) Then
				UtPris1 = "NULL"
			Else
				UtPris1 = rsID("Enhetspris")
				Call fjernKomma(UtPris1)
			End If

			Fakturanr = rsID("Fakturanr")
			Buntnr = rsID("Buntnr")

			Kontakt = rsID("Kontakt")
			if (isNULL(rsID("Kontakt"))) then
				Kontakt = "NULL"
			else
				Kontakt = rsID("Kontakt")
			end if
			
			if (isNULL(rsID("SOKontakt"))) then
				SOKontakt = "NULL"
			else
				SOKontakt = rsID("SOKontakt")
			end if	
			Hovedgruppe = 0
			Avdeling = rsID("avdeling")


			'	INSERT INTO EKSPORT_RUB_ORDRELINJE
			strSQL = "INSERT INTO [EKSPORT_RUB_ORDRELINJE] (" &_
				"[Fakturanr]," &_
				"[Formatnummer]," &_
				"[Klientnummer], " &_
				"[Registernummer], " &_
				"[Status], " &_
				"[Artikkelnummer], " &_
				"[Artikkelnavn], " &_
				"[MVAKode]," &_
				"[IOrdre], " &_
				"[Levert], " &_
				"[Utpris1], " &_
				"[Eksportert_Dato], " &_
				"[Kontakt], " &_
				"[SOKontakt], " &_
 				"[Buntnr], " &_
 				"[Hovedgruppe], " &_
 				"[Avdeling]" &_
				") VALUES (" &_
				rsID("Fakturanr") &_
				", 6001, 1 , 6, 1,'" &_
				Artikkelnummer & "','" &_
				Left(Artikkelnavn, 40) & "'," &_
				MVAKode & "," &_
				IOrdre & "," &_
				Levert & "," &_
				UtPris1 & "," &_
				"NULL," &_
				Kontakt & "," &_
				SOKontakt & "," &_
				Buntnr & "," &_
 				Hovedgruppe & "," &_
				Avdeling & ")"

			If ExecuteCRUDSQL(strSQL, Conn) = false then
				Conn.RollBackTrans
				CloseConnection(Conn)
				set Conn = nothing
				AddErrorMessage("En feil oppstod under sletting av gammelt datagrunnlag fra rubikon ordre linjer.")
				call RenderErrorMessage()
			End if
			rsID.MoveNext
			
		loop
		Conn.CommitTrans
		rsID.Close		
	end if
	Set rsID = Nothing
	CloseConnection(Conn)

	'Response.redirect "Eksp_RUB_01.asp?avd=" & valgt_avd
	Response.redirect "../WebUI/InvoiceHandling/VismaFileCreater.aspx?avd=" & valgt_avd & "&dato1=" & session("limitdato")
%>