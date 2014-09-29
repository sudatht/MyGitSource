<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<%

	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim Conn
	dim rsAnt
	dim strSQL
	dim strEndre
	dim strLinje
	dim strLinjeNr
	dim strOppdragID
	dim strKontaktID
	dim SOKontaktID
	dim strFirmaID
	dim strAvdeling
	dim strVikarID
	dim strArtikkelNr
	dim strTekst
	dim strAntall
	dim strEnhetspris
	dim strsplit

	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))	

	' Check parameters AND put into variables
	strEndre = lcase(Request("Endre"))
	strLinje = Request("Linje")
	strOppdragID = Request("OppdragID")
	strKontaktID = Request("Kontakt")
	SoKontaktID = Request("SOKontakt")
	strFirmaID = Request("FirmaID")
	strAvdeling = Request("Avdeling")
	strVikarID = Request("VikarID")
	strLinjeNr = Request("LinjeNr")
	strArtikkelNr = Request("Artikkelnr") 
	strTekst = Request("Tekst") 

	If Request("Antall") = "" Then strAntall = "NULL" Else strAntall =  Request("Antall")
	If Request("Enhetspris") = "" Then strEnhetspris = "NULL" Else strEnhetspris = Request("Enhetspris")

	strsplit = Request("split")
	If strsplit = "" Then strsplit = "NULL"

	strFakturanr = Request("Fakturanr")
	If strfakturanr = "" Then strFakturanr =" NULL"

	strFakturadato = Request("Fakturadato")
	If strFakturadato = "" Then
		strFakturadato = "NULL"
	Else
		strFakturadato = dbDate(strFakturadato)
	End If

	strStatus = Request("Status")

	'se strEndre & "  e<br>"
	'se strLinje & "  ln1<br>"
	'se strVikarID & "  v<br>"
	'se strLinjeNr & "  ln<br>"
	'se strFirmaID & "  f<br>"
	'se strKontaktID & "  k<br>"
	'se strOppdragID & "  o<br>"
	'se strArtikkelNr & "  art<br>"
	'se strTekst & "  txt<br>"
	'se strAntall & "  ant<br>"
	'se strEnhetspris & "  enhet<br>"
	'se strSplit & " <br>"
	'se strFakturanr & " <br>"
	'se strFakturadato & " <br>"
	'se strStatus & " <br>"

	'Start start transaction som that we can rollback if something goes wrong..
	Conn.Begintrans
	' Delete row in FAKTURALINJER
	If strEndre ="slett" Then

		strSQL = "DELETE FROM FAKTURAGRUNNLAG WHERE FakturalinjeID = "  & strLinje
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollBackTrans
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Feil oppstod under sletting av fakturagrunnlag.")
			call RenderErrorMessage()
		End if

	End if 'deleting

	If strTekst <> ""  Then  'minimum required fields

		' UPDATE row in FAKTURAGRUNNLAG
		If strEndre = "endre" Then

			If strAntall <> "NULL" AND strEnhetspris <> "NULL" Then
				Call SETtKomma(strEnhetspris): Call SETtKomma(strAntall)
				sumsum = strEnhetspris * strAntall
				Call fjernKomma(sumsum)
			Else
				sumsum = "NULL"
			End IF

			strSQL = "UPDATE FAKTURAGRUNNLAG SET" &_
				" FirmaID = " & strFirmaID &_
				", OppdragID = " & strOppdragID &_
				", VikarID = " & strVikarID &_
				", Antall = " & strAntall &_	
				", Linjesum = " & sumsum &_	
				", Status = " & strStatus &_	
				", Artikkelnr = '" & strArtikkelnr & "'" &_	
				", Tekst = '" & strTekst & "'" &_	
				", Enhetspris = " & strEnhetspris &_	
				", Fakturanr = " & strFakturanr &_	
				", Fakturadato = " & strFakturadato &_	
				", split = " & strSplit &_	
				" WHERE FakturalinjeID = " & strLinje


			If ExecuteCRUDSQL(strSQL, Conn) = false then
					Conn.RollBackTrans
					CloseConnection(Conn)
					set Conn = nothing
					AddErrorMessage("En feil oppstod under oppdatering av fakturagrunnlag.")
					call RenderErrorMessage()
			End if

		' Insert sort  into FAKTURAGRUNNLAG
		ElseIf strEndre = "insert" Then

			strLinjeNr = strLinjeNr + 1

			if (len(strKontaktID) > 0) then
				kontaktSQL = " Kontakt = " & strKontaktID
			else
				kontaktSQL = " SOKontakt = " & SOKontaktID
			end if

			strSQL = "SELECT a = COUNT(LinjeNr)" &_
				"FROM FAKTURAGRUNNLAG " &_
				"WHERE" &_
				kontaktSQL &_
				" AND OppdragID = " & strOppdragID &_
				" AND Status = " & strStatus &_
				" AND VikarID = " & strVikarID &_
				" AND LinjeNr >= " & strLinjeNr 

			Set rsAnt = GetFirehoseRS(strSQL, Conn)

			antRest = rsAnt("a").value
			rsAnt.Close
			set rsAnt = nothing

			if (len(strKontaktID) > 0) then
				kontaktSQL = " Kontakt = " & strKontaktID
			else
				kontaktSQL = " SOKontakt = " & SOKontaktID
			end if

			For i = (strLinjeNr + antRest) To  strLinjeNr Step -1
			
				strSQL = "UPDATE FAKTURAGRUNNLAG SET " &_
					"LinjeNr = " & (i + 1) &_
					" WHERE " & _
					kontaktSQL & _
					" AND OppdragID = " & strOppdragID &_
					" AND Status = " & strStatus &_
					" AND VikarID = " & strVikarID &_
					" AND LinjeNr = " & i

					If ExecuteCRUDSQL(strSQL, Conn) = false then
						Conn.RollBackTrans
						CloseConnection(Conn)
						set Conn = nothing
						AddErrorMessage("En feil oppstod under oppdatering av fakturagrunnlag.")
						call RenderErrorMessage()	
					End if
			Next

			If strAntall <> "NULL" AND strEnhetspris <> "NULL" Then
				sumsum = strEnhetspris * strAntall
				Call fjernKomma(sumsum)
			Else
				sumsum = "NULL"
			End IF
			Call fjernKomma(strEnhetspris)
			Call fjernKomma(strAntall)

			if (len(strKontaktID) > 0) then
				strKontaktIDInsert = strKontaktID
				SOKontaktIDInsert = "NULL"
			else
				SOKontaktIDInsert = SOKontaktID
				strKontaktIDInsert = "NULL"
			end if

			strSQL = "INSERT INTO FAKTURAGRUNNLAG (FirmaID, OppdragID, VikarID, Antall, Linjesum, Status, Artikkelnr, Tekst, " &_
				"Enhetspris, Kontakt, SOKontakt, LinjeNr, NyLinje, Fakturanr, Fakturadato, split, avdeling) " &_
				"values (" &_
				strFirmaID &_
				", " & strOppdragID &_
				", " & strVikarID &_
				", " & strAntall &_	
				", " & sumsum &_	
				", " & strStatus &_	
				", '" & strArtikkelnr &_	
				"','" & strTekst &_	
				"', " & strEnhetspris &_
				", " & strKontaktIDInsert &_
				", " & SOKontaktIDInsert &_
				", " & strLinjeNr	&_
				",1, " & strFakturanr &_
				", " & strFakturadato	&_
				", " & strSplit &_
				", " & strAvdeling &_
				")"
				
			 
			If ExecuteCRUDSQL(strSQL, Conn) = false then
				Conn.RollBackTrans
				CloseConnection(Conn)
				set Conn = nothing
				AddErrorMessage("En feil oppstod under oppretting av fakturagrunnlag.")
				call RenderErrorMessage()
			End if
				
		End If 'UPDATE or insert 
	End If 'filds have content 

	Conn.CommitTrans
	CloseConnection(Conn)
	set Conn = nothing

	Response.Redirect "Faktura_Vis.asp?OppdragID=" & strOppdragID & "&vikarid=" & strVikarID & "&FirmaID=" & strFirmaID & "&SOKontakt=" & SOKontaktID & "&Kontakt=" & strKontaktID & "&avdeling=" & strAvdeling 
%>