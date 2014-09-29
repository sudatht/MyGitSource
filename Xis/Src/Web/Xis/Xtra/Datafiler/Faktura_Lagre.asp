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
'session("debug") = true
	dim Conn
	dim strSQL
	dim strKontaktID
	dim SOKontaktID
	dim kontaktSQL
	dim strOppdragID
	dim strFirmaID
	dim strLoennsartNr
	dim graderingskode
	dim strFaktDato
	dim frakode
	dim VikarerHer
	dim strAvdeling

	'Funksjonsdefinsjoner
	Function NextFaktNr() 'Henter neste ledige fakturanummer (internt i internsyst)
		Dim strNr
		strSQL = "SELECT nr = MAX(Fakturanr) FROM FAKTURAGRUNNLAG"
		Set rsNr = GetFirehoseRS(strSQL, Conn)

		If IsNull(rsNr("nr")) Then
			strNr = 1
		Else
			strNr = rsNr("nr") + 1
		End If
		rsNr.Close
		set rsNr = nothing
		NextFaktNr = strNr
	End Function

	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))	

	' Check parameters AND put into variables
	strOppdragID = Request("OppdragID")
	strKontaktID = Request("Kontakt")
	SOKontaktID = Request("SOKontakt")
	strFirmaID = Request("FirmaID")
	strLoennsartNr = Request("LoennsartNr")
	graderingskode = lcase(Request("graderingskode"))
	strFaktDato = Request("FakturaDato")
	frakode = Request("frakode")
	VikarerHer = Request("VikarerHer")
	strAvdeling = Request("Avdeling")

	If graderingskode = "nedgrad" Then  'kalles fra Vikar_timeliste_fakt_list.asp 
		ddd = "NULL"
		sss = 1

		' Start transaction
		Conn.Begintrans
		
		if (len(strKontaktID) > 0) then
			kontaktSQL = " BestilltAv = " & strKontaktID
		else
			kontaktSQL = " SOBestilltAv = " & SOKontaktID
		end if		
			
		' UPDATE status in DAGSLISTE_VIKAR (timelisten)
		strSQL = "UPDATE DAGSLISTE_VIKAR set " &_
			"Fakturastatus = " & sss &_
			", Fakturadato = " & ddd &_
			" WHERE " &_
			kontaktSQL &_
			" AND Fakturastatus = 2"

			If ExecuteCRUDSQL(strSQL, Conn) = false then
				Conn.RollBackTrans
				CloseConnection(Conn)
				set Conn = nothing
				AddErrorMessage("En feil oppstod under oppdatering av dagsliste for vikar.")
				call RenderErrorMessage()	
			End if

		' UPDATE status in FAKTURALINJER
		if (len(strKontaktID) > 0) then
			kontaktSQL = " Kontakt = " & strKontaktID
		else
			kontaktSQL = " SOKontakt = " & SOKontaktID
		end if

		strSQL = "UPDATE FAKTURAGRUNNLAG SET " &_
			"Status = " & sss &_
			", Fakturadato = " & ddd &_
			" WHERE " &_
			kontaktSQL &_
			" AND Status = 2"

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollBackTrans
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("En feil oppstod under oppdatering av fakturagrunnlag.")
			call RenderErrorMessage()	
		End if

		if (len(strKontaktID) > 0) then
			kontaktSQL = " BestilltAv = " & strKontaktID
		else
			kontaktSQL = " SOBestilltAv = " & SOKontaktID
		end if

		' UPDATE status in VIKAR_UKELISTE
		strSQL = "UPDATE VIKAR_UKELISTE SET " &_
			"Overfort_fakt_status = " & sss &_
			", Fakturadato = " & ddd &_
			" WHERE " &_
			kontaktSQL &_
			" AND Overfort_fakt_status = 2"

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollBackTrans
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("En feil oppstod under oppdatering av Vikar ukeliste.")
			call RenderErrorMessage()	
		End if

		if (len(strKontaktID) > 0) then
			kontaktSQL = " AND Kontakt = " & strKontaktID
		else
			kontaktSQL = " AND SOKontakt = " & SOKontaktID
		end if

		' DELETE FROM EKSPORT_RUB_ORDRE
		strSQL = "DELETE FROM EKSPORT_RUB_ORDRE " &_
				"WHERE Status = 1" &_
				kontaktSQL

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollBackTrans
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("En feil oppstod under sletting av faktura hode eksport grunnlag.")
			call RenderErrorMessage()	
		End if

		' DELETE FROM EKSPORT_RUB_ORDRELINJE
		strSQL = "DELETE FROM EKSPORT_RUB_ORDRELINJE " &_
				"WHERE Status = 1 " &_
				kontaktSQL
				
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollBackTrans
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("En feil oppstod under sletting av fakturalinje eksport grunnlag.")
			call RenderErrorMessage()	
		End if

		Conn.CommitTrans
		CloseConnection(Conn)
		set Conn = nothing

		' Push another page
		Response.Redirect "faktura_list.asp?velgAvdeling=" & request("velgAvdeling") & "&viskode=" & request("viskode") & "&dato1=" & dato1

	Else  'Ikke nedgradering  (kalles fra Faktura_vis.asp ) 

		strDato = dbDate(Date)

		' Start transaction
		Conn.Begintrans

		' Find new fakturanr
		If Request("fakturanr") = "" Then
			nr = NextFaktNr()
		Else
			nr = Request("fakturanr")
		End If

		If Request("splitt") = "Ja" Then
			Ersplittet = 1
		Else
			Ersplittet = 0
		End IF

		if (len(strKontaktID) > 0) then
			kontaktSQL = " AND Kontakt = " & strKontaktID
		else
			kontaktSQL = " AND SOKontakt = " & SOKontaktID
		end if

		' UPDATE status in FAKTURALINJER
		strSQL = "SELECT FakturaLinjeID, VikarID " &_
				"FROM FAKTURAGRUNNLAG " &_
				" WHERE 1 = 1 " &_
				kontaktSQL &_
				" AND FirmaID = " & strFirmaID &_
				" AND VikarID IN (" & VikarerHer & ")" &_
				" AND Status < 3" &_
				" ORDER BY VikarID, LinjeNr"

		Set rsAnt = GetFirehoseRS(strSQL, Conn)

		If hasRows(rsAnt) Then
			VID = rsAnt("VikarID")
			lNr = 0
			do while Not rsAnt.EOF

				lNr = lNr + 1

				'Nytt fakturanummer dersom fakturaen er splittet. Må hente nytt fakturanummer i tilfelle
				'det har blitt godkjent fakturaer før denne ble splittet
				If VID <> rsAnt("VikarID") AND Request("splitt")="Ja" Then
					'nr = nr + 1 
					nr = NextFaktNr()
					VID = rsAnt("VikarID")
					lNr = 1
				End If

				strSQL = "UPDATE FAKTURAGRUNNLAG SET " &_
					" Status = 2" &_
					", Fakturanr = " & nr &_
					", LinjeNr = " & lNr &_
					", split = " & Ersplittet &_
					" WHERE FakturaLinjeID = " & rsAnt("FakturaLInjeID")

				If ExecuteCRUDSQL(strSQL, Conn) = false then
					Conn.RollBackTrans
					CloseConnection(Conn)
					set Conn = nothing
					AddErrorMessage("En feil oppstod under oppdatering av fakturagrunnlag.")
					call RenderErrorMessage()	
				End if

				rsAnt.MoveNext
			loop
			rsAnt.Close
		End If  'no rows
		Set rsAnt = Nothing

		if (len(strKontaktID) > 0) then
			kontaktSQL = " AND BestilltAv = " & strKontaktID
		else
			kontaktSQL = " AND SOBestilltAv = " & SOKontaktID
		end if				
		
		' UPDATE fakturastatus in DAGSLISTE_VIKAR (timelisten)
		strSQL = "UPDATE DAGSLISTE_VIKAR SET " &_
			"Fakturastatus = 2" &_ 
			" WHERE Fakturastatus = 1 " &_
			" AND FirmaID = " & strFirmaID &_
			kontaktSQL &_
			" AND TimelisteVikarStatus = 5" &_
			" AND VikarID IN (" & VikarerHer & ")" &_
			" AND Dato <= " & dbDate(session("limitDato"))

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollBackTrans
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("En feil oppstod under oppdatering av dagsliste for vikar.")
			call RenderErrorMessage()	
		End if

		' UPDATE status in VIKAR_UKELISTE
		strSQL = "UPDATE VIKAR_UKELISTE SET" &_
			" Overfort_fakt_status = 2" &_
			" WHERE FirmaID = " & strFirmaID &_
			kontaktSQL &_
			" AND VikarID IN (" & VikarerHer & ")" &_
			" AND StatusID = 5" &_
			" AND Overfort_fakt_status = 1 "

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollBackTrans
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("En feil oppstod under oppdatering av vikar ukeliste.")
			call RenderErrorMessage()	
		End if
		
		Conn.CommitTrans		
	End If 'kode
	CloseConnection(Conn)
	set Conn = nothing

	redir = "Faktura_vis.asp?OppdragID=" & strOppdragID &_
	"&FirmaID=" & strFirmaID & "&Kontakt=" & strKontaktID & "&SOKontakt=" & SOKontaktID &_
	"&Avdeling=" & strAvdeling

	Response.Redirect  redir 
%>