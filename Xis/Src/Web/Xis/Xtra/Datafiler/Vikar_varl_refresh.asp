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

	dim strProjektNr
	dim Conn
	dim rsOppdrag
	dim strVikarID
	dim strOppdragID
	dim strFirmaID
	dim TypeID
	dim strSQL

	Sub fjernKomma(strString )
		pos = Instr(strString, ",")
		If pos <> 0 Then
			mellom = Left(strString, pos-1) & "." & Left(Mid(strString, pos + 1), 2)
 			strString = mellom 
		End If
	End Sub

	'Oppdatere lønnsgrunnlaget for vikaren på bakgrunn av ukelistene (aggregerte timelister)

	' Connect to database
	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))	

	' parametere
	strVikarID = Request.QueryString("VikarID")
	strOppdragID = Request.QueryString("OppdragID")
	strFirmaID = Request.QueryString("FirmaID")

	' Sjekk vikartype, 1 = ordinær vikar, 2 selvstendig næringsdrivende, 3 = aksjeselskap (se tabell H_VIKAR_TYPE)
	strSQL = "SELECT [TypeID] FROM [VIKAR] WHERE [vikarID] = " & strVikarID

	set rsSelv = GetFirehoseRS(strSQL, Conn)
	TypeID = rsSelv("TypeID")
	rsSelv.Close
	Set rsSelv = Nothing

	'Start transaction
	Conn.Begintrans

	If Not TypeID = 3 Then 

		' SQL for AVDELING
		strSQL = "SELECT [AvdelingID], [TomID] FROM [OPPDRAG] WHERE [OppdragID] = " & strOppdragID

		Set rsOppdrag = GetFirehoseRS(strSQL, Conn)

		If HasRows(rsOppdrag) Then
			strAvdeling = rsOppdrag("AvdelingID")
			strProjektNr = rsOppdrag("TomID")
			rsOppdrag.Close
		Else
			strAvdeling = "NULL"
		End If
		set rsOppdrag = nothing

		' SQL for displaying data
		strSQL = "SELECT [Loennsartnr], [Ant]=SUM([Antall]), [Sats], [Bel]=SUM([Belop]), [StatusID] " &_ 
			"FROM [VIKAR_UKELISTE] " &_
			"WHERE [VikarID] = " & strVikarID &_
			" AND [StatusID] = 5 " &_
			" AND [Overfort_loenn_status] < 3 " &_
			"GROUP BY [Loennsartnr], [Sats], [StatusID] "

		Set rsVikar = GetFirehoseRS(strSQL, Conn)

		' Display data
		' sletting av eksisterende lønnsgrunnlag for denne vikaren.
		strSQL = "DELETE FROM [VIKAR_LOEN_VARIABLE] " &_
				"WHERE [VikarID] = " & strVikarId &_
				" AND [Overfor_loenn_status] < 3"

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
		strFirmaID = session("FirmaID")

		strLV = 150 'grunn lønnsarter, ordinær lønn
		str50 = 160 'grunn lønnsarter, overtid steg 1
		str100 = 163 'grunn lønnsarter, overtid steg 2

		do while not rsVikar.EOF

			If TypeID = 1 Then 'vanlig vikar

				strLArtNr = CStr(rsVikar("Loennsartnr"))

				SELECT Case strLArtNr
					Case 150 'ordinær lønn
						strLoennsArtNr = strLV
						strLV = strLV + 1 		
					Case 160 'overtid steg 1
						strLoennsArtNr = str50
						str50 = str50 + 1
					Case 163 'overtid steg 2
						strLoennsArtNr = str100
						str100 = str100 + 1
					Case Else
						strLoennsArtNr = strLArtNr
				End Select

			End If 'TypeId = 1

			If TypeID = 2 Then 'selvstendig næringsdrivende
				strLoennsartnr = 50 'lønnsart for selvstendig næringsdrivende
			End If

			strAnt = rsvikar("Ant")
			strSats = rsvikar("Sats")
			strBelop = rsVikar("Bel")
			strTStatus = rsVikar("StatusID")

			Call fjernKomma(strAnt)
			Call fjernKomma(strSats)
			Call fjernKomma(strBelop)

			strSQL = "INSERT INTO VIKAR_LOEN_VARIABLE (VikarID, Avdeling, Prosjektnr, Loennstakernr, " &_
				"Loennsartnr, Dato, Antall, Sats, Beloep, FirmaID, " &_
				"OppdragID, Overfor_Loenn_Status, TimelisteStatus) " &_
				"VALUES(" &_
				strVikarID & "," &_
				strAvdeling & "," &_
				strProjektNr & "," &_
				strLoennstakernr & ",'" &_
				strLoennsArtNr & "'," &_
				"NULL, " &_
				strAnt & "," &_
				strSats & "," &_
				strBelop & "," &_
				strFirmaID & "," &_
				strOppdragID & ", " &_
				status & "," &_
				strTStatus &  ")"

			If ExecuteCRUDSQL(strSQL, Conn) = false then
				Conn.RollBackTrans
				CloseConnection(Conn)
				set Conn = nothing
				AddErrorMessage("En feil oppstod under oppretting av vikar lønn variable.")
				call RenderErrorMessage()	
			End if
			rsVikar.MoveNext
		loop 
		rsVikar.Close: Set rsVikar = Nothing
	End If 'selvstendig
	
	Conn.CommitTrans
	'Conn.RollBackTrans
	CloseConnection(Conn)
	set Conn = nothing	

'response.End
	redir = "Vikar_varl_vis3.asp?VikarID=" & strVikarID & "&OppdragID=" & strOppdragID & "&FirmaID=" & strFirmaID
	Response.redirect redir
%>