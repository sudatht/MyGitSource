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

	dim Conn
	dim strSQL
	dim strVikarID 
	dim strOppdragID
	dim begDato
	dim endDato
	dim selectedPaymentType
	dim Timelonn
	dim Fakturapris
	dim VikarEndring
	dim strStartDato
	dim splittopt
	dim splittopt2

	' Connect to database
	SET Conn = GetClientConnection(GetConnectionstring(XIS, ""))

	' Check parameters AND put into variables
   	strVikarID = Request("VikarID")
	strOppdragID = Request("OppdragID")
	begDato = Request("begDato")
	endDato = Request("endDato")
	selectedPaymentType = Request("lstPaymentTypes")
	Timelonn = Request("Timelonn")
	Fakturapris = Request("Fakturapris")
	VikarEndring = Request("VikarEndring")
	strStartDato = Request("StartDato")
	splittopt = Request("splittopt")
	splittopt2 = Request("splittopt")

	'Start transaction
	Conn.Begintrans	
	
	if(len(begDato) = 0) then
		AddErrorMessage("Fra og med dato mangler.")
	end if
	
	if(len(endDato) = 0) then
		AddErrorMessage("Til og med dato mangler.")
	end if	

	if(len(strVikarID) = 0) then
		AddErrorMessage("VikarID mangler.")
	end if

	if(len(strOppdragID) = 0) then
		AddErrorMessage("Oppdragid mangler.")
	end if	

	If Request.Querystring("VikarEndring") = 1 Then ' SLETTE MELLOM DATOER

		if(HasError) then
			call RenderErrorMessage()
		end if
		
		' SLETTE MELLOM DATOER
		strSQL = "DELETE FROM DAGSLISTE_VIKAR " &_
			" WHERE VikarID = " & strVikarID &_
			" AND OppdragID = " & strOppdragID &_
			" AND Dato >= " & DbDate(begDato) &_
			" AND Dato <= " & Dbdate(endDato) &_
			" AND TimelisteVikarStatus < 5 "
	
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollBackTrans
			CloseConnection(Conn)
			SET Conn = nothing
			AddErrorMessage("En feil oppstod under sletting av timelister mellom datoer.")
			call RenderErrorMessage()	
		End if

	ElseIf Request.Querystring("VikarEndring") = 2 Then ' HINDRE FAKTURERING MELLOM DATOER

		if(HasError) then
			call RenderErrorMessage()
		end if

		' HINDRE FAKTURERING MELLOM DATOER
		strSQL = "UPDATE DAGSLISTE_VIKAR SET " &_
			"Fakturapris = 0, " &_
			"Fakturatimer = 0, " &_
			"loennsArt = '" & Request("lstPaymentTypes") & "'" &_
			" WHERE VikarID = " & strVikarID &_
			" AND OppdragID = " & strOppdragID &_
			" AND Dato >= " & DbDate(begDato) &_
			" AND Dato <= " & Dbdate(endDato) &_
			" AND TimelisteVikarStatus < 5 "

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollBackTrans
			CloseConnection(Conn)
			SET Conn = nothing
			AddErrorMessage("En feil oppstod under hindring av fakturering mellom datoer.")
			call RenderErrorMessage()	
		End if

	ElseIf Request.Querystring("VikarEndring") = 3 Then

		if(HasError) then
			call RenderErrorMessage()
		end if

		' ENDRE PRIS MELLOM DATOER
		Call fjernKomma(Fakturapris)
		Call fjernKomma(Timelonn)

		If (len(Timelonn) = 0 and len(Fakturapris) > 0)  Then
			strSQL = "UPDATE DAGSLISTE_VIKAR SET " &_
				"Fakturapris = " & Fakturapris &_
				" WHERE VikarID = " & strVikarID &_
				" AND OppdragID = " & strOppdragID &_
				" AND Dato >= " & DbDate(begDato) &_
				" AND Dato <= " & Dbdate(endDato) &_
				" AND TimelisteVikarStatus < 5 "
				
		ElseIf (len(Fakturapris) = 0 and len(Timelonn) > 0) Then
			strSQL = "UPDATE DAGSLISTE_VIKAR SET " &_
 				"Timelonn = " & Timelonn &_
				" WHERE VikarID = " & strVikarID &_
				" AND OppdragID = " & strOppdragID &_
				" AND Dato >= " & DbDate(begDato) &_
				" AND Dato <= " & Dbdate(endDato) &_
				" AND TimelisteVikarStatus < 5 "
		
		Elseif (len(Fakturapris) > 0 and len(Timelonn) > 0) then
			strSQL = "UPDATE DAGSLISTE_VIKAR SET " &_
 				"Timelonn = " & Timelonn &_
				", Fakturapris = " & Fakturapris &_
				" WHERE VikarID = " & strVikarID &_
				" AND OppdragID = " & strOppdragID &_
				" AND Dato >= " & DbDate(begDato) &_
				" AND Dato <= " & Dbdate(endDato) &_
				" AND TimelisteVikarStatus < 5 "
		Else
			AddErrorMessage("Du må fylle ut minst et av timelønn/timepris feltene.")
			call RenderErrorMessage()			
		End If

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollBackTrans
			CloseConnection(Conn)
			SET Conn = nothing
			AddErrorMessage("En feil oppstod under endring av pris mellom datoer.")
			call RenderErrorMessage()	
		End if

	ElseIf Request.Querystring("VikarEndring") = 4 Then ' ENDRE KLOKKESLETT MELLOM DATOER

		if(len(Request("tbxFraKl")) = 0) then
			AddErrorMessage("Fra klokkeslett mangler.")
		end if	

		if(len(Request("tbxTilKl")) = 0) then
			AddErrorMessage("Til klokkeslett mangler.")
		end if	

		if(len(Request("tbxLunsj")) = 0) then
			AddErrorMessage("Til klokkeslett mangler.")
		end if

		if(HasError) then
			call RenderErrorMessage()
		end if
		
		' ENDRE KLOKKESLETT MELLOM DATOER
		strSQL = "UPDATE DAGSLISTE_VIKAR SET " &_
			"starttid = '" & Request("tbxFraKl") &_
			"', sluttid = '" & Request("tbxTilKl") &_
			"', antTimer = " & Request("tbxTimerPrDag") &_
			", Fakturatimer = " & Request("tbxFaktTimerPrDag") &_
			", Lunch = '" & Request("tbxLunsj") & "'" &_	
			" WHERE VikarID = " & strVikarID &_
			" AND OppdragID = " & strOppdragID &_
			" AND Dato >= " & DbDate(begDato) &_
			" AND Dato <= " & Dbdate(endDato) &_
			" AND Not Datepart(Weekday, Dato) IN(1, 7)" &_
			" AND TimelisteVikarStatus < 5 "

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollBackTrans
			CloseConnection(Conn)
			SET Conn = nothing
			AddErrorMessage("En feil oppstod under endring av klokkeslett mellom datoer.")
			call RenderErrorMessage()	
		End if

	End If 'vikarendring
	'Conn.RollBackTrans
	Conn.CommitTrans
	CloseConnection(Conn)
	set Conn = nothing
	Response.Redirect "Vikar_timeliste_vis3.asp?VikarID=" & strVikarID & "&OppdragID=" & strOppdragID & "&FirmaID=" & strFirmaID & "&Startdato=" & strStartDato & "&kode=" & Request("kode") & "&splittuke=" & session("splittuke") & "&splittopt2=" & splittopt & "&splittopt=" & splittopt
%>