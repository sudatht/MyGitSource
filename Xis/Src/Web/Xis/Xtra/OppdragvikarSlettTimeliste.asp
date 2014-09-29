<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes/xis.rights.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim OppdragID : OppdragID = Request("tbxOppdragID")
	dim FirmaID : FirmaID = Request("tbxFirmaID")
	dim strSQL
	dim rsOppdragVikar
	dim Conn
	dim fmax
	dim tmax
	dim lmax
	dim rsMax

	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))
	
	' Sjekke om økonomi har rørt listene..
	strSQL = "SELECT fmax=Max(Fakturastatus), lmax=Max(Loennstatus), " &_
			"tmax=Max(timelistevikarstatus) FROM DAGSLISTE_VIKAR " &_
			"WHERE VikarID = " & Request("VikarID") &_
			" AND OppdragID = " & Request("OppdragID") &_
			" AND Dato <= '" & Request("tildato") &_
			"' AND Dato >= '" & Request("fradato") & "'"

	set rsmax = GetFirehoseRS(strSQL, Conn)
	fmax = rsmax("fmax")
	lmax = rsmax("lmax")
	tmax = rsmax("tmax")

	rsmax.Close: Set rsmax = Nothing
	
	'Sjekker om lønn eller faktura er godkjent. Du skal ikke kunne slette timelister hvis dette er tilfelle.
	If (fmax > 1 OR lmax > 1) Then
		AddErrorMessage("Du kan ikke slette timelistene fordi det finnes timelister der lønn og/eller faktura er godkjent. Nedgrader godkjent lønn og/eller faktura for å slette.")
	End if

	'Sjekker timelistestatus. Hvis timelister er godkjent av økonomi (status 5) skal de ikke kunne slettes.
	If tmax > 4 then
		AddErrorMessage("Du kan ikke slette fordi det finnes godkjente timelister (status 5). Nedgrader godkjente timelister for å slette.")
	End if

	if(HasError() = true) then
		CloseConnection(Conn)
		set Conn = nothing
		call RenderErrorMessage()
	end if
	
	Conn.BeginTrans	

	' slette timelistene...
	strSQL = "DELETE FROM DAGSLISTE_VIKAR " &_
		"WHERE VikarID = " & Request("VikarID") &_
		" AND OppdragID = " & Request("OppdragID") &_
		" AND Dato <= '" & Request("tildato") &_
		"' AND Dato >= '" & Request("fradato") & "'"

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollbackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("Feil oppstod under sletting av timeliste.")
		call RenderErrorMessage()
	End if

	' slette i oppdrag_vikar..
	strSQL = "DELETE FROM OPPDRAG_VIKAR WHERE OppdragVikarID = " & Request("Oppdragvikarid")
	
	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollbackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("Feil oppstod under sletting av Oppdrag info for vikar.")
		call RenderErrorMessage()
	End if

	'Dersom det fremdeles finnes vikarer tilknyttet oppdraget, skal vi benytte den siste utvidelsens tildato som sluttdato for oppdrag.
	strSQL = "SELECT MAX(Tildato) as [sluttdato] FROM [Oppdrag_Vikar] WHERE [OppdragID] = " & Request("OppdragID")
	set rsOppdragVikar = GetfireHoseRS(strSQL, Conn)
	if(not IsNull(rsOppdragVikar("sluttdato"))) then
		endOfTaskDate = dbdate(rsOppdragVikar("sluttdato").value)
		rsOppdragVikar.Close
		
		'Oppdater med slutt dato for siste oppdragsutvidelse
		strSQL = "UPDATE [Oppdrag] SET [Tildato] = " & endOfTaskDate & " WHERE [OppdragID] = " & Request("OppdragID")
		if (ExecuteCRUDSQL(strSQL, Conn) = false) then
			Conn.RollbackTrans
			CloseConnection(Conn)
			set Conn = nothing	   
			AddErrorMessage("Feil under oppdatering av oppdrag sluttdato.")
			call RenderErrorMessage()		
		end if
	end if
	set rsOppdragVikar = nothing

	Conn.CommitTrans
	CloseConnection(ConnTrans)
	set ConnTrans = nothing			
	'Response.End

	Response.redirect "WebUI/OppdragView.aspx?OppdragID=" & Request("OppdragID")
%>