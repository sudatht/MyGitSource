<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	strOppdragID = Request.QueryString("OppdragID")
	strFirmaID = Request.QueryString("FirmaID")
	strVikarID = Request.QueryString("VikarID")

	' Connect to database
	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))

	' Se etter om det er lønnet eller fakturert
	strSQL = "SELECT lmax = MAX(Loennstatus), fmax = MAX(Fakturastatus), " &_
		" tmax = MAX(timelistevikarstatus) FROM DAGSLISTE_VIKAR" &_
		" WHERE OppdragID = " & Request.QueryString("OppdragID") &_
		" AND FirmaID = " & Request.QueryString("FirmaID") &_
		" AND VikarID = " & Request.QueryString("VikarID")

	Set rsStat = GetfireHoseRS(strSQL, Conn)
	Lmax = rsStat("lmax")
	Fmax = rsStat("fmax")
	Tmax = rsStat("tmax")		
	rsStat.Close: Set rsStat = Nothing

	'Timelistene skal ikke kunne slettes dersom lønns- eller fakturagrunlag er godkjent
	If Lmax > 1 or Fmax > 1 then 
		AddErrorMessage("Du kan ikke slette timelistene fordi det finnes timelister der lønn og/eller faktura er godkjent. Nedgrader godkjent lønn og/eller faktura for å slette.")
		call RenderErrorMessage()		
	End if

	'Timelistene skal ikke kunne slettes dersom det er timelister dersom de er godkjent av økonomi.
	If Tmax > 4 then
		AddErrorMessage("Du kan ikke slette fordi det finnes godkjente timelister (status 5). Nedgrader godkjente timelister for å slette.")
		call RenderErrorMessage()			
	End if

	'Start transaction
	Conn.Begintrans

	' Sletter alle timelister på denne vikaren
	strSQL = "DELETE FROM DAGSLISTE_VIKAR " &_
		"WHERE OppdragID = " & Request.QueryString("OppdragID") &_
		" AND FirmaID = " & Request.QueryString("FirmaID") &_
		" AND VikarID = " & Request.QueryString("VikarID")

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under sletting fra ukeliste for vikar.")
		call RenderErrorMessage()	
	End if

	strSQL = "UPDATE OPPDRAG_VIKAR SET " &_
		"Timeliste = 0 " &_
		"WHERE OppdragID = " & Request.QueryString("OppdragID") &_
		" AND FirmaID = " & Request.QueryString("FirmaID") &_
		" AND VikarID = " & Request.QueryString("VikarID")
	
	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under oppdatering av oppdrag-vikar.")
		call RenderErrorMessage()	
	End if

	Conn.CommitTrans
	CloseConnection(Conn)
	set Conn = nothing

	' Push another page
	Redir = "Vikar_timeliste_vis3.asp?VikarID=" & strVikarID & "&OppdragID=" & strOppdragID & "&FirmaID=" & strFirmaID & "&startdato=" & strStartdato & "&Frakode=" & session("frakode")
	Response.Redirect Redir
%>	