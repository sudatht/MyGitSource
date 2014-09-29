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
	dim strSQL
	dim tildato
	dim ukedagnr

	' Check parameters AND put into variables
	tildato = Request("limitdato")
	tildato = DateValue(tildato)

	ukedagnr = WeekDay(tildato, 2)
	If ukedagnr < 4 Then
		tildato = tildato - ukedagnr
	Else
		tildato = tildato + (7 - ukedagnr)
	End If

	ukenr = Datepart("ww", tildato, 2)

	' Connect to database
	SET Conn = GetClientConnection(GetConnectionstring(XIS, ""))	

	Conn.BeginTrans
	
	' UPDATE DAGSLISTE_VIKAR
	strSQL = "UPDATE DAGSLISTE_VIKAR SET " &_
		"Timelistevikarstatus = 6 " &_
		"WHERE Fakturastatus = 3 " &_
		"AND Loennstatus = 3 " &_
		"AND Timelistevikarstatus = 5 " &_
		"AND Dato <= " & dbDate(tildato)

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("Feil oppstod under setting av timelistestatus til gammel på timelister.")
		call RenderErrorMessage()
	End if

	' UPDATE VIKAR_UKELISTE
	strSQL = "UPDATE VIKAR_UKELISTE SET " &_
		"StatusID = 6 " &_
		"WHERE Overfort_fakt_status = 3 " &_
		"AND Overfort_loenn_status = 3 " &_
		"AND StatusID = 5 " &_
		"AND Ukenr <= " & ukenr

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("Feil oppstod under setting av timelistestatus til gammel på timelister.")
		call RenderErrorMessage()
	End if

	Conn.CommitTrans
	CloseConnection(Conn)
	set Conn = nothing

	' Push another page
	redir = "timelisteMeny.asp"

	Response.redirect redir
%>