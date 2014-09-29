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
	dim strVikarID
	dim strDato
	dim strLoennDato
	dim nedgrad
	dim viskode
	dim strAvd
	dim strdato1
	dim strOppdragID
	dim strFirmaID

	' Check parameters AND put into variables
	strVikarID = Request("VikarID")
	strOppdragID = Request.QueryString("OppdragID")
	strFirmaID = Request("FirmaID")
	strDato = dbDate(Date)
	strLoennDato = Request("LoennDato")
	nedgrad = Request("nedgrad")
	viskode = session("viskode")
	strAvd = Request("avd")
	strdato1 = Request("dato1")

	'se strVikarID & "<br>"
	'se strDato & "<br>"
	'se strLoennDato & "<br>"
	'se nedgrad & "<br>"

	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))

If nedgrad = "Ja" Then  'kalles fra Vikar_timeliste_list3.asp

	' Start transaction
	Conn.Begintrans

	If Request("Loennstatus") = 3 Then
		LSTAT = 2
		LDATO = dbDate(strLoennDato)
	ElseIf Request("Loennstatus") = 2 Then
		LSTAT = 1
		LDATO = "NULL"
	End If

	' UPDATE status in DAGSLISTE_VIKAR    (timelisten)
	strSQL = "UPDATE DAGSLISTE_VIKAR SET " &_
	"Loennstatus = " & LSTAT & ", " &_
	"Loenndato = " & LDATO  &_
	" WHERE VikarID = " & strVikarID &_
	" AND Loenndato = " & dbDate(strLoennDato)

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under oppdatering av dagsliste for vikar.")
		call RenderErrorMessage()	
	End if

	' UPDATE status in VIKAR_LOEN_VARIABLE
	strSQL = "UPDATE VIKAR_LOEN_VARIABLE SET " &_
		"Overfor_Loenn_status = " & LSTAT & ", " &_
		"Loenndato = " & LDATO  &_
		" WHERE VikarID = " & strVikarID &_
		" AND Loenndato = " & dbDate(strLoennDato)

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under oppdatering av lønn variable for vikar.")
		call RenderErrorMessage()	
	End if

	' UPDATE status in VIKAR_UKELISTE
	strSQL = "UPDATE VIKAR_UKELISTE SET " &_
		"Overfort_loenn_status = " & LSTAT & ", " &_
		" Loenndato = " & LDATO  &_
		" WHERE VikarID = " & strVikarID &_
		" AND Loenndato = " & dbDate(strLoennDato)

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under oppdatering av ukelister for vikar.")
		call RenderErrorMessage()	
	End if

	Conn.CommitTrans
	'Conn.RollBackTrans
	CloseConnection(Conn)
	set Conn = nothing

	' redirect back to calling page
	Response.Redirect "Vikar_timeliste_list3.asp?viskode=" & viskode & "&avd=" & strAvd & "&dato1=" & strdato1

Else  'ikke nedgradering, men oppgradering

	' Start transaction
	Conn.Begintrans
	
	' UPDATE status in DAGSLISTE_VIKAR (timelisten)
	strSQL = "UPDATE DAGSLISTE_VIKAR SET " &_
		"Loennstatus = 2, Loenndato = " & strDato &_
		" WHERE Loennstatus = 1 " &_
		" AND Timelistevikarstatus = 5 " &_
		" AND VikarID = " & strVikarID 

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under oppdatering av dagslister for vikar.")
		call RenderErrorMessage()	
	End if

	' UPDATE VIKAR_LOEN_VARIABLE (called from Vikar_varl_vis3)
	strSQL = "UPDATE VIKAR_LOEN_VARIABLE SET " &_
			"Overfor_Loenn_status = 2, " &_
			"Loenndato= " & strDato &_
			" WHERE VikarID = " & strVikarID &_
			" AND Timelistestatus = 5 " &_
			" AND Overfor_Loenn_status = 1 "

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under oppdatering av lønn variable for vikar.")
		call RenderErrorMessage()	
	End if

	' UPDATE VIKAR_UKELISTE (called from Vikar_varl_vis3)
	strSQL = "UPDATE VIKAR_UKELISTE SET " &_
		"Overfort_loenn_status = 2," &_
		" Loenndato = " & strDato &_
 		" WHERE VikarID = " & strVikarID &_
		" AND Overfort_loenn_status = 1" &_
		" AND StatusID = 5"

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollBackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("En feil oppstod under oppdatering av ukelister for vikar.")
		call RenderErrorMessage()	
	End if

	Conn.CommitTrans
	CloseConnection(Conn)
	set Conn = nothing

	' Push another page
	redir = "Vikar_varl_vis3.asp?VikarID=" & strVikarID & "&OppdragID=" & strOppdragID & "&firmaID=" & strfirmaID

	Response.Redirect redir

End If 'nedgradering eller oppgradering
%>