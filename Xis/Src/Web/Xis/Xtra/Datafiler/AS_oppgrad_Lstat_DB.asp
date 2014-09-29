<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit 
Response.Expires = 0
%>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<%
	If (HasUserRight(ACCESS_ADMIN, RIGHT_ADMIN) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	
	' declaring variables
	Dim Conn
	Dim strSQL
	Dim VikarID
	
	VikarID = request("VikarID")

	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))

	' oppdaterer alle timelister for  AS-konsulenter som har t.stat=5 og l.stat = 3
	Conn.BeginTrans

	strSQL = "UPDATE DAGSLISTE_VIKAR SET" &_
		" Loennstatus = 3," &_
		" LoennDato  = getDate() " &_ 
		" where loennstatus = 1 " &_
		" and TimelisteVikarStatus = 5 "&_
		" and vikarID in(" & VikarID & ")"

	if (ExecuteCRUDSQL(strSQL, Conn) = false) then
		Conn.RollbackTrans
		CloseConnection(Conn)
		set Conn = nothing	   
		AddErrorMessage("Feil under oppdatering av timeliste.")
		call RenderErrorMessage()		
	end if

	strSQL = "UPDATE VIKAR_UKELISTE SET" &_
		" Overfort_loenn_status = 3," &_
		" LoennDato  = getDate() "&_	
		" where Overfort_loenn_status = 1 " &_
		" and StatusID = 5 " &_
		" and vikarID in(" & VikarID & ")"
		
	if (ExecuteCRUDSQL(strSQL, Conn) = false) then
		Conn.RollbackTrans
		CloseConnection(Conn)
		set Conn = nothing	   
		AddErrorMessage("Feil under oppdatering av ukeliste.")
		call RenderErrorMessage()		
	end if

	Conn.CommitTrans
	CloseConnection(Conn)
	set Conn = nothing

	Response.redirect "As_vis.asp"
%>
