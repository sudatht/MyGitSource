<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="includes/xis.rights.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	 
	dim brukerID
	dim KundeID
	dim oppdragID
	dim strParam
	dim kunde
	dim vikarID
	dim kode ' Kode = 1 oppdrag, 2 = kunde (utgått), 3 = vikar
	
	brukerID	= Session("BrukerID")
	kode		= Request("kode")
	KundeID		= Request("kundeNr")
	vikarID		= Request("vikarnr")
	oppdragID	= Request("oppdragID")

	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	if (kode = 1) Then 'Oppdrag
		strSQL = "INSERT INTO HOTLIST( BrukerID, navnID, navn, status, oppdragID)" &_           
		"Values( " & brukerID & "," & KundeID & "," &_
		"'" & Request("kundeNavn") & "'," & kode & "," & oppdragID & ")" 
	ElseIf (kode = 3) Then  'vikar
		strSQL = "INSERT INTO HOTLIST( BrukerID, navnID, navn, status)" &_           
		"Values( " & brukerID & "," & vikarID & "," &_
		"'" & Request("vikarNavn") & "'," & kode & ")"
	End if  
	  
	If ExecuteCRUDSQL(strSQL, Conn) = false then
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("Feil oppstod under hotlist tilknytning.")
		call RenderErrorMessage()
	End if	

	CloseConnection(Conn)
	set Conn = nothing

	if kode = 1 Then
		Response.redirect "WebUI/OppdragView.aspx?OppdragID=" & oppdragID
	ElseIf kode = 3 Then
		Response.redirect "VikarVis.asp?vikarID=" & vikarID
	End If
	%>