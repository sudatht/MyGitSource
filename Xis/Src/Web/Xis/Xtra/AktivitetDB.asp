<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim strAction
	dim lngVikarID
	dim lngFirmaID
	dim strOppdragID
	dim lngAktivitetID
	dim source
	dim rsFirma	
	dim cuid
	
	
	If  len(trim(Request("tbxAktivitetID"))) = 0 Then
		lngAktivitetID = 0
	Else
		lngAktivitetID = clng(Request("tbxAktivitetID"))
	End If

	If  len(trim(Request("tbxVikarID"))) = 0 Then
		lngVikarID = 0
	Else
		lngVikarID = clng(Request("tbxVikarID"))
	End If

	If  len(trim(Request("tbxFirmaID"))) = 0 Then
		lngFirmaID = 0
	Else
		lngFirmaID = clng(Request("tbxFirmaID"))
	End If

	If  len(trim(Request("tbxOppdragID"))) = 0 Then
		strOppdragID = 0
	Else
		strOppdragID = clng(Request("tbxOppdragID"))
	End If

	If  len(trim(Request("source"))) = 0 Then
		source = ""
	Else
		source = Request("source")
	End If

	If len(trim(Request("pbnDataAction"))) = 0 Then
	'Save is default action
		strAction = "Lagre aktivitet"
	Else
		strAction = trim(Request("pbnDataAction"))
	End If
	
	
	

'Response.Write "lngAktivitetID:" & lngAktivitetID & " lngVikarID:" & lngVikarID & " lngFirmaID:" & lngFirmaID & " strOppdragID:"  & strOppdragID & "<br>"

	If (lngAktivitetID = 0 AND lngVikarID = 0 AND lngFirmaID = 0 AND strOppdragID = 0) OR (len(source) = 0)  Then	
		AddErrorMessage("Systemfeil: Parametere mangler!")
		call RenderErrorMessage()
	End If

   ' Open database connection
   Set Conn = GetConnection(GetConnectionstring(XIS, ""))	
   
	' Action against database depending on Button pressed
	If strAction = "Lagre aktivitet" and lngAktivitetID = 0 Then
		
		' Create SQL-statement
		strSQL = "INSERT INTO [AKTIVITET]([AktivitetTypeID], [AktivitetDato], [VikarID], [FirmaID],[OppdragID], [RegistrertAvID], Notat ) " & _
			"Values( " & Request("dbxType") & "," & _
			DbDate( Request("tbxAktivitetDato") )& "," & _
			fixID(lngVikarID) & "," & _
			fixID(lngFirmaID)  & "," & _
			fixID(strOppdragID)  & "," & _
			Quote( Session("medarbID") )  & "," & _
			fixString(Request("txtNotat")) & ")"

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			AddErrorMessage("En eller flere feil oppstod under oppretting av aktivitet.")
			call RenderErrorMessage()
		End if

	ElseIf strAction = "Lagre aktivitet" and lngAktivitetID > 0 Then

		' Create SQL-statement
		strSql = "Update AKTIVITET SET " &_
			"[AktivitetTypeID] = " &  Request("dbxType") & _
			", [AktivitetDato] = " &  DbDate( Request("tbxAktivitetDato") ) & _
			", [Notat] = " &  Quote(padQuotes(Request("txtNotat"))) & _
			" WHERE [AktivitetID] = " & lngAktivitetID

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			AddErrorMessage("En eller flere feil oppstod under oppdatering av aktivitet.")
			call RenderErrorMessage()
		End if

	ElseIf (strAction = "Slette aktivitet" OR strAction = "slette") and lngAktivitetID > 0 Then
		' Create SQL-statement
		strSql = "DELETE AKTIVITET where AktivitetID = " & lngAktivitetID

		' Delete activity from database
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			AddErrorMessage("En eller flere feil oppstod under sletting av aktivitet.")
			call RenderErrorMessage()
		End if
		
	End If

	
	strSQL = "SELECT [SOCuid] FROM FIRMA WHERE FirmaID = " & lngFirmaID
	set rsFirma = GetFirehoseRS(strSQL, Conn)
	if(HasRows(rsFirma)) then
		cuid = 	rsFirma("SOCuid")
		rsFirma.close
	else
		cuid = 0
	end if
	set rsFirma = nothing

	CloseConnection(Conn)
	set Conn = nothing

	
	' Redirect to new page
	If source = "vikar"  Then
		strRedirect = "AktivitetVikar.asp"
	Elseif source = "oppdrag"  Then
		strRedirect = "AktivitetOppdrag.asp"
	Elseif source = "kontaktaktiviteter"  Then
		strRedirect = "AktivitetOppdragVikarList.asp"
	End If
	

	Response.Redirect strRedirect & "?cuid=" & cuid & "&FirmaID=" & lngFirmaID & "&VikarID=" &  lngVikarID & "&OppdragID="& strOppdragID & "&chkShowAutoRegister=" & trim(Request("chkShowAutoRegister"))
%>
