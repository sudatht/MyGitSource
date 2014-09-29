<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="../includes/Library.inc"-->
<%
'***************************************************************************************************
'Endringslogg
'Sist endret: 13.11.00  Endret av: LWS
'Endring: Gjort det mulig å velge avdeling når det legges til ekstra rader i variabel lønn
'**************************************************************************************************

'--------------------------------------------------------------------------------------------------
' Connect to database
'--------------------------------------------------------------------------------------------------

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

'--------------------------------------------------------------------------------------------------
' UPDATE OR NEW   VIKAR_LOEN_VARIABLE
'--------------------------------------------------------------------------------------------------

If Request("Ny") = "Ja" Then 

	'--------------------------------------------------------------------------------------------------
	' Check parameters and put into variables
	'--------------------------------------------------------------------------------------------------
	strVikarID = Request("VikarID")
	strOppdragID = Request("OppdragID")
	strLoennstakernr = strVikarID
	strDato = DbDate(Request("Dato") )
	strProsjektnr = Request("Prosjektnr")
	strLoennsartnr = Request("Loennsartnr")
	strAntall = Request("Antall")
	strSats = Request("Sats")
	strAvdeling = Request("Avdeling")
	strBeloep = Request("Beloep")
	strNavn = Request("Navn")
	strID = Request("ID")
	strSlett = Request("Slett")
	strEndre = Request("Endre")

	'Prosjektnr
	If strProsjektNr = "" Then 
		strProsjektnr = "Null"
	End if 	

	viskode = Request("Viskode")
	strFirmaID = session("FirmaID")
	If strFirmaID = "" Then
		strSQL = "Select FirmaID from OPPDRAG where OppdragID = " & strOppdragID
		Set f = conn.execute(strSQL)
		strFirmaID = f("FirmaID")
		f.Close: Set f = Nothing
	End If 'FirmaID = ""

	strOppdragID = session("OppdragID")
	strOppdragID = Request("OppdragID")
	strStatus = Request("status")

	If Request("Loenndato") = "" Then 
		strLoennDato = "NULL" 
	Else 
		strLoennDato = dbDate(Request("Loenndato"))
	End if	
	
	'--------------------------------------------------------------------------------------------------
	' Delete row in VIKAR_LOEN_VARIABLE
	'--------------------------------------------------------------------------------------------------
	If strSlett = "Ja" Then

		strSQL = "Delete from VIKAR_LOEN_VARIABLE where ID = "  & strID
		conn.Execute(strSQL)

	End IF 'deleting


	If  strLoennsartnr <> "" And strAntall <> "" And strSats <> "" And strAvdeling <> ""  Then

		'--------------------------------------------------------------------------------------------------
		' Update row in VIKAR_LOEN_VARIABLE
		'--------------------------------------------------------------------------------------------------
		If strEndre = "Ja" Then

			strBeloep = strAntall * strSats
			
			'Formatterer tallverdier
			call fjernKomma(strAntall)
			call fjernKomma(strSats)
			call fjernKomma(strBeloep)

			'Oppdaterer verdier
			strSQL = "Update VIKAR_LOEN_VARIABLE set" &_
				" Dato = " & strDato &_
				", Prosjektnr = " & strProsjektnr &_
				", Loennsartnr = '" & strLoennsartnr &_
				"', Antall = " & strAntall &_
				", Sats = " & strSats &_
				", Beloep = " & strBeloep &_
				", Avdeling = " & strAvdeling &_
				", Overfor_Loenn_Status = " & strStatus &_
				", Loenndato = " & strLoenndato &_
				" where ID = " & strID

			conn.Execute(strSQL)


		'--------------------------------------------------------------------------------------------------
		' Insert into database VIKAR_LOEN_VARIABLE
		'--------------------------------------------------------------------------------------------------
		ElseIf strEndre="Ny" Then 'insert (not delete or update)

			strBelop = strAntall * Request("Sats")

			IF strAvdeling = "" then strAvdeling = 0

			'Formatterer verdier
			Call fjernKomma(strBelop)
			Call fjernKomma(strAntall)
			Call fjernKomma(strSats)

			'Legger inn nye rader
			strSQL = "Insert into VIKAR_LOEN_VARIABLE (VikarID, Loennstakernr, Avdeling, Dato, Prosjektnr, Loennsartnr, Antall, Sats," &_
					"Beloep, FirmaID, OppdragID, Overfor_loenn_status, Loenndato, Timelistestatus, Nylinje) " &_
				"values (" &_
				strVikarID & "," &_
				strLoennstakernr & "," &_
				strAvdeling & "," &_
				strDato & "," &_
				strProsjektnr & ",'" &_
				strLoennsartnr & "'," &_
				strAntall & "," &_
				strSats & "," &_
				strBelop & "," &_
				strFirmaID & "," &_
				strOppdragID & "," &_
				strStatus & "," &_
				strLoenndato & "," &_
				"5, 1)"

			conn.Execute(strSQl)


		End If 'update or insert 

	End If 'fields have content

 	Response.Redirect "Vikar_varl_vis3.asp?VikarID=" & strVikarID & "&Avdeling=" & strAvdeling & "&viskode=" & Request("viskode") & "&OppdragID=" & strOppdragID

End If  'ny 


%>	