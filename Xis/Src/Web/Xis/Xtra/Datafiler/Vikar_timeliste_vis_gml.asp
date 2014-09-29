<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<%
	If (HasUserRight(ACCESS_ADMIN, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim Conn
	dim strSQL
	dim SOBestilltAv
	dim BestilltAv

	Function OnlyDigits( strString )
	' Remove all non-nummeric signs FROM string
		If Not IsNull(strString) Then
			For idx = 1 To Len( strString) Step 1
				Digit = Asc( Mid( strString, idx, 1 ) )
				If (( Digit > 47 ) AND ( Digit < 58 )) Then
					strNewstring = strNewString & Mid(strString, idx,1)
				End If
			Next
		End If
		Onlydigits = strNewString
	End Function

	' Connect to database
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))
	
	' prosessing parameters
	'strVikarID = Request("VikarID")
	strAnsattnummer = Request("ansattnr")
	strFirmaID = Request("FirmaID")
	strOppdragID = Request("OppdragID")

	If Request("frakode") = "0" Then
		If Request("FraDato2") <> "" AND Request("TilDato2") <> "" Then
			session("FraDato2") = Request("FraDato2")
			session("TilDato2") = Request("TilDato2")
		Else
			AddErrorMessage("Fyllinn fra- og tildato!")
			call RenderErrorMessage()	
		End If
	End If

	If strAnsattnummer = "" AND Request("fra") = "" Then
		redir = "Vikar_timeliste_gml_vis.asp"
		Response.redirect redir
	End If

	%>
	<html>
		<head>
			<title>Arbeid med datafiler</title>
			<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
			<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">	
		</head>
		<body>
			<div class="pageContainer" id="pageContainer">
				<div class="content">
					<%
					strSQL = "SELECT VIKAR.Vikarid, Navn=(VIKAR.Fornavn + ' ' + VIKAR.Etternavn) FROM VIKAR LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid WHERE ansattnummer = " & strAnsattnummer
					Set rsNavn = GetFirehoseRS(strSQL, Conn)

					strNavn = rsNavn("Navn")
					strVikarID = rsNavn("Vikarid")

					rsNavn.Close
					Set rsNavn = Nothing
					
					' SQL for displaying data
					strSQL = "SELECT Linje=TimelisteVikarID, OppdragID, DAGSLISTE_VIKAR.VikarID, Dato, status=TimelisteVikarStatus, " &_
						"Starttid, sluttid, AntTimer, FirmaID, Timelonn, Fakturastatus, " &_
						"BestilltAv, SOBestilltAv, Fakturapris, Notat, Fakturatimer, Loennstatus, Lunch, LoennsArt " &_
						"FROM DAGSLISTE_VIKAR LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON DAGSLISTE_VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " &_
						"WHERE VIKAR_ANSATTNUMMER.ansattnummer = " & strAnsattnummer &_
						" AND Dato <= " & dbDate(session("TilDato2")) &_
						" AND Dato >= " & dbDate(session("FraDato2")) &_
						" AND TimelisteVikarStatus = 6 "&_
						" ORDER BY Dato"

					Set rsTimeliste = GetFirehoseRS(strSQL, Conn)

					' If no record exsists
					If (NOT HasRows(rsTimeliste)) Then
						Response.write "Har ikke timeliste! <br>"
						ingenlinjer = 1
					Else	'if no record
						ingenlinjer = 0
						' SQL for firma name
						strFirmaID = rsTimeliste("FirmaID")
						strSQL = "SELECT FirmaID, Firma FROM Firma WHERE FirmaID = " & rsTimeliste("FirmaID")
						Set rsFirma = GetFirehoseRS(strSQL, Conn)
						strFirma = rsFirma("Firma")
						rsFirma.Close
						Set rsFirma = Nothing

						' Find week number
						strStartDato = Request.QueryString("startDato")
						If strStartDato = "" Then
							strStartDato = rsTimeliste("Dato")
						End If
						strWeekNumber = Datepart("ww", strStartDato, 2, 2)

						strBestilltAv = rsTimeliste("BestilltAv")
						SOBestilltAv = rsTimeliste("SOBestilltAv")
						strFirmaID = rsTimeliste("FirmaID")

						' Form to register AND change
						' Display data
						%>
						<div class="listing">
							<table>
								<tr>
									<th colspan="14"><% Response.write strAnsattnummer & " " & strNavn & " - " & strFirma & " - oppdragid: " & strOppdragID  %></th>
								</tr>
								<tr>
									<th>OppdragId</th>
									<th>Ukedag</th>
									<th>Dato</th>
									<th>Startkl.</th>
									<th>SluttKl.</th>
									<th>Timer</th>
									<th>Lunsj</th>
									<th>Timelønn</th>
									<th>F.Timer</th>
									<th>F.gr.</th>
									<th>L.status</th>
									<th>F.status</th>
									<th>Lønnsart</th>
									<th>Notat</th>
								</tr>
								<%
								' Prepare for Loop
								antRecord = 0
								antLoop = 0
								neste = False

								If Request.QueryString("startDato") = "" Then
 									strStartDato = rsTimeliste("Dato")
								Else
 									strStartDato = Request.QueryString("startDato")
								End If
								'Response.write strStartDato & "<br>"
								If Request.QueryString("TLONN") = "" Then
									TLONN = rsTimeliste("Timelonn")
								Else
									TLONN = Request.QueryString("TLONN")
								End If

								Session("displayWeek")  = Datepart("ww", strStartDato,2,2)
								Session("begDato") = rsTimeliste("Dato")
								ukeTeller = 0
								nowWeek =  0
								strSum = 0
								strFaktSum = 0
								tt = 0

								' Loop through recordset
								do while not rsTimeliste.EOF
									antRecord = antRecord + 1
									rstimeliste.MoveNext
								loop
								rsTimeliste.MoveFirst

								do while not rsTimeliste.EOF

									antLoop = antLoop + 1

									If Datepart("ww", rsTimeliste("Dato"),2,2) = Session("DisplayWeek") Then 'And (rsTimeliste("Timelonn") = CInt(TLONN))
										neste = True
										displayWeek  = Datepart("ww", rsTimeliste("Dato"),2,2)
									End If
									dag = WeekDayName(Weekday(rsTimeliste("Dato"),2),,2)

									' Determin when to display rows (weeks before are collected at the bottom of loop)
									' Display until sunday else (create buttons for stored weeks ) AND create buttons for next weeks
									If neste Then
										If displayWeek = Datepart("ww", rsTimeliste("Dato"),2,2) Then
											strStarttid = Left(TimeValue(rsTimeliste("Starttid")),5)
											strSluttid = Left(TimeValue(rsTimeliste("Sluttid")),5)

											strAnttimer = rsTimeliste("AntTimer")
											pos1 = Instr(1, strAnttimer, ",")
											If pos1 > 0 Then
												strAnttimer = Left(strAntTimer, pos1 + 2)
											End If
											strLunch = Left(TimeValue(rsTimeliste("Lunch")),5)
											strTimelonn = rsTimeliste("Timelonn")

											strFakturatimer = rsTimeliste("Fakturatimer")
											pos1 = Instr(1, strFakturatimer, ",")
											If pos1 > 0 Then
												strFakturatimer = Left(strFakturaTimer, pos1 + 2)
											End If
											strFakturapris = rsTimeliste("Fakturapris")
											strSum = strSum + strAnttimer
											strFaktSum = strFaktSum + strFakturatimer
     										sluttdato = rsTimeliste("Dato")
											fStatus = rsTimeliste("Fakturastatus")
											%>
											<tr class="right">
												<td><% =rsTimeliste("OppdragID") %></td>
												<td><% =dag %></td>
												<td><%=DateValue(rsTimeliste("Dato"))%></td>
												<td><%=strStarttid %></td>
												<td><%=strSluttid %></td>
												<td><%=strAnttimer %></td>
												<td><%=strLunch %>
												<% 
												name = CStr(OnlyDigits(rsTimeliste("Dato"))) 
												%>
												<td><% =rsTimeliste("Timelonn") %></td>
												<td><% =strFakturatimer %> </td>
												<td><% =rsTimeliste("FakturaPris") %></td>
												<td><% =rsTimeliste("Loennstatus") %></td>
												<td><% =rsTimeliste("FakturaStatus") %></td>
												<td><% =rsTimeliste("LoennsArt") %>&nbsp;</td>
												<td><% =rsTimeliste("Notat") %>&nbsp;</td>								
												<% 
												strStatus = rsTimeliste("Status") 
												%>
											</tr>
											<%
										Else
											neste = False
 										End If
									End If	 'første dato er startDato, visning av uke

									' This part runs allways in the loop but collects weeks only before display date
									If Not nowWeek = Datepart("ww", rsTimeliste("Dato"),2,2) Then
										If rsTimeListe("Fakturastatus") = 3 AND rstimeliste("Loennstatus") = 3 Then
											fff = "<table bgcolor=BLACK>"
										ElseIf rsTimeliste("status") = 5 Then
											fff = "<table bgcolor=GREEN>"
										ElseIF rsTimeliste("status") = 4 Then
											fff = "<table bgcolor=YELLOW>"
										ElseIF rsTimeliste("status") = 3 Then
											fff = "<table bgcolor=CYAN>"
										ElseIF rsTimeliste("status") = 2 Then
											fff = "<table bgcolor=#FF9900>"
										Else
											fff = "<table bgcolor=RED>"
										End If
										ukeTeller = ukeTeller + 1
										Redim Preserve ukeNrliste(ukeTeller)
										Redim Preserve ukeDatoListe(ukeTeller)
										Redim Preserve ukeFontListe(ukeTeller)
										Redim Preserve ukeTimelonnListe(ukeTeller)
										ukeNrliste(ukeTeller - 1) = Datepart("ww", rsTimeliste("Dato"),2,2)
										ukeDatoListe(uketeller - 1) = rsTimeliste("Dato")
										ukeFontListe(ukeTeller - 1) = fff
										ukeTimelonnListe(ukeTeller - 1) = rsTimeliste("Timelonn")
										nowWeek = Datepart("ww", rsTimeliste("Dato"), 2, 2)
										nylonn = 0
									End If
									' End of loop
									rsTimeliste.MoveNext
								loop

								dim ukeNr 
								ukeNr = Datepart("yyyy", strStartDato, 2, 2) & Datepart("ww", strStartDato, 2, 2)

								'Vanlige lønnede timer
								strSQL = "SELECT Antall = Sum(Antall) " & _
									" FROM VIKAR_UKELISTE, H_LOENNSART  " & _
									" WHERE VIKAR_UKELISTE.Loennsartnr = H_LOENNSART.Loennsartnr  " &_
									" AND Ukenr = " & ukeNr &_
									" AND VikarID = " & strVikarID &_
									" AND LoennRate = '1.0' "

								Set rsLn = conn.Execute(strSQL)
								If Not rsLn.EOF Then
									strVanlig = rsLn("Antall")
								Else
									strVanlig = ""
								End If

								'Vanlige fakt timer
								strSQL = "SELECT FakturertAntall = Sum(Antall) " & _
									" FROM VIKAR_UKELISTE, H_LOENNSART, H_FAKTURA_TYPE  " & _
									" WHERE VIKAR_UKELISTE.Loennsartnr = H_LOENNSART.Loennsartnr  " &_
									" AND H_FAKTURA_TYPE.FAKTURATYPE = VIKAR_UKELISTE.FAKTURATYPE  " &_
									" AND Ukenr = " & ukeNr &_
									" AND VikarID = " & strVikarID &_
									" AND Faktureres = 1 " &_ 
									" AND VIKAR_UKELISTE.fakturaSats = '1.0' "						

								Set rsLn = conn.Execute(strSQL)
								If Not rsLn.EOF Then
									strFaktVanlig = rsLn("FakturertAntall")
								Else
									strFaktVanlig = ""
								End If
								
								'Lønn overtid steg 1 timer
								strSQL = "SELECT Antall=SUM(Antall) " & _
									" FROM VIKAR_UKELISTE, H_LOENNSART " & _
									" WHERE VIKAR_UKELISTE.Loennsartnr  = H_LOENNSART.Loennsartnr " &_
									" AND Ukenr = " & ukenr &_
									" AND VikarID = " & strVikarID &_
									" AND LoennRate = '1.5' "
								
								Set rsLn = conn.Execute(strSQL)
								If Not rsLn.EOF Then
									str50 = rsLn("Antall")
									If str50 = 0 Then 
										str50 = ""
									end if
								End If								

								'Fakturerbar overtid timer
								strSQL = "SELECT FakturertAntall = Sum(Antall), VIKAR_UKELISTE.fakturaSats, VIKAR_UKELISTE.GroupKey " & _
									" FROM VIKAR_UKELISTE, H_LOENNSART, H_FAKTURA_TYPE  " & _
									" WHERE VIKAR_UKELISTE.Loennsartnr = H_LOENNSART.Loennsartnr  " &_
									" AND H_FAKTURA_TYPE.FAKTURATYPE = VIKAR_UKELISTE.FAKTURATYPE  " &_
									" AND Ukenr = " & ukenr &_
									" AND VikarID = " & strVikarID &_
									" AND Faktureres = 1 " &_ 
									" AND VIKAR_UKELISTE.fakturaSats > '1.0'  " &_	
									" GROUP BY VIKAR_UKELISTE.fakturaSats, VIKAR_UKELISTE.GroupKey " &_
									" Order by VIKAR_UKELISTE.GroupKey, VIKAR_UKELISTE.fakturaSats ASC "

								Set rsLn = conn.Execute(strSQL)

								If Not rsLn.EOF Then
									WHILE NOT (rsLn.EOF)
										select case rsLn("GroupKey").value
											case "step1" : 
												strFaktantallSteg1 = rsLn("FakturertAntall").value
												strFaktRateSteg1 = rsLn("fakturaSats").value												
											case "step2" : 
												strFaktAntallSteg2 = rsLn("FakturertAntall").value
												strFaktRateSteg2 = rsLn("fakturaSats").value
										end select
										rsLn.Movenext
									WEND
								End If
								rsLn.close
								set rsLn = nothing

								'Lønn overtid steg 2 timer
								strSQL = "SELECT Antall = Sum(Antall) " & _
									" FROM VIKAR_UKELISTE, H_LOENNSART " & _
									" WHERE VIKAR_UKELISTE.Loennsartnr  = H_LOENNSART.Loennsartnr " &_
									" AND Ukenr = " & ukeNr &_
									" AND VikarID = " & strVikarID &_
									" AND LoennRate = '2.0' "

								Set rsLn = conn.Execute(strSQL)
								If Not rsLn.EOF Then
									str100 = rsLn("Antall")
									If str100 = 0 Then 
										str100 = ""
									end if
								End If
								
								rsLn.Close
								set rsLn = nothing
								
								session("EndDato") = sluttdato

' Display SAVE AND NEW  BUTTON
%>
<tr>
<FORM NAME=TLISTE ACTION="Vikar_timeliste_lagre3.asp?Lagringskode=oppgrad&Startdato=<% =strStartdato %>&Sluttdato=<% =Sluttdato %>&VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&Timelonn=<% =strTimelonn %>&FirmaID=<% =strFirmaID %>&UkeNr=<% =strWeekNumber %>&limitDato=<% =limitDato %>&BestilltAV=<% =strBestilltAv %>&SOBestilltAV=<% =SOBestilltAv %>&Fakturapris=<% =strFakturapris %>&kode=1"  METHOD=POST ID="Form1">
<%

'Response.write strSum

If strVanlig = "" Or IsNull(strVanlig) Then
	 sss = strSum
	strVanlig = strSum
   Else
	sss = strVanlig
   End If
%>
<input type=hidden name=SUMM VALUE=<% =sss %> ID="Hidden1">
<input type=hidden name=SUMM2 VALUE=<% =strSum %> ID="Hidden2">

<%
If strFaktVanlig = "" Or IsNull(strFaktVanlig) Then
	 sss = strFaktSum
	strFaktVanlig = strFaktSum
   Else
	sss = strFaktVanlig
   End If
%>
<input type=hidden name=FAKTSUMM VALUE=<% =sss %> ID="Hidden3">
<input type=hidden name=FAKTSUMM2 VALUE=<% =strFaktSum %> ID="Hidden4">

<% '---------------viser samlesummer i tabell------------ %>
<th colspan=4 class=right>Sum:<th class="right"><% =strSum %></th>
<% sumsum = strSum - 37.5
 If sumsum > 0 Then  %>
<th  class=right>Overtid:<th class="right"><% =Left(sumsum,5) %></th>
<% Else %>
<TD><TD>
<% End If %>
<th class="right"><% =strFaktSum %></th>
<% '------------------------------------------------- %>
	<tr>
		<th colspan=4>
        <th class="right">Vanlig:</th>
        <th class="right"><% =strVanlig %></th>
		<th><th>
		<th class="right" ><% =strFaktVanlig %></th>
	</tr>
	<tr>
		<th></th>
		<th></th>
		<% 
		'Viser overtid i tabell
		If str50 <> "" Then 
			%>
				<th class="right" >50 %:</th><th class="right" ><% =str50 %></th>
			<% 
			Else 
			%>
			<th><th>
			<% 
		End If 
		If str100 <> "" Then 
			%>
				<th class="right" >100 %:</th><th class="right" ><% =str100 %></th>
			<% 
		Else 
			%>
			<th></th>
			<% 
		End If 
		%>
		<th></th>
		<% 
		If strFaktAntallSteg1 <> "" Then 
			%>
			<th class="right" >Overtid steg 1:</th><th class="right" ><% =strFaktAntallSteg1 %></th>
			<% 
		Else 
			%>
			<th><th>
			<% 
		End If 
		If strFaktAntallSteg2 <> "" Then 
			%>
			<th class="right" >Overtid steg 2:</th><th class="right" ><% =strFaktAntallSteg2 %> </th>
			<% 
		Else 
			%>
			<th><th>
			<% 		
		End If
	%>
	</table>
</form>
<%
' Display buttons for navigating between weeks
%>
<table cellpadding='0' cellspacing='0' ID="Table2">
	<tr><th>Uke:
<% 
For i = 0 To ukeTeller - 1
uk = uk + 1
If uk = 20 Then
	Response.write "<tr>"
	uk = 1
End If
 If ukeNrliste(i) = Session("displayweek") Then %>
<th><% =ukeFontListe(i) %><% =ukeNrListe(i)  %></table></th>
<% Else %>
<FORM ACTION="Vikar_timeliste_vis_gml.asp?ansattnr=<%=strAnsattnummer%>&OppdragID=<%=strOppdragID %>&FirmaID=<% =strFirmaID %>&startDato=<% =ukeDatoListe(i) %>&TLONN=<% =ukeTimelonnListe(i) %>" METHOD=POST ID="Form2">
<th><% =ukeFontListe(i) %><INPUT TYPE=SUBMIT VALUE="<% =ukeNrListe(i)  %>" ID="Submit1" NAME="Submit1"></table></th>
</form>
<% End If %>
<% Next %>
</table>
<%
end If
				' TILBAKEKNAPPER
				%>
				<table cellpadding='0' cellspacing='0' ID="Table3">
					<FORM ACTION=Vikar_timeliste_list_gml.asp METHOD=POST ID="Form3">
						<th colspan=2><INPUT TYPE=SUBMIT VALUE="                Nytt søk                " ID="Submit2" NAME="Submit2"></th>
					</form>
				</table>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>