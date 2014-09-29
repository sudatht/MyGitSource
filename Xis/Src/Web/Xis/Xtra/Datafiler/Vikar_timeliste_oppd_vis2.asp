<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<%
	dim conn
	dim strSQL
	dim kkk
	dim p
	dim pp
	dim k
	dim strVikarID	
	dim strNavn
	dim strOppdragID
	dim strLinje
	dim strFirmaId
	dim strStartDato
	dim strAntTimer
	dim status_kode
	dim strTimeLonn
	dim strFakturapris
	dim strNotat
	dim strBestilltav
	dim strSOBestilltav
	dim strKontaktNavn
	dim rsNavn
	dim kontaktSQL	
	dim personRs 			
	dim cts
	
	Function OnlyDigits( strString )
	' Remove all non-nummeric signs FROM string
		If Not IsNull(strString) Then
		For idx = 1 To Len( strString) Step 1
		Digit = Asc( Mid( strString, idx, 1 ) ) 
		If (( Digit > 47 ) AND ( Digit < 58 )) Then 
		strNewstring = strNewString & Mid(strString,idx,1)
		End If
		Next
		End If
		Onlydigits = strNewString
	End Function

%>
<html>
	<head>
		<title>Arbeid med datafiler</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script type="text/javascript" language="javascript" src="../js/javascript.js"></script>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Alle timelister for kontaktperson</h1>
			</div>
			<div class="content">
			<% 

			kkk = Request.QueryString("Notat") 
   				p = Len(kkk) - 2
			If p > 1  Then 
				kk = Mid(kkk,2,p)
			End If

		' Connect to database
		Set Conn = GetConnection(GetConnectionstring(XIS, ""))
			
		' prosessing parameters
		If lenb(Request.QueryString("VikarID")) = 0 Then
			strVikarID = Request.Form("VikarID")
			strNavn = Request.Form("Navn")
			strOppdragID = Request.Form("OppdragID")
			strLinje = Request.Form("Linje")
			strFirmaId = Request.Form("FirmaID")
		Else
			strVikarID = Request.QueryString("VikarID")
			strNavn = Request.QueryString("Navn")
			strOppdragID = Request.QueryString("OppdragID")
			strLinje = Request.QueryString("Linje")
			strFirmaId = Request.QueryString("FirmaID")
		End IF

		strStartDato = Request.QueryString("StartDato")
		strAntTimer = Request.QueryString("AntTimer")	
		status_kode = Request.QueryString("kode")
		strTimeLonn = Request.QueryString("Timelonn")
		strFakturapris = Request.QueryString("Fakturapris")
		strNotat = Request.QueryString("Notat")
		strBestilltav = Request.QueryString("Bestilltav")
		strSOBestilltav = Request.QueryString("SOBestilltav")

		if (LenB(strBestilltav) = 0 AND LenB(strSOBestilltav) > 0) then
			kontaktSQL = " SOBestilltAv = " & strSOBestilltav
		elseif (LenB(strBestilltav) > 0 AND LenB(strSOBestilltav) = 0) then
			kontaktSQL = " BestilltAv = " & strBestilltav
		end if

		' SQL for displaying data
		status_kode = 6 'Yippi, magic number!

		strSQL = "SELECT Linje=TimelisteVikarID, VIKAR.VikarID,Dato, status=TimelisteVikarStatus, " &_
			"Starttid, sluttid, AntTimer, FirmaID, TimeLonn, DAGSLISTE_VIKAR.Notat,  DAGSLISTE_VIKAR.OppdragID,FakturaPris, Fakturastatus, " &_
			"Vikarnavn=(VIKAR.Fornavn + ' ' + VIKAR.Etternavn), DAGSLISTE_VIKAR.Bestilltav " &_
			"FROM DAGSLISTE_VIKAR, VIKAR " &_
			"WHERE " &_
			kontaktSQL &_
			" AND DAGSLISTE_VIKAR.VikarID = VIKAR.VikarID " &_
			" AND TimelisteVikarStatus < " & status_kode &_
			" ORDER BY Dato"

		Set rsTimeliste = GetFirehoseRS(strSQL, Conn)

		' If no record exsists
		If HasRows(rsTimeliste) = false Then 
			Response.write "<H4><br> " & strNavn & " har ikke timeliste! <br></H4>"
		Else
			strVikarID = rsTimeliste("VikarID")
			strOppdragID = rsTimeliste("OppdragID")

			if (LenB(strBestilltav) = 0 AND LenB(strSOBestilltav) > 0) then
				Set cts = server.CreateObject("Integration.SuperOffice")
				set rsNavn = cts.GetPersonSnapshotById(clng(strSOBestilltav))
				if (not rsNavn.EOF) then
					if (isnull(rsNavn("middlename"))) then
						kontaktperson = rsNavn("firstname") & " " & rsNavn("lastname")
					else
						kontaktperson = rsNavn("firstname") & " " & rsNavn("middlename") & " " & rsNavn("lastname")
					end if
				end if
				rsNavn.close
				set rsNavn = nothing
				set cts = nothing
			elseif (LenB(strBestilltav) > 0 AND LenB(strSOBestilltav) = 0) then
				' SQL for finding contactpersons name
				strSQL = "SELECT Kontaktnavn=(Fornavn + ' ' + Etternavn) FROM Kontakt WHERE KontaktID = " & rsTimeliste("Bestilltav")
				Set rsNavn = GetFirehoseRS(strSQL, Conn)
				strKontaktNavn = rsNavn("Kontaktnavn")
				rsNavn.Close
				set rsNavn = nothing
			end if

			' SQL for firma name
			strFirmaID = rsTimeliste("FirmaID")
			strSQL = "SELECT Firma FROM Firma WHERE FirmaID = " & rsTimeliste("FirmaID")
			Set rsFirma = GetFirehoseRS(strSQL, Conn)
			strFirma = rsFirma("Firma")
			rsFirma.Close
			set rsFirma = nothing

			' Find week number
			If lenb(strStartDato) = 0 Then
				strStartDato = rsTimeliste("Dato")	
			End If
			strWeekNumber = Datepart("ww", strStartDato, 2)

		' Form to register AND change
		' If edit find the right row to display
		If (lenb(strLinje) > 0) Then ' edit

			strSQL = "SELECT Dato, Starttid, Sluttid " &_ 
				"FROM DAGSLISTE_VIKAR " &_
				"WHERE TimelisteVikarID = " & strLinje

			Set rsLinje = GetFirehoseRS(strSQL, Conn)
			strStarttid = Left(TimeValue(rsLinje("Starttid")), 5)
			strSluttid = Left(TimeValue(rsLinje("Sluttid")), 5)
			strUkedag = WeekDayName(Weekday(rsLinje("Dato"), 2) , , 2)
		End If 	

		' Form to register AND change
		Response.write "<br><H4>" & strKontaktnavn & ",  " & strFirma & "</H4>"

		If Request.QueryString("Endre") <> "" Then

			kk = strStartDato
			pp = Request.QueryString("Dato")

			If kk = pp Then
				kk = ""
			End If

			%>
			<div class="listing">
				<table>
					<FORM NAVN="OPPDRAG" ACTION="Vikar_timeliste_oppd_db2.asp" METHOD=POST>
						<input name="VikarID" TYPE=HIDDEN VALUE=<% =strVikarID %> >
						<input name="OppdragID" TYPE=HIDDEN VALUE="<% =strOppdragID %>" >
						<input name="FirmaID" TYPE=HIDDEN VALUE="<% =strFirmaID %>" >
						<input name="StartDato" TYPE=HIDDEN VALUE="<% =kk %>" >
						<input name="Endre" TYPE=HIDDEN VALUE=<% =Request.Querystring("Endre") %>>
						<input name="Linje" TYPE=HIDDEN VALUE="<% =strLinje %>">
						<input name="Bestilltav" TYPE=HIDDEN VALUE="<% =strBestilltav %>">

						<tr>
							<th>Ukedag<th>Dato<th>Fratid<th>Tiltid<th>Timelønn<th>Fakt.pris</th>
						<tr>
							<th><% If strLinje <> "" Then Response.write strUkedag Else Response.write "" %></th>
							<th><input type=text size=8 NAME=Dato <% If strLinje <> "" Then Response.write "VALUE=" & Request.QueryString("Dato") %>  ONBLUR="dateCheck(this.form,this.name) ></th>
							<th><input type=text size=6 NAME=Starttid <% If strLinje <> "" Then Response.write "VALUE=" & strStarttid %> ONBLUR=" timeCheck(this.form,this.name)" ></th>
							<th><input type=text size=6 NAME=Sluttid <% If strLinje <> "" Then Response.write "VALUE=" & strSluttid %> ONBLUR="timeCheck(this.form,this.name)" ></th>
							<th><input type=text size=6 NAME=Timelonn <% If strLinje <> "" Then Response.write "VALUE=" & strTimelonn %> ></th>
							<th><input type=text size=6 NAME=Fakturapris <% If strLinje <> "" Then Response.write "VALUE=" & strFakturapris %> ></th>
							<tr><th colspan=7 ><input type=text size=53 NAME=Notat <% If strLinje <> "" Then Response.write "VALUE=" & strNotat %> ></th>
						<tr>
							<th colspan=7 ><INPUT TYPE=SUBMIT	VALUE="                           Lagre                                   "><INPUT TYPE=RESET ></th>
					</form>
				</table>
			</div>
			<% 
		End If  'endre eller ny 
		If strLinje <> "" Then rsLinje.Close 
		' Display data
		%>
		<table>
			<tr class=right>
				<th>Endre</th>
				<th>Dag</th>
				<th>Dato</th>
				<th>Starttid</th>
				<th>Slutttid</th>
				<th>Timer</th>
				<th>Timelønn</th>
				<th>Fakt.gr.</th>
				<th>Status</th>
				<th>Slett</th>
				<th>Vikar</th>
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

			ukeTeller = 0
			nowWeek =  0 

			' Loop through recordset
			do while not rsTimeliste.EOF
				antRecord = antRecord + 1
				rstimeliste.MoveNext	
			loop	
			rsTimeliste.MoveFirst
		
			do while not rsTimeliste.EOF 

				antLoop = antLoop + 1

				If CStr(rsTimeliste("Dato")) = CStr(strStartDato) Then
					neste = True
					displayWeek  = Datepart("ww",rsTimeliste("Dato"), 2)
				End If
				dag = WeekDayName(Weekday(rsTimeliste("Dato"), 2) , , 2)

				' Determin when to display rows (weeks before are collected at the bottom of loop)
				' Display until sunday else (create buttons for stored weeks ) AND create buttons for next weeks 
				If neste Then
					If displayWeek = Datepart("ww",rsTimeliste("Dato"),2) Then
						strStarttid = Left(TimeValue(rsTimeliste("Starttid")),5)
						strSluttid = Left(TimeValue(rsTimeliste("Sluttid")),5)
						strAnttimer = rsTimeliste("AntTimer")
						%>
						<TR class=right>
							<% 
							'link to update
							update = "Endre=Ja&VikarID=" & rsTimeliste("VikarID") & "&OppdragID=" & strOppdragID & "&Linje=" & rsTimeliste("Linje") &_
							"&FirmaID=" & strFirmaID & "&Fratid=" & strStarttid & "&Tiltid=" & strSluttid &_
							"&Dato=" & rsTimeliste("Dato") & "&AntTimer=" & strAnttimer & "&StartDato=" & strStartDato &_
							"&kode=" & status_kode & "&Timelonn=" & rsTimeliste("Timelonn") &_
							"&Fakturapris="  & rsTimeliste("FakturaPris") & "&Notat='" & rsTimeliste("Notat") & "'&Bestilltav=" & strBestilltav
							%>
							<th><A HREF="Vikar_timeliste_oppd_vis2.asp?<% =update %>" >Endre</A></th>
							<th><% =dag %></th>
							<th><%=DateValue(rsTimeliste("Dato"))%></th>
							<th><%=strStarttid %></th>
							<th><%=strSluttid %></th>
							<th><%=strAnttimer %></th>
							<% name=CStr(OnlyDigits(rsTimeliste("Dato"))) %>
							<th><% =rsTimeliste("Timelonn") %>
							<th><% =rsTimeliste("FakturaPris") %>
							<th><% =rsTimeliste("FakturaStatus") %>
							<!------------------Link to DELETE-------------------->
							<th><A HREF=Vikar_timeliste_oppd_db2.asp?Linje=<%=rsTimeliste("Linje") %>&Slett=Ja&VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&StartDato=<% =strStartDato %>&FirmaID=<% =strFirmaID %>&kode=<% =status_kode %>&Dato=<%=DateValue(rsTimeliste("Dato"))%>&Bestilltav=<% =strBestilltav %> >Slett</A>

							<th><% =rsTimeliste("Vikarnavn") %>
							<th><% =rsTimeliste("Notat") %>

							<%
 						'If dag = "søndag" Or antLoop = antRecord Then 
					Else	
						neste = False
 					End If 
				End If	 'første dato er startDato, visning av uke
				' This part runs allways in the loop but collects weeks only before display date 
				If Not nowWeek = Datepart("ww",rsTimeliste("Dato"), 2) Then
					ukeTeller = ukeTeller + 1
					Redim Preserve ukeNrliste(ukeTeller)
					Redim Preserve ukeDatoListe(ukeTeller)
					ukeNrliste(ukeTeller - 1) = Datepart("ww",rsTimeliste("Dato"),2)
					ukeDatoListe(uketeller - 1) = rsTimeliste("Dato")
					nowWeek = Datepart("ww",rsTimeliste("Dato"),2)
				End If
				rsTimeliste.MoveNext
			loop 
			' End of loop 
			' Display buttons for navigating between weeks
			%>
		</table>
					<table cellpadding='0' cellspacing='0'>
						<tr><th>Uke:
						<% For i = 0 To ukeTeller - 1 
						If ukeNrliste(i) = Datepart("ww",strStartDato) Then %>
						<th><% =ukeNrListe(i) %></th>
						<% Else %>
						<FORM ACTION="Vikar_timeliste_oppd_vis2.asp?VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&FirmaID=<% =strFirmaID %>&startDato=<% =ukeDatoListe(i) %>&tilgang=<% =tilgang %>&kode=<% =status_kode %>&Bestilltav=<% =strBestilltav %>" METHOD=POST>
						<th><INPUT TYPE=SUBMIT VALUE=<% =ukeNrListe(i) %> ></th>
						<% End If %>
						</form>
						<% Next %>
							<th><th><th><th><th>
							<FORM ACTION="Vikar_timeliste_oppd_vis2.asp?Endre=Ny&VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&FirmaID=<% =strFirmaID %>&startDato=<% =ukeDatoListe(i) %>&tilgang=<% =tilgang %>&kode=<% =status_kode %>&Bestilltav=<% =strBestilltav %>" METHOD=POST >
							<th><INPUT TYPE=SUBMIT VALUE="  Ny  "></th>
							</form>
					</table>
					<%
					rsTimeliste.Close
				Else 	'no rows 
				session("sluttdato") = sluttdato 
				session("startdato") = strStartdato
				%>
				<table cellpadding='0' cellspacing='0'>
					<FORM ACTION=Vikar_timeliste_vis3.asp?VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&FirmaID=<% =strFirmaID %>&startDato=<% =ukeDatoListe(i) %>&tilgang=<% =tilgang %>&kode=<% =status_kode %>&Bestilltav=<% =strBestilltav %> METHOD=POST >
						<tr><th><INPUT TYPE=SUBMIT VALUE="              Tilbake             " ></th>
					</form>
				</table>
				<% End If %>
			</div>
		</div>
	</body>
</html>

