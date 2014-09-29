<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<!--#INCLUDE FILE="../includes/Xis.Settings.inc"-->
<!--#INCLUDE FILE="../includes/Xis.Economics.Constants.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	
		dim NedgradOK : NedgradOK = false
		dim strTimelonn : strTimelonn = ""
		dim strFakturaTimer : strFakturaTimer = ""
		dim strFakturapris : strFakturapris = ""
		dim strVikarID
		dim strOppdragID
		dim strLunch : strLunch = "00:00"
		dim strStarttid : strStarttid = ""
		dim strSluttid : strSluttid = ""
		dim strUkedag : strUkedag = ""
		dim strNotat : strNotat = ""
		dim strFaktAntallSteg1
		dim strFaktRateSteg1
		dim strFaktAntallSteg2
		dim strFaktRateSteg2
		dim Conn
		dim Ingenlinjer : Ingenlinjer = 0
		dim msg : msg = ""
		dim SOCuID
		dim strSQL
		dim rsCompanyInfo
		dim rsExtensions
		dim strFirmaID
		dim font
		dim fontEnd
		dim SOBestilltAv
		dim splittPart
		dim splittopt

		strVikarID = Request("VikarID")
		strOppdragID = Request("OppdragID")
		strLinje = Request("Linje")

		if strVikarID = "" and strOppdragID = "" then
			Response.Redirect "Vikar_timeliste_fakt_list.asp"
		end if
		
		if strVikarID = "" then
			strVikarID = "0"
		end if

		if strOppdragID = "" then
			strOppdragID = "0"
		else
			session("OppdragId") = Request("OppdragID")	
		end if

		If Request("FirmaID") <> "" Then
			session("FirmaId") = Request("FirmaID")
		End If
	
		Function OnlyDigits( strString )
		' Remove all non-nummeric signs from string
			If Not IsNull(strString) Then
				For idx = 1 To Len( strString) Step 1
					Digit = Asc( Mid( strString, idx, 1 ) )
					If (( Digit > 47 ) And ( Digit < 58 )) Then
						strNewstring = strNewString & Mid(strString, idx,1)
					End If
				Next
			End If
			Onlydigits = strNewString
		End Function
	
		'for å kompansere for feil ukenr i kalender..
		'FJM: This sub has sideeffects depending on value of fraverdi
		sub korrDatoer(inndato, fraverdi)
			'dy = inndato
			'call datoKorreksjon("d", dy)
			dy = datepart("d", inndato, 2, 2) 'day in year (1 - 357)

			'aarr = inndato
			'call datoKorreksjon("yyyy", aarr) 
			aarr = datepart("yyyy", inndato, 2, 2) 'get year

			uukkee = inndato

			uukkee = WeekFix(uukkee)

			if dy < 7 and uukkee and uukkee > 51 Then
				aarr = aarr - 1
			end if
			
			'Add leading zero if less than 10 (yyyyww)
			If uukkee < 10 then
				uukkee = "0" & uukkee
			end if

			if fraverdi = 1 Then 'update "global" weeknumber
				Session("displayWeek") = aarr & uukkee
			end if

			if fraverdi = 2 Then 'update local weeknumber
				ukenr = aarr & uukkee
			end if
		end sub

		Function GetTimeSheetEarlistDate()
			dim rsResult
			
			strSQL = "SELECT MIN(DV.Dato) AS [dato] " &_
				" FROM DAGSLISTE_VIKAR AS DV " &_
				" WHERE " &_
				" DV.VikarID = " & strVikarID &_
				" AND DV.OppdragID = " & strOppdragID &_
				" AND DV.TimelisteVikarStatus < 6 " &_
				" AND DV.Dato <= " & dbDate(limitDato)
		
			set rsResult = GetFirehoseRS(strSQL, Conn)
			if(HasRows(rsResult)) then
				GetTimeSheetEarlistDate = rsResult("dato")
				rsResult.close
			else
				GetTimeSheetEarlistDate = Now()
			end if
			set rsResult = nothing
		end function


		function RenderOvertimeDropDown(GroupKey, SelectedRate, dropDownName)
			dim returnValue : returnValue = ""
			dim selected
			dim lenSelectedRate
			
			if len(GroupKey) > 0 then
				lenSelectedRate = len(SelectedRate)
				strSQL = "SELECT [FakturaBeskrivelse], [FakturaSats], [IsDefault] " &_
					" FROM [H_FAKTURA_TYPE] " &_
					" WHERE [GroupKey] = '" & GroupKey & "' ORDER BY [FakturaSats] ASC"

				set rsOvertimeRates = GetFirehoseRS(strSQL, conn)
				Response.Write "<SELECT class='small' name='" & dropDownName & "' id='" & dropDownName & "'>"							
				WHILE NOT rsOvertimeRates.EOF
					selected = ""
					if lenSelectedRate > 0 then 'A selected rate was supplied
						if(SelectedRate = rsOvertimeRates("FakturaSats").value) then
							selected = "selected"
						end if
					elseif (rsOvertimeRates("IsDefault").value = true) then 'No rate select default 
						selected = "selected"
					end if
					Response.Write "<Option " & selected & " value='" & rsOvertimeRates("FakturaSats").value  &"'>" & rsOvertimeRates("FakturaBeskrivelse").value  & "</option>"				
					rsOvertimeRates.MoveNext
				WEND
				rsOvertimeRates.close
				set rsOvertimeRates = nothing
			end if
			RenderOvertimeDropDown = returnValue
		end function

		dim brukerID
		dim kkk
		dim p
	
		brukerID = Session("BrukerID")

		kkk = Request("Notat")
		p = Len(kkk) - 2
		If p > 1  Then
			kk = Mid(kkk, 2, p)
		End If

		' Get a database connection
		Set Conn = GetConnection(GetConnectionstring(XIS, ""))

		' prosessing parameters
		' Fixed Issue No - 51756 By TPH
		if len(trim(Request("TILDATOO"))) > 0 then
			Session("limitDato") = trim(Request("TILDATOO"))
		elseif Request.QueryString("dato1") <> "" then
			Session("limitDato") = Request.QueryString("dato1")
		elseif Request.QueryString("limitdato") <> "" then
			Session("limitDato") = Request.QueryString("limitdato")
		end if
		
		'if Request("TILDATOO") <> "" then 
		'	Session("limitDato")= Request("TILDATOO")
		'end if

		If Request("frakode") <> "" Then
			session("frakode") = Request("frakode")
		End If

		'gjør om "limitDato" til siste dag i uken
		'sub ligger i "library.inc"
		limitDato = Session("limitDato")
		Call newLimitDate(limitDato)

		strEndre = Request("Endre")
		strNy = Request("Ny")

		gmlEnr = Request("gmlEnr")
		splittopt = Request("splittopt2") 'fra denne siden
'Response.Write splittopt
		If splittopt = ".1" Then
			splittopt = 1
		ElseIf splittopt = ".2" Then
			splittopt = 2
		Else
			splittopt = "NULL"
		End If

		If Request("splittopt") <> "" Then  'fra andre sider
			splittopt = Request("splittopt")
		End If
		
		session("splittopt") = splittopt
		session("splittuke") = Request("splittuke")

		'Sets week to display		
		If lenb(Request("startDato")) > 0 Then
 			strStartDato = cDate(Request("startDato"))
		else
			strStartDato = GetTimeSheetEarlistDate()
		End If
		call korrDatoer(strStartDato, 1) 


		' SQL for å sjekke om det er ok å godkjenne nye uker
		' nye uker kan ikke godkjennes hvis det finne linjer med faktura- eller lønnsstatus 2
		Dim godkjennOK

		strSQL = " SELECT DISTINCT [VikarID] FROM [DAGSLISTE_VIKAR]" &_
			" WHERE ([fakturastatus] = 2 or [Loennstatus] = 2)" &_
			" AND [oppdragid] = " & strOppdragID & " AND [vikarid] = " & strVikarID

		Set rsGodkjennTest = GetFirehoseRS(strSQL, Conn)
		if HasRows(rsGodkjennTest) THEN
			godkjennOK = 1
			rsGodkjennTest.Close
		else
			godkjennOK = -1
		end if
		
		Set rsGodkjennTest = nothing
				

		' SQL for finding name
		' Update SQL 27.11.2001 E.L
		strSQL = "SELECT " & _
			"Navn = (VIKAR.Fornavn + ' ' + VIKAR.Etternavn), " & _
			"VIKAR.TypeID, " & _
			"VIKAR_ANSATTNUMMER.ansattnummer " & _
			"FROM VIKAR " & _
			"LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " & _
			"WHERE VIKAR.Vikarid = '" & strVikarID & "' "

		Set rsNavn = GetFirehoseRS(strSQL, Conn)
		If (not HasRows(rsNavn)) Then
			Set rsNavn = Nothing
			AddErrorMessage("Vikaren ble ikke funnet!")
			call RenderErrorMessage()
		else
			strAnsattnummer = rsNavn("ansattnummer").Value
			strNavn = rsNavn("Navn").Value
			strVikarType = rsNavn("TypeID").Value	
		End If
		rsNavn.Close
		Set rsNavn = Nothing

		' SQL for displaying data
		'frakode = 1 (-1), personalkonsulent
		'frakode > 2, økonomi

		'fra personalkonsulent
		If session("frakode") = 1 Or CInt(session("frakode")) = -1 Then
			strSQL = "SELECT Linje = DV.TimelisteVikarID, DV.VikarID, DV.Dato, status = DV.TimelisteVikarStatus, " &_
				"DV.Starttid, DV.sluttid, DV.AntTimer, F.FirmaID, F.firma, DV.Timelonn, DV.Fakturastatus, " &_
				"DV.BestilltAv, DV.SOBestilltAv, DV.Fakturapris, DV.Notat, DV.Fakturatimer, DV.Loennstatus, DV.Lunch, DV.LoennsArt, DV.splittuke " &_
				" FROM DAGSLISTE_VIKAR AS DV, Firma AS F " &_
				" WHERE DV.Firmaid = F.firmaid " & _
				" AND DV.VikarID = " & strVikarID &_
				" AND DV.OppdragID = " & strOppdragID &_
				" AND DV.TimelisteVikarStatus < 6 " &_
				" ORDER BY DV.Dato"			
		
		'fra økonomi
		elseIf session("frakode") > 1 Then

			If gmlEnr <> "" Then
				sstat = 7
			Else
				sstat = 6
			End If

			strSQL = "SELECT Linje = DV.TimelisteVikarID, DV.VikarID, DV.Dato, status = DV.TimelisteVikarStatus, " &_
				"DV.Starttid, DV.sluttid, DV.AntTimer, F.FirmaID, F.firma, DV.Timelonn, DV.Fakturastatus, " &_
				"DV.BestilltAv, DV.SOBestilltAv, DV.Fakturapris, DV.Notat, DV.Fakturatimer, DV.Loennstatus, DV.Lunch, DV.LoennsArt, DV.splittuke " &_
				" FROM DAGSLISTE_VIKAR AS DV, Firma AS F " &_
				" WHERE DV.FirmaID = F.FirmaID " &_
				" AND DV.VikarID = " & strVikarID &_
				" AND DV.OppdragID = " & strOppdragID &_
				" AND DV.TimelisteVikarStatus < " & sstat &_
				" AND DV.Dato <= " & dbDate(limitDato) &_
				" ORDER BY DV.Dato"
		End If

		Set rsTimeliste = GetFirehoseRS(strSQL, Conn)

		' If no record exsists
		If (rsTimeliste.EOF) Then
			Response.Write "Har ikke timeliste! <br>"
			Ingenlinjer = 1
		Else
			Ingenlinjer = 0

			' SQL for firma name
			ID = rsTimeliste("FirmaID")
			strFirma = rsTimeliste("Firma")
			' Find week number
			If ((strStartDato = "") or (IsNull(strStartDato))) Then
				strStartDato = rsTimeliste("Dato")
			End If

			strWeekNumber = Datepart("yyyy", strStartDato, 2, 2) & Datepart("ww", strStartDato, 2, 2)
			' MK: Workaround if the page is opened for the first time and Session variable DisplayWeek has not been set
		    Session("DisplayWeek_temp") = strWeekNumber

			strBestilltAv = rsTimeliste("BestilltAv")
			strSOBestilltAv = rsTimeliste("SOBestilltAv")
			strFirmaID = rsTimeliste("FirmaID")

			' Form to register and change
			' If edit find the right row to display
			If strLinje <> "" Then ' edit
		
				strSQL = "SELECT Dato, Starttid, Sluttid, AntTimer, Timelonn, " &_
					"Fakturapris, Fakturatimer, Lunch, Notat, LoennsArt, splittuke, " &_
					"Fakturastatus, loennstatus " &_
					"from DAGSLISTE_VIKAR " &_
					"where TimelisteVikarID = " & strLinje

				Set rsLinje = GetFirehoseRS(strSQL, Conn)

				strDato2 = DateValue(rsLinje("Dato"))
				strStarttid = Left(TimeValue(rsLinje("Starttid")), 5)
				strSluttid = Left(TimeValue(rsLinje("Sluttid")), 5)
				strUkedag = WeekDayName(Weekday(rsLinje("Dato"), 2),, 2)
				strAntTimer = rsLInje("AntTimer")
				strTimelonn = rsLinje("Timelonn")
				strFakturapris = rsLinje("Fakturapris")
				strLunch = Left(TimeValue(rsLinje("Lunch")), 5)
				strFakturatimer = rsLinje("Fakturatimer")
				strNotat = rsLinje("Notat")
				strLoennsArt = rsLinje("LoennsArt")
				strSplitt = rsLinje("splittuke")
				strFakturastatus = rsLinje("Fakturastatus")
				strLoennstatus = rsLinje("Loennstatus")

				rsLinje.Close
				set rsLinje = nothing

			End If	
			%>
			<html>
				<head>
					<title>Timeliste</title>
					<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
					<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
					<script type="text/javascript" language="javascript" src="../js/javaScript.js"></script>
					<script language="javaScript" type="text/javascript">
						function settUkesplitt(adr)
						{
							location = adr;
						}

						function hentFakttimer(p)
						{
							bb = document.all('tbxTimerPrdag').value;
							document.all('tbxFakttimerPrDag').value = bb
						}

						function hentFakttimer2(fnr)
						{
							var bb;
							bb = document.forms[fnr].elements('tbxTimerPrdag').value;
							document.forms[fnr].elements('tbxFakttimerPrDag').value = bb;
						}

						function hentVanligetimer() 
						{
							k = document.all('SUMM2').value;
							p = k.indexOf(",");

							if (p != -1) 
							{
								kk = k.substring(0, p) + "." + k.substring(p + 1);
							}
							else
							{
								kk = k;
							}

							ov1 = document.all('PR50').value;
							if (ov1 == 0) 
							{
								ov11 = 0;
							} 
							else 
							{
								p = ov1.indexOf(",");
								if (p != -1) 
								{
									ov11 = ov1.substring(0, p)+ "." + ov1.substring(p+1);
								}
								else
								{
									ov11 = ov1;
								}
							}
				
							ov2 = document.all('PR100').value;
							if (ov2 == 0) 
							{
								ov22 = 0;
							}
							else 
							{
								p = ov2.indexOf(",");
								if (p != -1) 
								{
									ov22 = ov2.substring(0, p)+ "." + ov2.substring(p+1);
								}
								else
								{
									ov22 = ov2;
								}
							}
							sum1 = parseFloat(kk) - (parseFloat(ov11) + parseFloat(ov22));
							document.all('SUMM3').value = sum1
							document.all('SUMM').value = sum1
							document.all('PRFAKT50').value = ov1
							document.all('PRFAKT100').value = ov2
						}


						function hentVanligetimer2() 
						{
							k = document.all('FAKTSUMM2').value;
							p = k.indexOf(",");

							if (p != -1) 
							{
								kk = k.substring(0, p)+ "." + k.substring(p + 1);
							}
							else
							{
								kk = k;
							}

							ov1 = document.all('PRFAKT50').value;
							if (ov1 == 0) 
							{
								ov11 = 0;
							} 
							else 
							{
								p = ov1.indexOf(",");
								if (p != -1) 
								{
									ov11 = ov1.substring(0, p)+ "." + ov1.substring(p+1);
								}
								else
								{
									ov11 = ov1;
								}
							}
							
							ov2 = document.all('PRFAKT100').value;
							if (ov2 == 0) {
								ov22 = 0;
							} else {
								p = ov2.indexOf(",");
								if (p != -1) {
									ov22 = ov2.substring(0, p)+ "." + ov2.substring(p+1);
								}else{
									ov22 = ov2;
								}
							}
							sum1 = parseFloat(kk) - (parseFloat(ov11) + parseFloat(ov22));

							document.all('FAKTSUMM3').value = sum1
							document.all('FAKTSUMM').value = sum1
						}
						
						function ikkeFakturerOmSyk(strFakturaTimer)
						{
						    if(document.all('SYK').value == 'EM' || document.all('SYK').value == 'HD' || document.all('SYK').value == 'SB' || document.all('SYK').value == 'SM')
						    {
						        document.all('tbxFaktTimerPrDag').value = 0;
						    }
						    //else document.all('tbxFaktTimerPrDag').value = strFakturaTimer;
						}
					</SCRIPT>
				</head>
					<body>
						<div class="pageContainer" id="pageContainer">		
							<div class="contentHead1">
								<h1>Timelister</h1>
							</div>
							<div class="content">
								
								<h2>Vikar:&#160;<%=strAnsattnummer%>&#160;<%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "/xtra/VikarVis.asp?VikarID=" & strVikarID, strNavn, "Vis vikar " & strNavn )%><br>
								<%
								if (len(strFirmaID) > 0) then
									strSQL = "SELECT [SOCuID] FROM [Firma] WHERE [firmaId]=" & strFirmaID
									set rsCompanyInfo = GetFirehoseRS(StrSQL, Conn)
									if not rsCompanyInfo.EOF Then
										SOCuID = rsCompanyInfo("SOCuID").value
									end if
									rsCompanyInfo.close
									set rsCompanyInfo = nothing
									%>
									Kontakt:&#160;<%=CreateSONavigationLink(SUPEROFFICE_PANEL_CONTACT_URL, SUPEROFFICE_PANEL_CONTACT_URL, SOCuID, strFirma, "Vis kontakt '" & strFirma & "'")%><br>
									<%
								end if
								%>
								Oppdrag:&#160;<%=CreateSOLink(SUPEROFFICE_XIS_TASK_URL, "", "/xtra/WebUI/OppdragView.aspx?OppdragID=" & strOppdragID, strOppdragID, "Vis oppdrag " & strOppdragID )%></h2>
							</div>
							<%
							'Beskjed til økonomi og beskjed fra vikar
							If session("frakode") < 10 Then							
								%>
								<div class="contentHead"><h2>Informasjon fra vikar / ansvarlig</h2></div>
								<div class="content">
									<%
									dim strMessage
									strMessage = "<p><strong>Ingen melding fra vikar.</strong></p>" 'Default no message
									
									dim tempWeekNumber
									if(IsNull(Session("DisplayWeek"))) then
									    tempWeekNumber = Session("DisplayWeek_temp")
									else
									    tempWeekNumber = Session("DisplayWeek")
									end if
									
									StrSQL = "SELECT [kommentar] " & _
									"FROM [UKE_KOMMENTAR] " & _
									"WHERE [oppdragid] = " & strOppdragID & _
									" AND [vikarid] = " & strVikarID & _
									" AND [ukenr] = " & tempWeekNumber '& Session("DisplayWeek")
									
									set tempWeekNumber = nothing


									Set rsKommentar = GetFirehoseRS(StrSQL, Conn)
									if (HasRows(rsKommentar)) THEN
										if(lenb(trim(rsKommentar("kommentar").value))) > 2 then 'This value contains ascii char 13 when "empty"
											strMessage = "<p><strong>Beskjed fra vikar:</strong>&#160;" & rsKommentar("kommentar").value & "</p>"
										end if
										rsKommentar.Close
									end if
									set rsKommentar = nothing
									Response.Write strMessage
									
									strSQL = "SELECT [NotatOkonomi] FROM [oppdrag] WHERE [oppdragID] = " & session("OppdragId") & " AND len(NotatOkonomi) > 0 " 
									Set rsMelding = GetFirehoseRS(strSQL, Conn)
									If HasRows(rsMelding) THEN
										strMessage = "<p><strong>Beskjed til økonomi om oppdraget:</strong>&#160;" & rsMelding("NotatOkonomi") & "</p>"
										rsMelding.close
									else
										strMessage = "<p><strong>Ingen melding fra ansvarlig.</strong></p>"
									end if
									set rsMelding = nothing								
									Response.Write strMessage
									%>
								</div>							
								<%
							end if
							' Form to register and change
							If strEndre = "Ja" Or strNy = "Ja" Then
							
								kk = strStartDato
								pp = Request("Dato")

								If kk = pp Then
									kk = ""
								End If

								If request("Ny")= "Ja" Then 
									splittopt = request("ukedel")
								end if
							
								if splittopt="1" OR splittopt="2" then 
									strSplitt = splittopt
								end if
								%>
								<div class="contentHead"><h2>Dagsliste</h2></div>
								<div class="content">
									<form name="EDITER" action="Vikar_timeliste_db3.asp?kode=1&VikarID=<%=strVikarID%>&OppdragID=<% =strOppdragID %>&Linje=<% =strLinje %>" method=POST id="Form1">
										<input name="FirmaID" type="HIDDEN" value="<% =strFirmaID %>" id="Hidden1">
										<input name="StartDato" type="HIDDEN" value="<% =kk %>" id="Hidden2">
										<input name="Endre" type="HIDDEN" value="<% =strEndre %>" id="Hidden3">
										<input name="Endret" type="HIDDEN" value="<% =strSplitt %>" id="Hidden4">
										<input name="BestilltAv" type="HIDDEN" value="<% =strBestilltAv %>" id="Hidden5">
										<input name="SOBestilltAv" type="HIDDEN" value="<% =strSOBestilltAv%>" id="Hidden6">
										<input name="tbxTimerPrDag" type="HIDDEN" value="<% =strAnttimer %>" id="Hidden7">
										<input name="limitDato" type="HIDDEN"  value="<% =limitDato %>" id="Hidden8">
										<input name="splittopt2" type="HIDDEN" value="<% =splittopt %>" id="Hidden9">
										<div class="listing">
											<table id="Table1">
												<tr>
													<th>Ukedag</th>
													<th>Dato</th>
													<th>Fra kl</th>
													<th>Til kl</th>
													<th>Lunsj</th>
													<th>Timelønn</th>
													<th>F.timer</th>
													<th>F.pris</th>
												</tr>
												<tr>
													<td><%=strUkedag%></th>
													<td><input type="text" size="8" class="right" maxlength="8" name="Dato" value="<% If strLinje <> "" Then response.write strDato2  %>" onblur="dateCheck(this.form, this.name)" id="Text1"></td>
													<td><input type="text" size="5" class="right" maxlength="5" name="tbxFraKl" value="<%=strStarttid%>" onchange="timeCheck(this.form, this.name), workTime(this.form, this.name), hentFakttimer()" id="Text2"></td>
													<td><input type="text" size="5" class="right" maxlength="5" name="tbxTilKl" value="<%=strSluttid%>" onchange="timeCheck(this.form, this.name), workTime(this.form, this.name), hentFakttimer()" id="Text3"></td>
													<td><input type="text" size="5" class="right" maxlength="5" name="tbxLunsj" value="<%=strLunch %>" onchange="timeCheck(this.form, this.name), workTime(this.form, this.name), hentFakttimer()" id="Text4"></td>
													<td><input type="text" size="4" class="right" name="Timelonn" value="<%=strTimelonn%>" id="Text5"></td>
													<td><input type="text" size="5" class="right" maxlength="5" name="tbxFaktTimerPrDag" value="<%=strFakturaTimer%>" id="Text6"></td>
													<td><input type="text" size="4" class="right" name="Fakturapris" value="<%=strFakturapris%>" id="Text7"></td>
												</tr>
												<tr>
													<th>Notat:</th>
													<td colspan="4"><input type="text" size="36" name="Notat" value='<%=strNotat%>' id="Text8"></td>
													<th>Lønnsart:</th>
													<td colspan="2">
														<select name="SYK" id="Select1" onchange="ikkeFakturerOmSyk(<%=strFakturaTimer%>);">
														<%
														strSQL = "SELECT loennsart, loennsartKode FROM H_LOENNSART WHERE LoennsartKode is not null ORDER BY LoennsartKode"
														Set rsPaymentCodes = GetFirehoseRS(strSQL, Conn)
														While not rsPaymentCodes.EOF
															if (strLoennsart = rsPaymentCodes("loennsartKode").value) then
																sel = "selected"
															else
																sel = ""
															end if
															%>
															<option value="<%=rsPaymentCodes("loennsartKode").value%>" <%=sel%>><%=rsPaymentCodes("loennsart").value%></option>
															<%
															rsPaymentCodes.MoveNext
														wend
														rsPaymentCodes.close
														set rsPaymentCodes = nothing														
														%>
														</select>
													</td>
												</tr>
												
												<tr>
													<th colspan="6">
														<input type="SUBMIT" value="Lagre" id="Submit1" name="Submit1">
														<input type="RESET" value="Tilbakestill" id="Reset1" name="Reset1">
													</th>
													<th colspan="2">Splitt uke:
														<% 
														If strSplitt = 1 or strSplitt = 2 Then
															opt = "splitt checked disabled"
														Else
															opt = "splitt"
														end if														
														%>
														<input type="CHECKBOX" name="splittuke" class="checkbox" value=<%=opt%> id="Checkbox1">
													</th>
												</tr>
											</table>
										</div>
									</form>
								</div>
								<%
							End If  ' Endre = ja
							' Display data
							%>
							<div class="contentHead"><h2>Ukeliste</h2></div>
							<div class="content">
								<input type="hidden" size="3" name='ukedel' value=<% =splittopt %> id="Hidden10">
								<div class="listing">
									<table id="Table2">
										<tr>
											<th>&#160;</th>
											<th>Ukedag</th>
											<th>Dato</th>
											<th>Start kl</th>
											<th>Slutt kl</th>
											<th>Timer</th>
											<th>Lunsj</th>
											<th>Timelønn</th>
											<th>F.timer</th>
											<th>F.gr.</th>
											<th>L.status</th>
											<th>F.status</th>
											<th>T.status</th>
											<th>L.art</th>
											<th>Notat</th>
											<th>Slette?</th>
										</tr>
										<%
										' Prepare for Loop
										antRecord = 0
										antLoop = 0
										neste = False
						
										If Request("TLONN") = "" Then
											TLONN = rsTimeliste("Timelonn")
										Else
											TLONN = Request("TLONN")
										End If
										
										If LenB(Request("startDato")) > 0 Then
 											strStartDato = cDate(Request("startDato"))
										Else
											strStartDato = rsTimeliste("Dato")
										End If

										call korrDatoer(strStartDato, 1)										

										If Request("TLONN") = "" Then
											TLONN = rsTimeliste("Timelonn")
										Else
											TLONN = Request("TLONN")
										End If

										Session("begDato") = rsTimeliste("Dato")
										dagteller = 0
										session("dagTeller") = 0
										flereDg = -1
										ukedel_1 = -1
										ukedel_2 = -1
										ukeTeller = 0
										nowWeek =  0
										strSum = 0
										strFaktSum = 0
										setVar = True
										tt = 0
										splitt = 0
										strStatus2 = 6 'fanger opp den minste timelistestatusen denne uken

										' Get number of records
										do while not rsTimeliste.EOF
											antRecord = antRecord + 1
											rstimeliste.MoveNext
										loop
										rsTimeliste.MoveFirst

										do while not rsTimeliste.EOF
											antLoop = antLoop + 1

											'for å kompansere for feil ukenr i kalender..
											ukenr = ""
											call korrDatoer(rsTimeliste("Dato"), 2)

											If ukenr = Session("DisplayWeek") Then
												If Request("splittuke") = "Ja" And Request("startdato")<> "" Then
    												If Datevalue(Request("startdato")) = Datevalue(rsTimeliste("Dato")) Then
														neste = True
    												End If
												Else
													neste = True
												End If
												displayWeek  = ukenr
											End If
											dag = WeekDayName(Weekday(rsTimeliste("Dato"), 2), , 2)

											' Determine when to display rows (weeks before are collected at the bottom of loop)
											' Display until sunday else (create buttons for stored weeks ) and create buttons for next weeks
											If neste Then
												tt = tt + 1
												If tt > 1 And rsTimeliste("splittuke") = 2 Then  'ikke slå inn når 'splitt' er første dagen..
													splitt = -1
												End If
												If (displayWeek = ukenr And Not splitt) or (displayWeek = ukenr And (request("splittopt2")=".1" or request("splittopt2")="1") and rsTimeliste("splittuke") = 1 ) or (displayWeek = ukenr And (request("splittopt2")=".2" or request("splittopt2")="2") and rsTimeliste("splittuke") = 2 ) then
													If setVar Then
 		 												strStartDato = rsTimeliste("Dato")
														setVar = False
													End If
													strStarttid = Left(TimeValue(rsTimeliste("Starttid")),5)
													strSluttid = Left(TimeValue(rsTimeliste("Sluttid")),5)
													strAnttimer = FormatNumber(rsTimeliste("AntTimer"),2,-1)
													'pos1 = Instr(1, strAnttimer,",")
													'If pos1 > 0 Then
													'	strAnttimer = Left(strAntTimer, pos1 + 2)
													'End If
													strLunch = Left(TimeValue(rsTimeliste("Lunch")),5)
													strTimelonn = rsTimeliste("Timelonn")

													strFakturatimer = FormatNumber(rsTimeliste("Fakturatimer"),2,-1)
													'pos1 = Instr(1, strFakturatimer, ",")
													'If pos1 > 0 Then
													'	strFakturatimer = Left(strFakturaTimer, pos1 + 2)
													'End If
													strFakturapris = rsTimeliste("Fakturapris")
													strSum = strSum + strAnttimer
													strFaktSum = strFaktSum + strFakturatimer
     												sluttdato = rsTimeliste("Dato")
													fStatus = rsTimeliste("Fakturastatus")
													lStatus = rsTimeliste("Loennstatus")
													tStatus = rsTimeliste("Status")
													'Tilat kun nedgradering dersom faktura- og lønnstatus er 1, og timeliste status er 5.
													if cInt(fStatus) = 1 and cInt(lStatus) = 1 and cInt(tStatus) = 5 Then 
														NedgradOK = true 
													end if
													%>
													<tr>
														<% 
														If tStatus < 5 Or gmlEnr <> "" Then 
															update="Endre=Ja&VikarID=" & strVikarID & "&OppdragID=" & strOppdragID & "&Linje=" & rsTimeliste("Linje") &_
                											"&StartDato=" & strStartDato & "&splittuke=" & session("splittuke") & "&splittopt2=" & splittopt & "&splittopt=" & splittopt
															%>
															<th><a href="Vikar_timeliste_vis3.asp?<%=update%>">Endre</a></th>
															<% 
														Else 
															%>
															<td>&nbsp;</td>
															<% 
														End If 
														%>
														<td><%=dag%></td>
														<td><%=DateValue(rsTimeliste("Dato"))%></td>
														<td class="right"><%=strStarttid %></td>
														<td class="right"><%=strSluttid %></td>
														<td class="right"><%=strAnttimer %></td>
														<td class="center">
															<%
															Response.Write strLunch
															name = CStr(OnlyDigits(rsTimeliste("Dato"))) 
															%>
														</td>
														<td class="right">
															<% 
															If session("frakode") < 10 Then
																Response.Write rsTimeliste("Timelonn")
															End If 
															%>
														</td>
														<td class="right"><%=strFakturatimer%></td>
														<td class="right">
															<% 
															If session("frakode") < 10 Then 
																Response.Write  rsTimeliste("FakturaPris")
															End If 
															%>
														</td>
														<td class="center"><% =rsTimeliste("Loennstatus") %></td>
														<td class="center"><% =rsTimeliste("FakturaStatus") %></td>
														<%
														If tStatus < strStatus2 Then 
															strStatus2 = tStatus
														end if
														%>
														<td class="center"><% =tStatus %></td>
														<td><% =rsTimeliste("LoennsArt")%>&nbsp;</td>
														<td><% =rsTimeliste("Notat") %>&nbsp;</td>											
														<% 
														If rsTimeliste("status") < 5 Then
															dagteller = dagteller + 1
															session("dagTeller") = dagteller	
															'Link to DELETE
															delete = "Slett=Ja&VikarID=" & strVikarID & "&OppdragID=" & strOppdragID & "&Linje=" & rsTimeliste("Linje") &_
                       										"&StartDato=" & strStartDato & "&splittuke=" & session("splittuke") & "&splittopt2=" & splittopt
															%>
															<td><a href=Vikar_timeliste_db3.asp?<% =delete %>&kode=1&sletteDato=<%=DateValue(rsTimeliste("Dato"))%> >Slett</a></td>
															<% 
														else
															%>
															<td>&nbsp;</td>
															<%
														End If 
														%>
													</tr>
													<%
												Else
													neste = False
 												End If
											End If	 'første dato er startDato, visning av uke (neste)

											' This part runs allways in the loop but collects weeks before and after display week
											If rsTimeliste("splittuke") = 1 Then
												splittuke = 1
											ElseIf rsTimeliste("splittuke") = 2 Then
												splittuke = 2
											Else
												splittuke = 0
											End If

											if nowWeek <> ukenr then
												flereDg = -1
												ukedel_1 = -1
												ukedel_2 = -1
											end if

											If Not nowWeek = ukenr Or splittuke <> 0  Then
												if flereDg or (ukedel_1 or ukedel_2 AND not splittuke = gmlUkedel) then
													gmlUkedel = splittuke
													If rsTimeListe("Fakturastatus") = 3 And rstimeliste("Loennstatus") = 3 Then
														fff = "<table bgcolor='black'>"
													ElseIf rsTimeliste("status") = 5 Then
														fff = "<table bgcolor='green'>"
													ElseIF rsTimeliste("status") = 4 Then
														fff = "<table bgcolor='yellow'>"
													ElseIF rsTimeliste("status") = 3 Then
														fff = "<table bgcolor='cyan'>"
													ElseIF rsTimeliste("status") = 2 Then
														fff = "<table bgcolor='blue'>"
													Else
														fff = "<table bgcolor='red'>"
													End If

													ukeTeller = ukeTeller + 1
													Redim Preserve ukeNrliste(ukeTeller)
													Redim Preserve ukeDatoListe(ukeTeller)
													Redim Preserve ukeFontListe(ukeTeller)
													Redim Preserve ukeSplittliste(ukeTeller)

													If splittuke = 1 And ukedel_1 Then
														ukeNrliste(ukeTeller - 1) = ukenr & ".1"
														ukedel_1 = 0
														flereDg = 0
													ElseIf splittuke = 2 And ukedel_2 Then
														ukeNrliste(ukeTeller - 1) = ukenr & ".2"
														ukedel_2 = 0
														flereDg = 0
													Else
														ukeNrliste(ukeTeller - 1) = ukenr
													End If

													ukeDatoListe(uketeller - 1) = rsTimeliste("Dato")
													ukeFontListe(ukeTeller - 1) = fff

													If splittuke Then
														ukeSplittliste(ukeTeller - 1) = "Ja"
													End If

													nowWeek = ukenr
												end if 'flereDg
											End if
											' End of loop
											rsTimeliste.MoveNext
										loop
										If strStatus2 < 5 Then 
											%>
											<tr>
												<form name="aaa" action="Vikar_timeliste_vis3.asp?Ny=Ja&VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&FirmaID=<% =strFirmaID %>&TimeLonn=<% =strTimelonn %>&UkeNr=<% =session("displayweek") %>&Startdato=<% =strStartdato %>&limitDato=<% =limitDato %>&ukedel=<% =splittopt %>&splittopt2=.<%=splittopt%>&splittuke=<%=session("splittuke")%>" method="post" id="Form2">
													<th colspan="16"><input type=SUBMIT value=" Ny " id="Submit2" name="Submit2"></th>
												</form>										
											</tr>
											<%
										End If 
										%>
									</table>
									<table id="Table3">
										<%
										' SQL to find overtime
										If lcase(splittopt) = "null" Then
											notat = "IS NULL Or Notat = ' ' Or Notat like 'NULL'"
										Else
											notat = "= '" & splittopt & "'"
										End If

										'Vanlige lønnede timer
										strSQL = "SELECT Antall = Sum(Antall) " & _
											" FROM VIKAR_UKELISTE, H_LOENNSART  " & _
											" WHERE VIKAR_UKELISTE.Loennsartnr = H_LOENNSART.Loennsartnr  " &_
											" AND Ukenr = " & session("displayWeek") &_
											" AND VikarID = " & strVikarID &_
											" AND OppdragID = " & strOppdragID &_
											" AND LoennRate = '1.0' " &_ 
											" AND (Notat " & notat & ")"

										Set rsLn = GetFirehoseRS(strSQL, Conn)
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
											" AND Ukenr = " & session("displayWeek") &_
											" AND VikarID = " & strVikarID &_
											" AND OppdragID = " & strOppdragID &_
											" AND Faktureres = 1 " &_ 
											" AND fakturaSats = '1.0' " &_	
											" AND (Notat " & notat & ")"							

										Set rsLn = GetFirehoseRS(strSQL, conn)
										If Not rsLn.EOF Then
											strFaktVanlig = rsLn("FakturertAntall")
										Else
											strFaktVanlig = ""
										End If
										rsLn.close
										set rsLn = nothing

										'Lønn overtid steg 1 timer
										strSQL = "SELECT Antall=SUM(Antall) " & _
											" FROM VIKAR_UKELISTE, H_LOENNSART " & _
											" WHERE VIKAR_UKELISTE.Loennsartnr  = H_LOENNSART.Loennsartnr " &_
											" AND Ukenr = " & session("displayWeek") &_
											" AND VikarID = " & strVikarID &_
											" AND OppdragID = " & strOppdragID &_
											" AND LoennRate = '1.5' " &_ 
											" AND (Notat " & notat & ")"
										
										Set rsLn = GetFirehoseRS(strSQL, Conn)
										If Not rsLn.EOF Then
											str50 = rsLn("Antall")
											If str50 = 0 Then 
												str50 = ""
											end if
										End If
										rsLn.close
										set rsLn = nothing
										
										'Fakturerbar overtid timer
										strSQL = "SELECT FakturertAntall = Sum(Antall), fakturaSats, GroupKey " & _
											" FROM VIKAR_UKELISTE, H_LOENNSART, H_FAKTURA_TYPE  " & _
											" WHERE VIKAR_UKELISTE.Loennsartnr = H_LOENNSART.Loennsartnr  " &_
											" AND H_FAKTURA_TYPE.FAKTURATYPE = VIKAR_UKELISTE.FAKTURATYPE  " &_
											" AND Ukenr = " & session("displayWeek") &_
											" AND VikarID = " & strVikarID &_
											" AND OppdragID = " & strOppdragID &_
											" AND Faktureres = 1 " &_ 
											" AND fakturaSats > '1.0'  " &_	
											" AND (Notat " & notat & ")" &_
											" GROUP BY fakturaSats, GroupKey " &_
											" Order by GroupKey, fakturaSats ASC "
'response.Write "strSQL:" & strSQL & "<br><br>"
										Set rsLn = GetFirehoseRS(strSQL, Conn)
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
										strSQL = "SELECT Antall=Sum(Antall) " & _
											" FROM VIKAR_UKELISTE, H_LOENNSART " & _
											" WHERE VIKAR_UKELISTE.Loennsartnr  = H_LOENNSART.Loennsartnr " &_
											" AND Ukenr = " & session("displayWeek") &_
											" AND VikarID = " & strVikarID &_
											" AND OppdragID = " & strOppdragID &_
											" AND LoennRate = '2.0' " &_ 
											" AND (Notat " & notat & ")"

										Set rsLn = GetFirehoseRS(strSQL, Conn)
										If Not rsLn.EOF Then
											str100 = rsLn("Antall")
											If str100 = 0 Then 
												str100 = ""
											end if
										End If
										rsLn.close
										set rsLn = nothing										

										session("EndDato") = sluttdato
										' Display save and new button
										%>
										<tr>
											<form name="TLISTE" action="Vikar_timeliste_lagre3.asp?Lagringskode=oppgrad&Startdato=<% =strStartdato %>&Sluttdato=<% =Sluttdato %>&VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&Timelonn=<% =strTimelonn %>&FirmaID=<% =strFirmaID %>&UkeNr=<% =session("displayweek") %>&limitDato=<% =limitDato %>&BestilltAV=<% =strBestilltAv %>&SOBestilltAV=<% =strSOBestilltAv %>&Fakturapris=<% =strFakturapris %>&kode=1&splittopt=<%=splittopt%>" method="POST" id="Form3">
												<%
												If strVanlig = "" Or IsNull(strVanlig) Then
													sss = strSum
													strVanlig = strSum
												Else
													sss = strVanlig
												End If
												%>
												<input type="hidden" name="SUMM" value=<% =sss %> id="Hidden11">
												<input type="hidden" name="SUMM2" value=<% =strSum %> id="Hidden12">
												<%
												If strFaktVanlig = "" Or IsNull(strFaktVanlig) Then
													sss = strFaktSum
													strFaktVanlig = strFaktSum
												Else
													sss = strFaktVanlig
												End If
												%>
												<input type="hidden" name="FAKTSUMM" value="<% =sss %>" id="Hidden13">
												<input type="hidden" name="FAKTSUMM2" value="<% =strFaktSum %>" id="Hidden14">
												<% 
												'viser samlesummer i tabell
												%>
												<td>Registrerte timer:</td>
												<th class="right"><% =FormatNumber(strSum, 2) %></th>
												<td>Timer til fakturering:</td>
												<th class="right"><% =FormatNumber(strFaktSum, 2) %></th>
												<td>Timer utover ordinær arbeidstid:</td>
												<% 
												sumsum = strSum - XIS_WORKHOURS_PER_WEEK
												If sumsum > 0 Then  
													%>
													<th class="right"><% =FormatNumber(Left(sumsum, 5), 2) %></th>
													<% 
												Else 
													%>
													<th>&nbsp;</th>
													<% 
												End If 
												%>
											</tr>
											<% 			
'Response.Write "strStatus2:" & strStatus2 & "<br>"
'Response.Write "godkjennOK:" & godkjennOK & "<br>"
											If strStatus2 < 5  and godkjennOK = -1 Then 
												%>
												<tr>
													<td>Ordinære lønnede timer:</td>
													<th class="right"><input name="SUMM3" type="TEXT" size="4" value="<% =strVanlig %>" readonly id="Text9"></th>
													<td>Ordinære fakturerte timer:</td>
													<th class="right"><input name=FAKTSUMM3 type="text" size=4 value="<% =strFaktVanlig %>" readonly id="Text10"></th>
												</tr>
												<tr>
													<td>Lønnede timer 50%:</td>
													<th class="right"><input name=PR50 type="text" size=4 value="<% =str50 %>" onblur="hentVanligetimer(), hentVanligetimer2()" id="Text11"></th>
													<td>Fakturerte timer <%=RenderOvertimeDropDown("step1", strFaktRateSteg1, "lstRateSteg1")%></td>
													<th class="right"><input name=PRFAKT50 type="text" size=4 value="<% =strFaktAntallSteg1 %>" onblur=hentVanligetimer2() id="Text13"></th>
												</tr>
												<tr>
													<td>Lønnede timer 100%:</td>
													<th class="right"><input name=PR100 type="text" size=4 value="<% =str100 %>" onblur="hentVanligetimer(), hentVanligetimer2()" id="Text12"></th>
							      					<td>Fakturerte timer <%=RenderOvertimeDropDown("step2", strFaktRateSteg2, "lstRateSteg2")%></td>
													<th class="right"><input name=PRFAKT100 type="text" size=4 value="<% =strFaktAntallSteg2 %>" onblur=hentVanligetimer2() id="Text14"></th>
												</tr>
												<tr>
													<th colspan=6><input type=SUBMIT value="Godkjenn timelisten" id="Submit3" name="Submit3"></th>
												</tr>
												<% 
											Else 
												If strFaktAntallSteg1 <> "" Then 'rate for overtime step 1
													strFaktRateSteg1 = (strFaktRateSteg1 - 1) * 100
												else
													strFaktRateSteg1 = 50
													strFaktAntallSteg1 = 0
												end if
												If strFaktAntallSteg2 <> "" Then 'rate for overtime step 2
													strFaktRateSteg2 = (strFaktRateSteg2 - 1) * 100
												else
													strFaktRateSteg2 = 100
													strFaktAntallSteg2 = 0
												end if													
												%>												
												<tr>
													<td>Ordinære lønnede timer:</td>
													<th class="right"><%=strVanlig%></th>
													<td>Ordinære fakturerte timer:</td>
													<th class="right"><%=strFaktVanlig%></th>
												</tr>
												<tr>
													<!-- viser overtids summer i glugger -->
													<td>Lønnede timer 50%:</td>
													<th class="right"><%=str50 %></th>
													<td>Fakturerte timer <%=strFaktRateSteg1%>%</td>
													<th class="right"><% =strFaktAntallSteg1 %></th>													
												</tr>																								
												<tr>
													<td>Lønnede timer 100%:</td>
													<th class="right"><%=str100 %></th>													
      												<td>Fakturerte timer <%=strFaktRateSteg2%>%</td>
													<th class="right"><% =strFaktAntallSteg2 %></th>
												</tr>												
												<%
											End If 'status < 5 
										%>
									</form>
								</table>
								<%
								' Display buttons for navigating between weeks
								%>
							</div>
							<table cellpadding='0' cellspacing='0' id="Table4">
								<tr>
									<%
									For i = 0 To ukeTeller - 1
										uk = uk + 1
										If uk = 20 Then
											Response.Write "</tr><tr>"
											uk = 1
										End If
										If ukeFontListe(i) = "<table bgcolor='black'>" then
											Font = ""
											FontEnd = ""
										Else
											font = "<font color='white'>"
											fontEnd = "</font>"
										End if

										If i = 0 and Left(ukeNrliste(i),6) = CStr(Session("displayweek")) And (ukeSplittLIste(i) <> "ja" And Request("splittuke") <> "Ja") And strNy <> "Ja" Then
											valgtUkedel = Left(ukeNrliste(i), 6)
												if ukeSplittLIste(i)="Ja" then 
													%>
													<script>settUkesplitt('Vikar_timeliste_vis3.asp?VikarID=<%=strVikarID%>&OppdragID=<%=strOppdragID%>&frakode=<%=frakode%>&splittuke=Ja&splittopt2=<%=Mid(ukeNrListe(i),7)%>')</script>
													<%
												end if
												%>
												<th>
													<%=ukeFontListe(i)%>
														<tr>
															<td><%=font%><% =Mid(ukeNrListe(i), 5) %><%=fontEnd%></td>
														</tr>
													</table>
												</th>
											<%
										ElseIf Not (request("ukedel")="1" or request("ukedel")="2") and i > 0 and Left(ukeNrliste(i), 6) <> valgtUkedel and Left(ukeNrliste(i), 6) = CStr(Session("displayweek")) And (ukeSplittLIste(i) <> "ja" And Request("splittuke") <> "Ja") Then
											%>
											<th>
												<%=ukeFontListe(i)%>
													<tr>
														<td><%=font%><% =Mid(ukeNrListe(i),5) %><%=fontEnd%></td>
													</tr>
												</table>
											</th>
											<%		
										Else
											If (Request("splittuke") = "Ja" OR request("ukedel")="1" OR request("ukedel")="2") And (Mid(ukeNrListe(i),7) = request("splittopt2") or Mid(ukeNrListe(i),8) = 		request("splittopt2")) And ukeSplittListe(i) = "Ja" And Left(ukeNrliste(i),6) = CStr(Session("displayweek")) Then
												if Mid(ukeNrListe(i), 8) = "2" then
													hideWeek = Mid(ukeNrListe(i), 5, 2)
												end if
												%>
												<th>
													<%=ukeFontListe(i)%>
														<tr>
															<td><%=Font%><% =Mid(ukeNrListe(i), 5)%><%=fontEnd%></td>
														</tr>
													</table>
												</th>
												<%		
											ElseIf not hideWeek = Mid(ukeNrListe(i), 5, 2) then	
												%>
												<form name="bbb" action="Vikar_timeliste_vis3.asp?VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&FirmaID=<% =strFirmaID %>&startDato=<% =ukeDatoListe(i) %>&splittuke=<% =ukeSplittListe(i) %>&splittopt2=<% =Right(ukeNrliste(i),2) %>" method=POST id="Form4">
													<th>
														<%=ukeFontListe(i)%>
															<tr>
																<td>
																	<input type="SUBMIT" value="<% =Mid(ukeNrListe(i),5) %>" id="Submit4" name="Submit4">
																</td>
															</tr>
														</table>
													</th>
												</form>
												<%
												if Mid(ukeNrListe(i), 8) = "2" then
													hideWeek = Mid(ukeNrListe(i), 5, 2)
												end if
											End If
										End If
									Next 
									%>
								</tr>
							</table>
							<%
							' Display more buttons
							frakode = session("frakode")
							If frakode < 10 Then
								If frakode > 1 Then  
									%>
									<table id="Table5">
										<tr>
											<form name="ccc" action="Vikar_timeliste_vis3s.asp?VikarID=<% =strVikarID %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&splittuke=<% =splittuke %>" method=POST id="Form5">
												<input type=SUBMIT value="          Sammendrag           " id="Submit5" name="Submit1">
											</form>										
											<% 
											If frakode > 1 Then 
												If frakode = 3 Then fkode2 = 3 Else fkode2 = 1 
												%>
												<form  name="ddd" action="Vikar_varl_vis3.asp?VikarID=<% =strVikarID %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&lagringskode=oppgrad&UkeNr=<% =session("DisplayWeek") %>&frakode2=<% =fkode2 %>" method=POST id="Form6">
													<th><input type=SUBMIT value="Variabel lønn" id="Submit6" name="Submit6"></th>
												</form>
												<%
												strSQL="SELECT AvdelingID FROM Oppdrag WHERE OppdragID = " & strOppdragID
												Set rsAvd = GetFirehoseRS(strSQL, Conn)
												strAvdeling = rsAvd("AvdelingID")
												rsAvd.Close
												Set rsAvd = Nothing
												
												if (lenb(splittopt) > 0) then
													if Left(splittopt, 1) <> "." then
														splittPart = "." & splittopt
													end if
												end if
												%>
												<form name="eee" action="Faktura_vis.asp?Kontakt=<%=strBestilltAv %>&VikarID=<% =strVikarID %>&OppdragID=<% =strOppdragID %>&SOKontakt=<%=strSOBestilltAv %>&FirmaID=<% =strFirmaID %>&frakode=<% =frakode %>&Avdeling=<% =strAvdeling %>" method="POST" id="Form7">
													<th>
														<input type=SUBMIT value="   Faktura   " id="Submit7" name="Submit7">
													</th>
												</form>
												<form name="GGG" action="Faktura_refresh.asp?VikarID=<% =strVikarID %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&lagringskode=nedgrad&UkeNr=<% =strWeekNumber %>&StartDato=<% =strStartDato %>&BestilltAV=<% =strBestilltAv %>&SOBestilltAV=<% =strSOBestilltAv %>&kode=1" method=POST id="Form8">
													<td>
														<input type=SUBMIT value="Hent til faktura" id="Submit8" name="Submit8">
													</td>
												</form>
												<% 
												if NedgradOK = true THEN
													%>
													<form name="fff" action="Vikar_timeliste_lagre3.asp?VikarID=<% =strVikarID %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&lagringskode=nedgrad&UkeNr=<% =session("displayweek") %>&StartDato=<% =strStartDato %>&BestilltAV=<% =strBestilltAv %>&SOBestilltAV=<% =strSOBestilltAv %>&kode=1&splittopt=<% =splittopt%>" method=POST id="Form9">
														<th>
															<input type="SUBMIT" title="Åpne timeliste for vikar" value="Nedgrader til t.status 1" id="Submit9" name="Submit9">
														</th>
													</form>
													<form name="xxx" action="Vikar_timeliste_lagre3.asp?VikarID=<% =strVikarID %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&lagringskode=nedgrad&UkeNr=<% =session("displayweek") %>&StartDato=<% =strStartDato %>&BestilltAV=<% =strBestilltAv %>&SOBestilltAV=<% =strSOBestilltAv %>&kode=1&tilStatus=2&splittopt=<% =splittopt %>" method=POST id="Form10">
														<th>
															<input type="SUBMIT" title="Nedgrader timelistestatus til 'godkjent av vikar'" value="Nedgrader til t.status 2" id="Submit10" name="Submit10">
														</th>
													</form>
													<% 
												End If 'NedgradOK 
											End If 
											viskode = Session("viskode") 
											%>
											<form name="hhh" action="Vikar_timeliste_vis3.asp?VikarID=<% =strVikarID %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&lagringskode=oppgrad&UkeNr=<% =strWeekNumber %>&frakode2=<% =fkode2 %>" method="POST" id="Form11">
												<td><input type="text" name=TILDATOO size=6 value=<% =session("limitdato") %> id="Text15"></td>
												<td><input type=SUBMIT value="Ny tildato" id="Submit11" name="Submit11"></td>
											</form>
										</tr>
									</table>
									<% 
								End If 'frakode > 1 
							End If 'frakode < 10 
							%>
							<table id="Table6">
								<tr>
									<% 
									If frakode = 2 Then 
										%>
										<form name="jjj" action=Vikar_timeliste_list3.asp?visKode=<% =viskode  %>&dato1=<% =session("limitDato") %>  method="POST" id="Form12">
											<th colspan="3">
												<input type="button" onclick="history.back(1);" value=" Tilbake  " id="Button1" name="Button1">
											</th>
										</form>
										<% 
									End If
									If frakode = 3 Then 
										%>
										<form name="kkk" action=Vikar_timeliste_fakt_list.asp?visKode=<% =viskode  %>&dato1=<% =session("limitDato") %>  method="POST" id="Form13">
											<th colspan="3">
												<input type="button"  onclick="history.back(1);" value="  Tilbake   " id="Button2" name="Button2">
											</th>
										</form>
										<% 
									End If 
									%>
								</tr>
							</table>
							<%
							session("sluttdato") = sluttdato
							session("startdato") = strStartdato

							' NEW MENUES AT THE BOTTOM
							If LenB(Request("VikarEndring")) > 0 Then
								Select Case Request("VikarEndring")
									Case 1: msg = "Slette timelister mellom datoer"
									Case 2: msg = "Hindre fakturering mellom datoer"
									Case 3: msg = "Endre pris mellom datoer"
									Case 4: msg = "Endre tider mellom datoer"
									Case Else
								End Select
							End If							
							If LenB(Request("VikarEndring")) > 0 Then 
								%>
								</div>
								<div class="contentHead"><h2><%=msg%></h2></div>
								<div class="content">
									<table id="Table7">
										<form name="ww" action="Vikar_timeliste_oppd_db3.asp?VikarEndring=<% =Request("VikarEndring") %>&VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&FirmaID=<% =strFirmaID %>&startDato=<% =strStartDato %>" method="POST" id="Form14">

											<tr>
												<th>F.o.m</th>
												<th>T.o.m</th>
												<% 
												If Request("VikarEndring") = 4 Then 'Endre tider mellom datoer
													%>
													<th>Fratid</th>
													<th>Tiltid</th>
													<th>Lunsj</th>
													<% 
												End If 
												If Request("VikarEndring") = 3 Then 'Endre pris mellom datoer
													%>
													<th>Timelønn</th>
													<th>Faktgrunnlag</th>
													<% 
												End If 
												%>
											</tr>
											<tr>
												<th><input class="mandatory" type="text" size=6 name=begdato onblur="dateCheck(this.form, this.name)" id="Text16"></th>
												<th><input class="mandatory" type="text" size=6 name=enddato onblur="dateCheck(this.form, this.name)" id="Text17"></th>
												<% 
												If Request("VikarEndring") = 3 Then 'Endre pris mellom datoer
													%>
													<th class="mandatory"><input type="text" size=6 name=Timelonn id="Text18"></th>
													<th class="mandatory"><input type="text" size=6 name=Fakturapris id="Text19"></th>
													<% 
												elseIf Request("VikarEndring") = 4 Then 'Endre tider mellom datoer
													%>
													<th><input type="text" class="mandatory" size=6 name=tbxFraKl onchange="timeCheck(this.form, this.name), workTime(this.form, this.name), hentFakttimer2(form.name) " id="Text20"></th>
													<th><input type="text" class="mandatory" size=6 name=tbxTilKl onchange="timeCheck(this.form, this.name), workTime(this.form, this.name), hentFakttimer2(form.name) " id="Text21"></th>
													<th><input type="text" class="mandatory"  size=4 name=tbxLunsj value="00:00" onchange="timeCheck(this.form, this.name), workTime(this.form, this.name), hentFakttimer2(form.name)"  id="Text22"></th>
													<input type=hidden name="tbxTimerPrDag" value="" id="Hidden15">
													<input type=hidden name="tbxFaktTimerPrDag" value="" id="Hidden16">
													<% 
												End If 
												%>
												<th><input type=SUBMIT value=" Lagre  " id="Submit12" name="Submit12"></th>
											</tr>
										</form>
									</table>
								<% 
							End If  'vikarendring

							If frakode < 10 Then
								%>
								<h5>|
									<%
									If fStatus < 3 Then
										if strStatus2 < 5 then
											%>
											<a href="Vikar_timeliste_vis3.asp?VikarEndring=1&VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&FirmaID=<% =strFirmaID %>&startDato=<% =strStartDato %>&splittopt2=<% =splittopt %>&splittopt=<% =splittopt %><% if splittopt="2" OR splittopt="1" then response.write "&splittuke=Ja"%>" >
											Slette mellom datoer</a> |
											<%
											If HasUserRight(ACCESS_TASK, RIGHT_ADMIN) = true OR HasUserRight(ACCESS_ADMIN, RIGHT_WRITE) = true Then
												%>
												<a href="Vikar_timeliste_slett_Alle.asp?VikarEndring=3&VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&FirmaID=<% =strFirmaID %>&startDato=<% =strStartDato %>&splittopt2=<% =splittopt %>&splittopt=<% =splittopt %>">
												Slett timelisten </a>|
												<% 
											End If 
											%>
											<a href="Vikar_timeliste_vis3.asp?VikarEndring=2&VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&FirmaID=<% =strFirmaID %>&startDato=<% =strStartDato  %>&splittopt2=<% =splittopt %>&splittopt=<% =splittopt %><% if splittopt="2" OR splittopt="1" then response.write "&splittuke=Ja"%>">
											Hindre fakturering mellom datoer</a>|
											<a href="Vikar_timeliste_vis3.asp?VikarEndring=3&VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&FirmaID=<% =strFirmaID %>&startDato=<% =strStartDato %>&splittopt2=<% =splittopt %>&splittopt=<% =splittopt %><% if splittopt="2" OR splittopt="1" then response.write "&splittuke=Ja"%>">
											Endre pris mellom datoer </a>|
											<br>|<a href="Vikar_timeliste_vis3.asp?VikarEndring=4&VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&FirmaID=<% =strFirmaID %>&startDato=<% =strStartDato %>&splittopt2=<% =splittopt %>&splittopt=<% =splittopt %><% if splittopt="2" OR splittopt="1" then response.write "&splittuke=Ja"%>">
											Endre tider mellom datoer </a>|
											<%
										end if 'strStatus2< 5
									End If 'fStatus < 3 
									%>
									<a href="Vikar_timeliste_oppd_vis2.asp?Bestilltav=<% =strBestilltav %>&SOBestilltav=<% =strSOBestilltAv%>&vikarid=<%=strVikarID %>&oppdragid=<%=strOppdragID %>&splittopt2=<% =splittopt %>&splittopt=<% =splittopt%>">Se alle timelister til denne kontaktpersonen</a>|<br>									
								</h5>
								<% 
							End If 'frakode < 10 
							rsTimeliste.Close 

							'Fjm:
							'Uttrekk for sist forlenge av oppdrag, dersom før 01.07.05 skal oppdraget ha andre %-satser hos økonomi
							If frakode < 10 Then
								StrSQL = "SELECT MAX([Fradato]) AS [Forlengelse] FROM [Oppdrag_Vikar] WHERE [OppdragID] = " & strOppdragID & " AND [VikarID] = " & strVikarID 

								Set rsExtensions = Conn.Execute(StrSQL)
								if (NOT rsExtensions.EOF) then
									%>
									<strong>Tidligere forlengelser:</strong>
									<%
									Response.Write rsExtensions("Forlengelse").value
									rsExtensions.close
									%>
									<br>
									<%
								end if
								set rsExtensions = nothing
							End If 'frakode < 10 

							'Get other links
							frakode = session("frakode")
							If frakode < 10 Then
								strSQL = "SELECT DISTINCT OppdragId FROM DAGSLISTE_VIKAR " &_
									"WHERE vikarID = " & strVikarID &_
									" and Loennstatus < 3 " &_
									" and Dato < " & dbDate(limitDato) &_
									" order by OppdragID"

								set rsOppdrag = GetFirehoseRS(strSQL, Conn)
								If Not rsOppdrag.EOF Then
									%>
									Oppdrag denne vikaren har:<br>
									<%
									do while Not rsOppdrag.EOF
										If CStr(rsOppdrag("OppdragID")) = strOppdragID Then %>
											| <% =rsOppdrag("OppdragID") %>
										<% Else %>
											| <a href="Vikar_timeliste_vis3.asp?VikarID=<% =strVikarID %>&OppdragID=<% =rsOppdrag("OppdragID") %>&frakode=<%=frakode%>" ><% =rsOppdrag("OppdragID") %></a>
										<% End IF
										rsOppdrag.MoveNext
									loop
									rsOppdrag.Close: set rsOppdrag = Nothing
									%>|<%
								End If 'ingen linjer i rsOppdrag
								dim SQLContact
								If (frakode > 1) Then
									if (IsNull(strBestilltAv)) then
										SQLContact =  " D.SOBestilltAv = " & strSOBestilltAv
									else
										SQLContact =  " D.BestilltAv = " & strBestilltAv
									end if
								
									strSQL = "SELECT DISTINCT VikarID, D.OppdragId FROM " &_
										"DAGSLISTE_VIKAR D, OPPDRAG " &_
										"WHERE " &_
										SQLContact  &_
										" and D.OppdragID = OPPDRAG.OppdragID" &_
										" and Fakturastatus < 3 " &_
										" and OPPDRAG.AvdelingID = " & strAvdeling &_
										" and Dato < " & dbDate(limitDato) &_
										" order by VikarID"
									set rsOppdrag = GetFirehoseRS(strSQL, Conn)
									If Not rsOppdrag.EOF Then
										%>
										<br>Vikarer på faktura:<br>
										<%
										do while Not rsOppdrag.EOF
											If CStr(rsOppdrag("VikarID")) = strVikarID Then 
												%>
												| <% =rsOppdrag("VikarID") %>
												<% 
											Else 
												%>
												| <a href="Vikar_timeliste_vis3.asp?VikarID=<% =rsOppdrag("VikarID") %>&OppdragID=<% =rsOppdrag("OppdragID") %>&frakode=<% =frakode %>" ><% =rsOppdrag("VikarID") %></a>
											<% 
											End If
											rsOppdrag.MoveNext
										loop
										rsOppdrag.Close: set rsOppdrag = Nothing
										%> |<%
									End If 'ingen linjer i rsOppdrag
								End If 'frakode > 1
							End If 'frakode < 10
						End If 	'No rows 

						' TILBAKEKNAPPER
						frakode = session("frakode")
						If Ingenlinjer = 1 Then
							%>
							<br>
							<table id="Table8">
								<tr>
									<% 
									If frakode = 1 Then 'Fra oppdrag
										%>
										<form name="mmm" action=../WebUI/OppdragView.aspx?OppdragID=<% =strOppdragID %>  method=POST id="Form15">
											<th colspan=2>
												<input type="submit"  value="                 Tilbake                " id="Submit13" name="Submit13">
											</th>
										</form>
										<% 
									End If 
									If frakode = 2 Then 'Timeliste (?)
										%>
										<form name="nnn" action=Vikar_timeliste_list3.asp?visKode=<% =viskode  %>&dato1=<% =session("limitDato") %>  method=POST id="Form16">
											<th colspan="2">
												<input type="submit"  value="                 Tilbake                " id="Submit14" name="Submit14">
											</th>
										</form>
										<% 
									End If 
									If frakode = 3 Then 'Fra faktura?
										%>
										<form name="ooo" action=Vikar_timeliste_fakt_list.asp?visKode=<% =viskode  %>&dato1=<% =session("limitDato") %>  method="POST" id="Form17">
											<th colspan="2">
												<input type="submit" value="                 Tilbake                " id="Submit15" name="Submit15">
											</th>
										</form>
										<% 
									End If 
									%>
								</tr>
							</table>
							</div>
						</div>
					</div>
				</body>
			</html>
			<%
		End If 'ingenlinjer 
		CloseConnection(Conn)
		set Conn = nothing		
		%>