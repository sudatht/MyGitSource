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
	dim Conn
	dim strSQL
	dim intValgtAvd 'Id til valgt avdeling fra kriterie side
	dim strAvdWHERE 'SQL for betingelser på avdeling	
	
	sub hentOvertid(AVD, tiluke, Otype)
		if Otype = 1 then ' hvis ikke fakturert
			'henter og aggregerer overtid for avdelingen i perioden - gjelder ikke fakturert'
			strSQL = " SELECT FTimer=SUM(V.Fakturatimer), "&_
				" FBelop = SUM((V.Fakturabeloep/(F.FakturaSats*100))*((F.FakturaSats*100)-100)), " &_				
				" T.navn, O.tomid " &_
	 			" FROM OPPDRAG O, VIKAR_UKELISTE V, VIKAR VK, TJENESTEOMRADE T, H_Faktura_Type F " &_
				" WHERE V.Overfort_fakt_status < 3 "&_
				" AND T.tomid = O.tomid "&_
				" AND V.Fakturapris > 0 "&_
				" AND F.FakturaSats > 1.0 " &_
				" AND V.OppdragID = O.OppdragID "&_
				" AND V.FakturaType = F.FakturaType  "&_
				" AND O.AvdelingID = "& AVD &_
				" AND VK.TypeID != 3 "&_
				" AND VK.vikarID = V.vikarID "&_
				" AND V.ukenr <= "& tiluke &_
				" AND not V.ID in(SELECT ID from vikar_ukeliste WHERE ukenr=" & tiluke &" AND notat like '2') "&_
				" GROUP BY O.AvdelingID, T.navn, O.tomid "

			Set rsOvertid = GetFirehoseRS(strSQL, Conn)

		end if ' ikke fakturert
		if Otype = 2 Then 'ikke lønnet
			'* henter og aggregerer overtid for avdelingen i perioden - gjelder ikke lønnet *'
			strSQL = " SELECT LTimer=SUM(V.antall), "&_
				" LBelop=SUM((V.Belop/(L.LoennRate*100))*((L.LoennRate*100)-100)) "&_
	 			" FROM OPPDRAG O, VIKAR_UKELISTE V, VIKAR VK, H_loennsart L "&_
				" WHERE V.Overfort_loenn_status < 3 "&_
				" AND V.Belop > 0 " &_
				" AND V.OppdragID = O.OppdragID " &_
				" AND V.Loennsartnr = L.Loennsartnr " &_
				" AND L.LoennRate > 1.0 "&_
				" AND O.AvdelingID = "& AVD &_
				" AND VK.TypeID != 3 "&_
				" AND VK.vikarID = V.vikarID "&_
				" AND V.ukenr <= "& tiluke &_
				" AND not V.ID IN(SELECT ID from vikar_ukeliste WHERE ukenr="& tiluke &" AND notat like '2') "&_
				" GROUP BY O.AvdelingID "

			Set rsOvertid = GetFirehoseRS(strSQL, Conn)
			if HasRows(rsOvertid) Then
				LTimer = rsOvertid("LTimer")
				LBelop = rsOvertid("LBelop")
				rsOvertid.close
			end if
			set rsOvertid = nothing
			OvertAvd(Avd)= LBelop
		end if
	end sub
	
	'Henter inn valgte verdier
	'fradato = Request("fradato")
	tildato = Request("tildato")
	lb_loennet = Request("ikke_loennet")
	lb_fakt = Request("ikke_fakt")
	intValgtAvd = Request("dbxAvdeling")

	if intValgtAvd > 0 then
		strAvdWHERE = " AND o.avdelingid = " & intValgtAvd
	else
		strAvdWHERE = ""
	end if

	if tildato = "" then
		tildato = 0
	end if

	'Feilmelding hvis verken lønn eller fakturert er valgt
	If lb_loennet <> "on" AND lb_fakt <> "on" then
		AddErrorMessage("Verken ikke lønnet eller ikke fakturert er valgt. Gå tilbake og kryss av i minst en av boksene for å få fram rapporten.")
		call RenderErrorMessage()			
	End if
	
%>
<html>
	<head>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script type="text/javascript" language="javascript" src="../js/javascript.js"></script>
		<title>Periodisering</title>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<a id="Top"></a>
			<div class="contentHead1">
				<h1>Periodisering - vanlig</h1>
			</div>
			<div class="content">
				<div class="listing">
				<%
				' Get a database connection
				Set Conn = GetConnection(GetConnectionstring(XIS, ""))

				' The size of the two OvertAvd and SumAvd arrays below need to be defined AFTER we know the number of
				' branches. They could be defined beforehand and oversized, but that would be a waste of memory.
				' They are ReDim-ed a few lines lower, when we know the number of branches.
				dim OvertAvd() 	    'Overtid for hver av avdelingene
				dim SumAvd()	 	'Sum for hver av avdelingene
				dim AntAvd	 		'Antall avdelinger
				dim intAntTOM		'Totalt antall tjenesteområder
				dim rsTOM			'Recordset for antall tjenesteområder
				dim AToms(20,4)		'Setter av god plass for evnt. nye tjenesteområder
				dim intForrigeTOM	'Id'en til forrige tjenesteområde i loopen
				dim rsOvertid		'RC for overtid

				'Atoms(x,0) = Tomid
				'Atoms(x,1) = Navn
				'Atoms(x,2) = Avdelingssum
				'Atoms(x,3) = Sum alle avdeling

				'Finner totalt antall avdelinger som er registrert i basen
				strSQL = "SELECT stoerste = MAX(AvdelingID) FROM avdeling"
				set rsAntallAvd = GetFirehoseRS(strSQL, Conn)
				If (HasRows(rsAntallAvd)) then
					AntAvd = rsAntallAvd("stoerste")
					rsAntallAvd.close
				Else
					AntAvd = 1
				End If
				Set rsAntallAvd = Nothing
				
				' NOW, knowing the number of branches, we can define the size of the arrays initially at line 113 and 114:
				ReDim OvertAvd(cint(AntAvd))
				ReDim SumAvd(cint(AntAvd))

				'Finner totalt antall tjenesteområder som er registrert i basen
				strSQL= "SELECT Tomid, navn from tjenesteomrade ORDER BY tomid"
				set rsTOM = GetFirehoseRS(strSQL, Conn)
				intAntTOM = 0
				while not rsTOM.EOF
					Atoms(intAntTOM,0) = rsTOM("tomid").value
					Atoms(intAntTOM,1) = rsTOM("navn").value
					Atoms(intAntTOM,2) = 0
					Atoms(intAntTOM,3) = 0

					intAntTOM = intAntTOM + 1
					rsTOM.movenext
				wend
				rsTOM.close
				Set rsTOM = Nothing

				intAntTOM = intAntTOM - 1
				intForrigeTOM = 0

			'for å kompansere for feil ukenr i kalender..
			aarr = tildato
			call datoKorreksjon("yyyy", aarr)
			uukkee = tildato
			call datoKorreksjon("ww", uukkee)
			If uukkee < 10 then uukkee = "0" & uukkee
			tiluke = aarr & uukkee
			
			'Sjekker om ikke fakturert skal være med
			If lb_fakt = "on" then 'Ta med ikke fakturert

				' sql for ikke fakturert
				strSQL = "SELECT dmin = MIN(dato), dmax = MAX(dato), O.AvdelingID, O.tomid, O.OppdragID, T.FirmaID, F.Firma, T.VikarID, Navn=(V.Fornavn + ' ' + V.Etternavn)," &_
					" FTimer=SUM(T.Fakturatimer)," &_
					" T.Fakturapris," &_
					" VIKAR_ANSATTNUMMER.ansattnummer" & _
					" from DAGSLISTE_VIKAR T, OPPDRAG O, FIRMA F, VIKAR V LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON V.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " &_
					" WHERE T.OppdragID = O.OppdragID" &_
					strAvdWHERE &_
					" AND T.VikarID = V.VikarID" &_
					" AND T.FirmaID = F.FirmaID" &_
					" AND T.Fakturastatus < 3" &_
					" AND T.Fakturapris > 0" &_
					" AND T.Dato <= " & dbDate(tildato) &_
					" AND not V.TypeID = 3 "&_
					" GROUP BY O.AvdelingID, O.OppdragID, O.tomid, T.FirmaID, F.Firma, T.VikarID, V.Fornavn, V.Etternavn, VIKAR_ANSATTNUMMER.ansattnummer" &_
					", T.Fakturapris" &_
					" ORDER BY O.AvdelingID, O.OppdragID, O.tomid "

				Set rsPeriode = GetFirehoseRS(strSQL, Conn)
				If not rsPeriode.EOF Then
					' Utlisting av data
					%>
					<table>
					<% 
					'overskrift - avdlingsnavn
					strSQL = "SELECT Avdeling from AVDELING WHERE AvdelingID = " & rsPeriode("AvdelingID")
					Set rs = GetFirehoseRS(strSQL, Conn)
					aa = rs("Avdeling"): rs.Close: Set rs = Nothing
					%>
					<tr>
						<th colspan="8" >Ikke fakturert&nbsp;Avdeling:&nbsp;<%=aa%>&nbsp;Periode:&nbsp;<% =fradato %> - <% =tildato %>
					</tr>
					<tr>
						<th>OppdNr</th>
						<th>Uke</th>
						<th>Kontakt</th>
						<th>Ansattnr</th>
						<th>Navn</th>
						<th>Timer</th>
						<th>Pris</th>
						<th>Sum</th>
					</tr>
					<%
					sumsum = 0
					sumsumsum = 0
					sumtim = 0
					sumkurs = 0: sumdata = 0: sumdok = 0: sumintern=0

					'Nuller ut variable for alle avdelinger

					For teller = 1 to AntAvd
						sumAvd(Teller) = 0
					Next

					For teller = 1 to AntAvd
						overtAvd(teller) = 0
					Next

					'overtKurs = 0: overtData = 0: overtDok = 0: overtIntern = 0
					AVD = rsPeriode("AvdelingID")

					do while (not rsPeriode.EOF)
						If AVD <> rsPeriode("AvdelingID") Then 'skjer ved ny avdeling
							for TomTeller = 0 to intAntTOM
								if Atoms(Tomteller,2) > 0 then
									response.write "<tr><td colspan='5'>Sum for " & Atoms(Tomteller,1) & " (" & Atoms(Tomteller,0) & ")</td><td>&nbsp;</td><td>&nbsp;</td><td class='right'>" & formatNumber(Atoms(Tomteller,2),2) & "</td></tr>"
								end if
								Atoms(TomTeller,2) = 0
							next
							%>
							<tr>
								<TD colspan="5">Sum alle tjenesteområder:</TD>
								<TD class="right"><% = formatNumber(sumtim, 2) %></TD>
								<TD>&nbsp;</TD>
								<TD class='right'><% = formatNumber(sumsum, 2) %></TD>
							</tr>
							<%
							FTimer=0
							FBelop=0
							call hentOvertid(AVD, tiluke, 1)
							if(HasRows(rsOvertid)) then
								while not rsOvertid.EOF
									response.write "<tr><TD colspan='5' >Sum overtidstillegg for " & rsOvertid("navn") & " (" & rsOvertid("tomid") & ")</TD>"
									response.write "<TD class='right'>" & formatNumber(rsOvertid("FTimer"),2) & "</TD>"
									response.write "<TD>&nbsp;</TD>"
									response.write "<TD class='right'>" & formatNumber(rsOvertid("FBelop"),2) &"</TD></TR>"
									FBelop = FBelop + rsOvertid("FBelop")
									FTimer = FTimer + rsOvertid("FTimer")
									rsOvertid.movenext
								wend
								rsOvertid.close
							end if
							set rsOvertid = nothing
							response.write "<tr><TD colspan='5'>Sum overtidstillegg totalt:</TD>"
							response.write "<TD class='right'>" & formatNumber(FTimer, 2) & "</TD>"
							response.write "<TD>&nbsp;</TD>"
							response.write "<TD class='right'>" & formatNumber(FBelop, 2) &"</TD></TR>"
							OvertAvd(AVD) = FBelop
							%>
							<tr>
								<TD colspan="5">Sum total:</TD>
								<TD class="right"><% = formatNumber(sumtim,2) %></TD>
								<TD>&nbsp;</TD>
								<TD class="right"><% = formatNumber(sumsum + FBelop,2) %></TD>
							</tr>
							</table>
							<p id="sideskift" class="pageBreakAfter">&nbsp;</p>
							<table>
							<% 
							strSQL = "SELECT Avdeling from AVDELING WHERE AvdelingID = " & rsPeriode("AvdelingID")
							Set rs = GetFirehoseRS(strSQL, Conn)
							aa = rs("Avdeling"): rs.Close: Set rs = Nothing 
							%>
							<tr>
								<th colspan="8">Ikke fakturert
								&nbsp;
								Avdeling:&nbsp;<% =aa %>&nbsp;
								&nbsp;
								Periode:&nbsp;<% =fradato %> - <% =tildato %>
								</th>
							</tr>
							<tr>
								<th>OppdNr</th>
								<th>Uke</th>
								<th>Kontakt</th>
								<th>Ansattnr</th>
								<th>Navn</th>
								<th>Timer</th>
								<th>Pris</th>
								<th>Sum</th>
							</tr>
							<%
							'summere og sette variabler på nytt
							sumsumsum = sumsumsum + sumsum
							sumsum = 0
							sumtim = 0
							AVD = rsPeriode("AvdelingID")
						End If 'ny avdeling 
						%>
						<tr>
							<TD class=right><% =rsPeriode("OppdragID") %>
							<%
							dmin = rsPeriode("dmin"): dmax = rsPeriode("dmax")
							call datoKorreksjon("ww", dmin)
							call datoKorreksjon("ww", dmax)
							%>
							<TD><%=dmin%> - <%=dmax%>
							<TD><% =rsPeriode("Firma") %>
							<TD class="right"><% =rsPeriode("ansattnummer") %>
							<TD><% =rsPeriode("Navn") %>
							<% 
							ft = formatNumber(rsPeriode("FTimer"), 2)
							%>
							<TD class="RIGHT" ><% =ft %>
							<TD class=right ><% = rsPeriode("Fakturapris")%>
							<% 
							sum = ft * rsPeriode("Fakturapris") 
							%>
							<TD class=right><% =formatNumber(sum, 2) %></TD>
						</tr>
						<%
						'summering per avd.
						sumsum = sumsum + sum
						sumtim = sumtim + ft
						SumAvd(AVD) = sumsum
						for TomTeller = 0 to intAntTOM
							if Atoms(Tomteller,0) = rsperiode("tomid") then
								Atoms(Tomteller,2) = Atoms(Tomteller,2) + sum
								Atoms(Tomteller,3) = Atoms(Tomteller,3) + sum
								exit for
							end if
						next
						rsPeriode.MoveNext
					loop
					rsPeriode.Close
					Set rsPeriode = Nothing
					for TomTeller = 0 to intAntTOM
						if Atoms(Tomteller,2) > 0 then
							response.write "<tr><td colspan='5'>Sum for " & Atoms(Tomteller,1) & " (" & Atoms(Tomteller,0) & ")</td><td>&nbsp;</td><td>&nbsp;</td><td class='right'>" & formatNumber(Atoms(Tomteller,2),2) & "</td></tr>"
						end if
						Atoms(TomTeller,2) = 0
					next
					%>
					<tr>
						<TD colspan="5">Sum alle tjenesteområder:</TD>
						<TD class='right'><% = formatNumber(sumtim,2) %></TD>
						<TD>&nbsp;</TD>
						<TD class='right'><% = formatNumber(sumsum,2) %></TD>
					</tr>
					<%
					FTimer=0
					FBelop=0
					call hentOvertid(AVD, tiluke, 1)
					if(HasRows(rsOvertid)) then
						while not (rsOvertid.EOF)
							response.write "<tr><TD colspan='5'>Sum overtidstillegg for " & rsOvertid("navn") & " (" & rsOvertid("tomid") & ")</TD>"
							response.write "<TD class='right'>" & formatNumber(rsOvertid("FTimer"),2) & "</TD>"
							response.write "<TD>&nbsp;</TD>"
							response.write "<TD class='right'>" & formatNumber(rsOvertid("FBelop"),2) &"</TD></TR>"
							FBelop = FBelop + rsOvertid("FBelop")
							FTimer = FTimer + rsOvertid("FTimer")
							rsOvertid.movenext
						wend
						rsOvertid.close						
					end if
					set rsOvertid = nothing					
					response.write "<tr><TD colspan='5'>Sum overtidstillegg totalt:</TD>"
					response.write "<TD class='right'>" & formatNumber(FTimer,2) & "</TD>"
					response.write "<TD>&nbsp;</TD>"
					response.write "<TD class='right'>" & formatNumber(FBelop,2) &"</TD></TR>"
					OvertAvd(AVD) = FBelop
					%>
					<tr>
						<TD colspan="5">Sum total:</TD>
						<TD class="right"><% = formatNumber(sumtim,2) %></TD>
						<TD>&nbsp;</TD>
						<TD class="right"><% = formatNumber(sumsum+FBelop,2) %></TD>
					</TR>
				</table>
				<p id="sideskift" class="pageBreakAfter">&nbsp;</p>
				<%
				sumsumsum = sumsumsum + sumsum
				%>
				<table>
				<tr>
					<th colspan="2" >Samlet ikke fakturert i perioden&nbsp;<% =fradato %> - <% =tildato %></th>
					<%
					For teller = 1 to AntAvd
						strSQL = "SELECT Avdeling from AVDELING WHERE AvdelingID = " & Teller
						Set rs = GetFirehoseRS(strSQL, Conn)
						aa = rs("Avdeling")
						rs.Close: Set rs = Nothing
						sumsumsum = sumsumsum + OvertAvd(teller)
						%>
						<tr>
							<TD><%=aa%></TD>
							<TD class="right"><%=formatNumber(sumAvd(Teller) + overtAvd(Teller),2) %></TD>
						</tr>
						<%
					next
					%>
					<tr>
						<th>Sum</th>
						<th class="right"><% = formatNumber(sumsumsum, 2) %></th>
					</tr>
				</table>
			</table><br>
			<p id="sideskift" class="pageBreakAfter">&nbsp;</p>
			<%
			For teller = 1 to AntAvd
				OvertAvd(teller) = 0
			Next
			%>
			<table>
				<tr>
					<th colspan="2">Samlet ikke fakturert per tjenesteområde i perioden&nbsp;<% =fradato %> - <% =tildato %></th>
				</tr>
				<tr>
					<TD>Tjenesteområde</TD><TD class="RIGHT">Beløp</TD>
				</tr>
					<%
					sumsumsum = 0
					for TomTeller = 0 to intAntTOM
						response.write "<tr><td>Samlet sum for " & Atoms(Tomteller, 1) & " (" & Atoms(Tomteller,0) & ")</td><td class='right'>" & formatNumber(Atoms(Tomteller,3),2) & "</td></tr>"
						sumsumsum = sumsumsum + Atoms(Tomteller, 3)
					next
					response.write "<tr><td>Samlet sum alle tjenesteområder</td><td class='right'>" & formatNumber(sumsumsum,2) & "</td></tr>"
					%>
				</TR>
			</table><br>
			<table>
				<tr>
					<th colspan="3">Samlet ikke fakturert overtid per tjenesteområde i perioden&nbsp;<% =fradato %> - <% =tildato %></th>
					<%
					strSQL = " SELECT FTimer=SUM(V.Fakturatimer), "&_
					" FBelop = SUM((V.Fakturabeloep/(F.FakturaSats*100))*((F.FakturaSats*100)-100)), " &_	
					" T.navn, T.tomid "&_
	 				" FROM OPPDRAG O, VIKAR_UKELISTE V, VIKAR VK, TJENESTEOMRADE T, H_Faktura_Type F " &_
					" WHERE V.Overfort_fakt_status < 3 "&_
					strAvdWHERE & _
					" AND T.tomid = O.tomid "&_
					" AND F.FakturaSats > 1.0 " &_
					" AND V.OppdragID = O.OppdragID " &_
					" AND V.FakturaType = F.FakturaType " &_
					" AND V.Fakturapris > 0 "&_
					" AND VK.TypeID != 3 "&_
					" AND VK.vikarID = V.vikarID "&_
					" AND V.ukenr <= "& tiluke &_
					" AND not V.ID IN (SELECT ID from vikar_ukeliste WHERE ukenr="& tiluke &" AND notat like '2') "&_
					" GROUP BY T.navn, T.tomid "

					Set rsOvertid = GetFirehoseRS(strSQL, Conn)

					sumsumsum = 0
					sumsumtimer = 0
					response.write "<tr><td>Tjenesteområde</td><td class='right'>Timer</td><td class='right'>Beløp</td></tr>"
					if (HasRows(rsOvertid)) then
						while not rsOvertid.EOF
							response.write "<tr><td>Samlet overtid for " & rsOvertid("navn") & " (" & rsOvertid("tomid") & ")</td><td class='right'>" & formatNumber(rsOvertid("ftimer"),2) & "</td>"
							response.write "<td class='right'>" & formatNumber(rsOvertid("fbelop"), 2) & "</td></tr>"
							sumsumsum = sumsumsum + rsOvertid("ftimer")
							sumsumtimer = sumsumtimer + rsOvertid("fbelop")
							rsOvertid.movenext
						wend
						rsOvertid.close
					end if
					set rsOvertid = nothing
					response.write "<tr><td>Samlet overtid alle tjenesteområder</td><td class='right'>" & formatNumber(sumsumsum,2) & "</td>"
					response.write "<td class='right'>" & formatNumber(sumsumtimer, 2) & "</td></tr>"
					%>
				</TR>
			</table>
			<%
		Else
			Response.Write "Ikke fakturert: Ingen forekomster!"
		End If 'ingen rader
		End if 'lb_fakt = "on"

			If lb_loennet = "on" then
				' sql for ikke lønnet
				strSQL = "SELECT T.VikarID, T.OppdragID, T.FirmaID, F.Firma, O.AvdelingID, Navn=(V.Fornavn + ' ' + V.Etternavn)" &_
					", LTimer=SUM(T.AntTimer)" &_
					", T.Timelonn" &_
					", VIKAR_ANSATTNUMMER.ansattnummer" & _
					" FROM DAGSLISTE_VIKAR T, OPPDRAG O, FIRMA F, VIKAR V LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON V.Vikarid = VIKAR_ANSATTNUMMER.Vikarid" &_
					" WHERE T.OppdragID = O.OppdragID" &_
					strAvdWHERE &_
					" AND T.VikarID = V.VikarID" &_
					" AND T.FirmaID = F.FirmaID" &_
					" AND T.Loennstatus < 3" &_
					" AND T.Timelonn > 0" &_
					" AND not V.TypeID = 3 "&_
					" AND T.Dato <= " & dbDate(tildato) &_
					" GROUP BY O.AvdelingID, T.OppdragID, T.FirmaID, F.Firma, T.VikarID, V.Fornavn, V.Etternavn, VIKAR_ANSATTNUMMER.ansattnummer" &_
					", T.Timelonn" &_
					" ORDER BY O.AvdelingID, T.OppdragID"

					Set rsPeriode = GetFirehoseRS(strSQL, Conn)
					If not rsPeriode.EOF Then
					' opplisting av data
					%>
					<table>
						<% 
						strSQL = "SELECT Avdeling from AVDELING WHERE AvdelingID = " & rsPeriode("AvdelingID")
						Set rs = GetFirehoseRS(strSQL, Conn)
						aa = rs("Avdeling"): rs.Close: Set rs = Nothing 
						%>
						<tr>
							<th colspan=8 >Ikke lønnet&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Avdeling:&nbsp;<% =aa %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Periode:&nbsp;<% =fradato %> - <% =tildato %>
							</th>
						</tr>
						<tr>
							<th>OppdNr</th>
							<th>Uke</th>
							<th>Kontakt</th>
							<th>Ansattnr</th>
							<th>Navn</th>
							<th>Timer</th>
							<th>Pris</th>
							<th>Sum</th>
						</tr>
						<%
						sumsum = 0
						sumsumsum = 0
						sumtim = 0
						'sumkurs = 0: sumdata = 0: sumdok = 0: sumintern=0
						For teller = 1 to AntAvd
							SumAvd(teller) = 0
						Next
						AVD = rsPeriode("AvdelingID")
						do while not rsPeriode.EOF
							If AVD <> rsPeriode("AvdelingID") Then 'for hver avdeling
								LTimer=0
								LBelop=0
								call hentOvertid(AVD, tiluke, 2)
								%>
								<tr>
									<TD colspan="5">Sum:</TD>
									<th class=right><% = sumtim %></th>
									<TD>&nbsp;</td>
									<th class=right><% = formatNumber(sumsum,2) %></th>
								</tr>
								<tr>
									<TD colspan="5">Sum overtidstillegg:</TD>
									<th class=right><% = formatNumber(LTimer,2) %></th>
									<TD>&nbsp;</td>
									<th class=right><% = formatNumber(LBelop,2) %></th>
								</tr>
								<tr>
									<TD colspan="5">Sum total:</TD>
									<th class=right><% = formatNumber(sumtim,2) %>
									<TD>&nbsp;<th class=right><% = formatNumber(sumsum+LBelop,2) %></th>
								</tr>
								</table>
								<p id="sideskift" class="pageBreakAfter">&nbsp;</p>
								<table >
								<% 
								strSQL = "SELECT Avdeling from AVDELING WHERE AvdelingID = " & rsPeriode("AvdelingID")
								Set rs = GetFirehoseRS(strSQL, Conn)
								aa = rs("Avdeling"): rs.Close: Set rs = Nothing 
								%>
								<tr>
									<th colspan=8 >Ikke lønnet&nbsp;Avdeling:&nbsp;<% =aa %>&nbsp;Periode:&nbsp;<% =fradato %> - <% =tildato %></th>
								</tr>
								<tr>
									<th>OppdNr</th>
									<th>Uke</th>
									<th>Kontakt</th>
									<th>Ansattnr</th>
									<th>Navn</th>
									<th>Timer</th>
									<th>Pris</th>
									<th>Sum</th>
								</tr>
								<%
								sumsumsum = sumsumsum + sumsum

								sumsum = 0
								sumtim = 0
								AVD = rsPeriode("AvdelingID")
							End If 'for hver avdeling 
							%>
							<tr>
								<TD class=right><% =rsPeriode("OppdragID") %></td>
								<% 
								strSQL = "SELECT dmin=Min(dato), dmax=Max(dato)" &_
									" FROM Dagsliste_vikar" &_
									" WHERE OppdragID = " & rsPeriode("OppdragID") &_
									" AND VikarID = " & rsPeriode("VikarID") &_
									" AND Loennstatus < 3" &_
									" AND Fakturapris > 0" &_
									" AND Dato <= " & dbDate(tildato)
					
								set ddd = GetFirehoseRS(strSQL, Conn)
								dmin = ddd("dmin")
								call datoKorreksjon("ww", dmin)
								dmax = ddd("dmax")
								call datoKorreksjon("ww", dmax)
								%>
								<TD><% =dmin %>-<% =dmax %></td>
								<% 
								ddd.Close: Set ddd = Nothing 
								%>
								<TD><% =rsPeriode("Firma") %></td>
								<TD class=right><% =rsPeriode("ansattnummer") %></td>
								<TD><% =rsPeriode("Navn") %>
								<% 
								lt = FormatNumber(rsPeriode("LTimer"),2) 
								%>
								<TD class="right" ><% =lt %></td>
								<TD class="right" ><% =rsPeriode("Timelonn") %></TD>
								<% 
								sum = rsPeriode("LTimer") * rsPeriode("timelonn") 
								%>
								<TD class="right" ><% =formatNumber(sum,2) %></TD>
								<%
								sumtim = sumtim + lt
								sumsum = sumsum + sum
								SumAvd(AVD) = sumsum
								rsPeriode.MoveNext
							loop
							rsPeriode.Close: Set rsPeriode = Nothing
							LTimer=0
							LBelop=0
							call hentOvertid(AVD, tiluke, 2)
							%>
							<tr>
								<TD colspan="5">Sum:</TD>
								<th class=right><% = sumtim %></th>
								<TD>&nbsp;</td>
								<th class=right><% = formatNumber(sumsum,2) %></th>
							</tr>
							<tr>
								<TD colspan="5">Sum overtidstillegg:</TD>
								<th class=right><% = formatNumber(LTimer,2) %></th>
								<TD>&nbsp;</td>
								<th class=right><% = formatNumber(LBelop,2) %></th>
							</tr>
							<tr>
								<TD colspan="5">Sum total:</TD>
								<th class=right><% = formatNumber(sumtim,2) %></th>
								<TD>&nbsp;</td>
								<th class=right><% = formatNumber(sumsum+LBelop,2) %></th>
							</tr>
							<%
							for TomTeller = 0 to intAntTOM
								response.write "<tr><td colspan=5>Sum totalt for " & Atoms(Tomteller,1) & "</td><td>&nbsp;</td><td>&nbsp;</td><td class='right'>" & formatNumber(Atoms(Tomteller,3),2) & "</td></tr>"
							next
							%>
						</table>
						<p id="sideskift" class="pageBreakAfter">&nbsp;</p>
						<table>
						<%
						sumsumsum = sumsumsum + sumsum
						%>
						<tr>
							<th colspan="2">Samlet ikke lønnet i perioden&nbsp;<% =fradato %> - <% =tildato %></th>
						</tr>
						<% 
						For teller = 1 to AntAvd
							strSQL = "SELECT Avdeling from AVDELING WHERE AvdelingID = " & Teller
							Set rs = GetFirehoseRS(strSQL, Conn)
							aa = rs("Avdeling"): rs.Close: Set rs = Nothing
							sumsumsum=sumsumsum + overtAvd(teller)
							%>
							<tr>
								<TD><% =aa %></td>
								<TD class=right><% = formatNumber(sumAvd(Teller)+overtAvd(Teller),2) %></TD>
							</tr>
							<%
						next
						%>
						<tr>
							<th>Sum</th>
							<th class="right"><% = formatNumber(sumsumsum,2) %></th>
						</tr>
					</table>
				</table>
					<%
					Else
						Response.Write "Ikke lønnet: Ingen forekomster!"
					End If 'ingen rader
				End if 'lb_loennet = "on"
				%>
						<a href="#top"><img src="../Images/icon_GoToTop.gif" alt="Til toppen">Til toppen<a>
					</div>
				</div>
			</div>
		</body>
	</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>
