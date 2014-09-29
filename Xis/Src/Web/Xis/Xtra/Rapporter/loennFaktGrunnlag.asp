
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Reports.Utils.inc"-->
<%
	If (HasUserRight(ACCESS_REPORT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim strSQL
	dim Conn
%>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en" "http://www.w3.org/tr/rec-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<meta name="generator" content="Microsoft Visual Studio 6.0">
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
	<%
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

	sub overtid(vikarNR, gmlUke, nyPeriode, ukeSumUtbetalt, Dag, inn)

		DagNr = datepart("d", Dag)
		UkedgNr = datepart("w", Dag)
		ukenrMnd = datepart("ww", dateAdd("d",-DagNr+1, Dag), 2, 2)
		splittKorr = ""

		if DagNr < 8 AND ukenrMnd = Datepart("ww", Dag, 2, 2) Then
			splittKorr = " AND not notat like '1' "
		end if

		a = dateAdd("d",-DagNr+1, Dag)
		'se a
		b = DateAdd("m",1, a) 
		'se b
		c = Datepart("d", DateAdd("d",-8, b))

		sisteUkeMnd =  datepart("ww", DateAdd("d",-1, b),2,2)
		ukenrMnd = datepart("ww", dag,2,2)
		 
		if DagNr > c AND ukenrMnd = sisteUkeMnd Then
			splittKorr = " AND not notat like '2' "
		end if
										
		strSQL = "SELECT overtidType= (H.LoennRate*100)-100, "&_
			" OppdragID, Antall, Sats, Belop, loennperiode, overfort_loenn_status, notat "&_
			" FROM VIKAR_UKELISTE V " &_
			" INNER JOIN H_loennsArt AS H ON V.loennsartnr = H.loennsartnr " &_
			" WHERE vikarid =" & vikarNR  &_
			" AND H.LoennRate > 1.0 "&_
			" AND Ukenr = "& gmlUke &_
			splittKorr &_
			" ORDER BY loennperiode, H.LoennRate, overfort_loenn_status"

		set rsOvertid = GetFireHoseRS(strSQL, Conn)
			
		if HasRows(rsOvertid) THEN 'hvis det er overtid registrert
			do until rsOvertid.EOF
				overtidBelop = rsOvertid("Belop")
				periode="" & rsOvertid("loennperiode")
				ukedel = ""
				
				if not trim(periode)= "" Then
					ukeSumUtbetalt=ukeSumUtbetalt + overtidBelop
					'sumTotal = sumTotal + overtidBelop
				end if
				if trim(periode)= "" Then
					gmlUkesumUbet = gmlUkesumUbet + overtidBelop
					periode="ikke lønnet"
					korrBelop = korrBelop + overtidBelop / (rsOvertid("overtidType")*100)*((rsOvertid("overtidType")*100)-100)
				end if
				if  not rsOvertid("notat") = " " then 
					ukedel="- ukedel " & rsOvertid("notat")
				End If	
				%>
				<tr>
					<td colspan="2">Overtid <%=rsOvertid("overtidType")%>%<%=ukedel%></td>
					<td><%=rsOvertid("oppdragID")%></td>
					<td>Periode: <%=periode%></td>
					<td>&nbsp;</td>
					<td><%=rsOvertid("Antall")%></td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td class="right"><%=formatNumber(overtidBelop, 2)%></td>
				</tr>
				<%	
			rsOvertid.MoveNext
			Loop
			rsOvertid.close				
		end if  
		set rsOvertid = nothing
	end sub

	sub ukesum(vikarNR, gmlUke, nyPeriode, ukeSumUtbetalt, Dag)
	
		'kode for å håndtere måneds-, årsskifte og splittuker
		DagNr = datepart("d", Dag)
		UkedgNr = datepart("w", Dag)
		ukenrMnd = datepart("ww", dateAdd("d", -DagNr+1, Dag), 2, 2)
		'se ukenrMnd
		'se UkedgNr& "  "& DagNr
		splittKorr = ""

		if DagNr < 8 AND ukenrMnd = Datepart("ww", Dag, 2, 2) Then
			splittKorr = " AND not notat like '1' "
		end if

		a = dateAdd("d", -DagNr+1, Dag)
		'se a
		b = DateAdd("m", 1, a) 
		'se b
		c = Datepart("d", DateAdd("d",-8, b))

		sisteUkeMnd =  datepart("ww", DateAdd("d", -1, b), 2, 2)
		ukenrMnd = datepart("ww", dag, 2, 2)
 
		if DagNr > c AND ukenrMnd = sisteUkeMnd Then
			splittKorr = " AND not notat like '2' "
		end if

		strSQL = "SELECT  V.OppdragID, V.Antall, V.Sats, V.Belop, V.Loennsartnr, V.loennperiode, V.notat, V.overfort_loenn_status "&_
				 " FROM VIKAR_UKELISTE AS V " &_
				 " INNER JOIN H_loennsArt AS H ON V.loennsartnr = H.loennsartnr " &_
				 " WHERE vikarid =" & vikarNR  &_
				 " AND H.LoennRate = 1.0 "&_
				 " AND Ukenr = " & gmlUke &_
				 splittKorr &_
				 " ORDER BY  V.Loennsartnr, V.loennperiode "

		set rsUkeloenn = Conn.Execute(strSQL)

		if HasRows(rsUkeloenn) THEN 

			oppdragID=0
			loennsart=0
			antall=0
			sats=0
			belop=0
			gmlLart=0
			gmlOpdrID=0
			gmlSats=0
			
			do until (rsUkeloenn.EOF)
				loennperiode=""&rsUkeloenn("loennperiode")
				loennsart=rsUkeloenn("Loennsartnr")
				oppdragID=rsUkeloenn("oppdragID")
				'antall=antall+rsUkeloenn("Antall")
				antall=rsUkeloenn("Antall")
				sats=rsUkeloenn("Sats")
				'belop=belop+rsUkeloenn("Belop")
				belop=rsUkeloenn("Belop")
				ukedel = ""
				loennstatus = rsUkeloenn("overfort_loenn_status")
				
				if  rsUkeloenn("notat") <> " "  Then
					'Response.Write (rsukeloenn("notat")) 
					ukedel = "-" &rsUkeloenn("notat")
				End If
					
				if not trim(loennperiode)="" then ukeSumUtbetalt=ukeSumUtbetalt+rsUkeloenn("Belop")
				'if not trim(loennperiode)="" then sumTotal = sumTotal + ukeSumUtbetalt 'rsUkeloenn("Belop")
					
				if  loennperiode = "" Then 
					if loennstatus > 2 then
						loennperiode = "ukjent"
					else
						loennperiode = "ikke utbetalt"						
					End if
				End if	
					
	 			if belop <> 0 then call skrivUtUkesum(loennsart, gmlUke, loennperiode, oppdragID, antall, sats, belop, ukedel)	
					
				gmlLart=loennsart
				gmlOpdrID=oppdragID
				gmlSats=sats
				rsUkeloenn.MoveNext
			
			Loop						
		end if  
		rsUkeloenn.close
		set rsUkeloenn = nothing
		sumTotal = sumTotal + ukeSumUtbetalt	
		call skrivUtUkesumTotal(gmlUke, gmlOpdrID, ukeSumUtbetalt, loennperiode )
		ukeSumUtbetalt = 0
	end sub

	sub skrivUtUkesum(loennsart, gmlUke, periode, oppdragID, antall, sats, belop, ukedel)
		%>
		<tr>
			<td colspan="4"><strong>Sum L.art: <%=loennsart%> uke: <%=gmlUke&ukedel%> Periode: <% =periode%> </strong></td>
			<td></td>
			<td><%=antall%></td>
			<td class="right"><%=sats%></td>
			<td></td>
			<td class="right"><%=formatNumber(belop,2)%></td>
		</tr>
		<%
	end sub

	sub skrivUtUkesumTotal(gmlUke, oppdragID, belop, loennperiode )
		%>
		<tr>
			<td colspan="4"><strong>Sum utbetalt uke: <%=gmlUke%> Periode: <% =loennperiode%> </strong></td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td class="right"><%=formatNumber(belop,2)%></td>
		</tr>
		<%
	end sub

	i = 0
	For each y in request("valg")
		i = i + 1
	next

	vikarid = request("valg")

	'Henter parametere for valgt måned og vikar

	periode =  request("periode")

	'se periode

	mnd =  mid(periode, 5)
	aar = left(periode, 4)

	nyPeriode = mnd
	if mnd<10 then nyPeriode="0" & mnd
	nyPeriode = aar & nyPeriode

	SELECT	CASE mnd
			CASE 1 
				mnd2 = "Januar" 
			CASE 2  
				mnd2 = "Februar" 
			CASE 3  
				mnd2 = "Mars"	
			CASE 4  
				mnd2 = "April" 
			CASE 5 
				mnd2 = "Mai"
			CASE 6 
				mnd2 = "Juni" 
			CASE 7  
				mnd2 = "Juli" 
			CASE 8  
				mnd2 = "August"
			CASE 9 
				mnd2 = "September"
			CASE 10  
				mnd2 = "Oktober" 
			CASE 11  
				mnd2 = "November" 
			CASE 12  
				mnd2 = "Desember"
		END SELECT

	'dato for generering av rapport
	Set dato = GetFireHoseRS("SELECT dato=getdate()", Conn)
	generertDato = dato("dato")
	dato.Close
	Set dato = nothing

	'henter lønn-/fakturagrunnlag for perioden
	strSQL = "SELECT O.avdelingID, D.vikarID, D.firmaID, O.ansmedID, uke=CONVERT(int, CONVERT(char(4), datepart(year, D.Dato))+CONVERT(char(2), datepart(week, D.Dato))), "&_
		" D.Loennstatus, D.splittuke, "&_
		" ukedag = datepart(weekday , D.Dato), "&_	
		" D.Dato, D.OppdragID, kunde=F.Firma, "&_
		" fraTidTT =  datepart(hh, D.Starttid), "&_
		" fraTidMM = "&_
		" CASE CONVERT(char(2), datepart(mi, D.starttid)) "&_
		"	WHEN '0' THEN '00' "&_
		"	ELSE  CONVERT(char(2), datepart(mi, D.starttid)) "&_
		" END, "&_
		" tilTidTT = datepart(hh, D.sluttid), "&_
		" tilTidMM =  "&_
		" CASE CONVERT(char(2), datepart(mi, D.sluttid)) "&_
		"	WHEN '0' THEN '00' "&_
		"	ELSE  CONVERT(char(2), datepart(mi, D.sluttid)) "&_
		" END, " & _
		" D.AntTimer, D.TimeLonn "&_
		" FROM DAGSLISTE_VIKAR D, FIRMA F, OPPDRAG O"&_	
		" WHERE D.VikarID in("& vikarid &")" &_
		" AND O.avdelingID IN("&request("avdelinger")&")"&_
		" AND O.oppdragID = D.oppdragID "&_
		" AND NOT D.antTimer = 0 "&_
		" AND D.FirmaID = F.FirmaID "&_
		" AND datepart(month, D.dato)="& mnd &_
		" AND datepart(year, D.dato)="& aar &_
		" ORDER BY  O.avdelingID, D.vikarID, D.dato, D.oppdragid "

	set rsOppdrag = GetFireHoseRS(strSQL, Conn)			
%>
	<div class="contentHead1">
		<h1>Lønn-/ fakturagrunnlag for <% =mnd2&"  "&aar %> </h1>
	</div>
	<div class="content">
	<h2>Rapport generert <% =generertDato %></h2>
	<p>
		(NB! Hvis ukesplitt ikke er i månedskiftet, vil dagene som vises ikke være i overenstemmelse med ukesummer.<br>
		Det er ukesummene som stemmer, da hele ukedeler må tas med. Ukedel 2 før månedskiftet og ukedel 1 etter månedskiftet tas med i summeringer.)
	</p>
	<div class="listing">
	<table id="Table1">
		<%
		'looper gjennom og skriver ut oppdrags-/timelistelinjer
		If NOT HasRows(rsOppdrag) Then 'hvis det ikke er registrert oppdrag i perioden
			response.write "<p class='warning'> det er ikke registrert oppdrag i perioden</p>"
		else

			SumAvdeling = 0
			SumAlle = 0

			Dim sumTotal
			sumTotal = 0

			do until rsOppdrag.EOF 'starter looping
				avdID = rsOppdrag("avdelingID")
				nyAvdID = avdID
				splittmerke =""

				'Navn på avdeling, funksjon fra library.inc
				avdeling = HentAvdNavn(AvdId)
				%>
					<tr>
						<td colspan="9">
							<h3>Avdeling: <%=avdeling%></h3>
						</td>
						<%	
						do until nyAvdID <> avdID 'starter sortering på avdeling
							
								vikarNR = rsOppdrag("vikarID")
								nyVikarID = vikarNR
								dag1 = rsOppdrag("Ukedag")
								
								strSQL = "SELECT v.vikarid, Navn=(v.Fornavn+' '+v.Etternavn), a.adresse, a.postnr, a.poststed " &_
										" FROM vikar v, adresse a "&_
										" WHERE v.vikarid in( " &vikarNR &")" &_
										" AND a.adresseRelID = v.vikarid "
								
								set RsVikar = GetFireHoseRS(strSQL, Conn)

								vikar = rsVikar("Navn")
								postadresse = rsVikar("adresse")
								poststed = rsVikar("postnr") & "  " & rsVikar("poststed")
								gmlUkedagNr = 1
								rsVikar.Close
								Set rsVikar = nothing						
								%>
								</tr>
								<tr>
									<td colspan="9">
										<p>
											<strong><% =vikar %></strong><br>
											<% =postadresse %><br>
											<% =poststed %><br>
										</p>
									</td>
								</tr>
								<tr>
									<td colspan="9">
										<table cellspacing="1" cellpadding="3" id="Table2">
											<tr>
												<th>Ukedag</th>
												<th>Dato</th>
												<th>OppdNr</th>
												<th>Kunde</th>
												<th>Tid</th>
												<th>Timer</th>
												<th>Lønn</th>
												<th>Utbet</th>
												<th>Sum</th>
									<%
									start=true
									ukeSumUbet=0
									gmlUkesumUbet=0
									ukeSumUtbetalt=0
		'***************************** START LOOP *****************************************************************
		do until nyVikarID <> vikarNR OR rsOppdrag.EOF 'starter looping
			
			uke = rsOppdrag("uke")
			uke = "" & uke			
			if len(uke)= 5 Then
				uke = (trim(left(uke,4) & "0" & right(uke,1)))
			end If				
			uke = cLng(uke)			
			ukedag = rsOppdrag("Ukedag")
			ukedagNr = rsOppdrag("Ukedag")
			if Not rsOppdrag("splittuke") = "" AND not IsNull(rsOppdrag("splittuke")) Then splittmerke = " - " & rsOppdrag("splittuke")			
			utbetalt = rsOppdrag("Loennstatus")
			sumDag = ( Round((rsOppdrag("antTimer")),2) * rsOppdrag("Timelonn"))
			tempSumDag=sumDag
				SELECT	CASE utbetalt
					CASE 3 
						utbet = "ja" 
						SumDag=0
		
					CASE ELSE
						utbet = "nei"
						ukeSumUbet=ukeSumUbet+sumDag	
				END SELECT

			'sumTotal = sumTotal + sumDag
				SELECT	CASE ukedag
					CASE 1 
						Ukedag = "Søndag" 
					CASE 2  
						Ukedag = "Mandag" 
					CASE 3  
						Ukedag = "Tirsdag"	
					CASE 4  
						Ukedag = "Onsdag" 
					CASE 5 
						Ukedag = "Torsdag"
					CASE 6 
						Ukedag = "Fredag"
					CASE 7  
						Ukedag = "Lørdag" 
				END SELECT
			

aarr =  rsOppdrag("dato")
aarMinus = 0

if datepart("d", aarr)< 8 AND datepart("m", aarr) = 1 AND not datepart("ww", rsOppdrag("dato"),2,2) = 1 Then aarMinus = 1 
	call datoKorreksjon("yyyy", aarr)
	aarr=aarr-aarMinus
	uukkee = rsOppdrag("dato")
	call datoKorreksjon("ww", uukkee)
If uukkee < 10 then uukkee = "0" & uukkee
overtidUke = aarr & uukkee

uke=overtidUke

'se gmlUke & " - " & Uke & " - " & ukedagNr 

			if  ukedagNr > 1 AND  not uke=gmlUke AND not start=true   then 'gmlUkedagNr  Then 'ny uke 
				ukeTest=false
				korrBelop=0
				%>
				<%
				call overtid(vikarNR, gmlUke, nyPeriode, ukeSumUtbetalt, mndDag,1)
				call ukesum(vikarNR, gmlUke, nyPeriode, ukeSumUtbetalt, mndDag)
				if rsOppdrag("Loennstatus")< 3 Then
					gmlUkesumUbet=UkeSumUbet-sumDag+korrBelop
					sumTotIkkeUtbet = sumTotIkkeUtbet + gmlUkesumUbet					
				else 
					gmlUkesumUbet=0										
				end if
				%>
				<tr>
					<td colspan="8"><strong>Sum ikke utbetalt for uke <% =gmlUke %></strong></td>
					<td class="right"><%=formatNumber(gmlUkesumUbet,2)%></td>
				</tr>
				<%
				
				ukeSumUbet=sumDag

			end IF 'ny uke					
			%>
				<tr>
					<td><% =ukedag%></td>
					<td><% =rsOppdrag("Dato")&splittmerke%></td>
					<td><% =rsOppdrag("OppdragID")%></td>
					<td><% =rsOppdrag("Kunde")%></td>
					<td><% =rsOppdrag("fratidTT")&":"&rsOppdrag("fraTidMM")&"-"& rsOppdrag("tiltidTT")&":"&rsOppdrag("tiltidMM")%></td>
					<td><% =Round((rsOppdrag("antTimer")),2) %></td>
					<td class="right"><% =rsOppdrag("Timelonn") %></td>
					<td><%=utbet%></td>
					<td class="right"><%=formatNumber(tempSumDag,2)%></td>
				</tr>
			<%		
			splittmerke=""
			mndDag = rsOppdrag("dato")
'se mndDag
			if ukedag = dag1 Then
				dag1 = ""
			end if
			gmlUkedagNr = ukedagNr
			gmlUkeTemp=gmlUke
			'se gmlUkeTemp&"  DDDDD"& uke
			gmlUke = uke
			ukeTest=true
			start=false
	
			rsOppdrag.MoveNext			
			'gmlUkedagNr=""
			if NOT rsOppdrag.EOF THEN
				nyVikarID = rsOppdrag("vikarID")
			End if
			if nyVikarID<>vikarNR OR rsOppdrag.EOF then 'AND NOT IsEmpty(sumTotal) Then 'skriver ut totalsum for måneden
				korrBelop=0
				
				%>
				<%

				call overtid(vikarNR, uke, nyPeriode, ukeSumUtbetalt, mndDag, 2)
				call ukesum(vikarNR, uke, nyPeriode, ukeSumUtbetalt, mndDag)
				
				gmlUkesumUbet=UkeSumUbet+korrBelop

				sumTotIkkeUtbet = sumTotIkkeUtbet + gmlUkesumUbet				
				
			%>
				<tr>
					<td colspan="8"><strong>Sum ikke utbetalt for uke <% =gmlUke %></strong></td>
					<td class="right"><%=formatNumber(gmlUkesumUbet,2)%></td>
				</tr>
				<tr>
					<td colspan="8"><strong>Sum total for hele måneden - utbetalt - inkludert eventuell overtid</strong></td>
					<td class="right"><%=formatNumber(sumTotal,2)%></td>
				</tr>
				<tr>
					<td colspan="8"><strong>Sum total for hele måneden - ikke utbetalt - inkludert eventuell overtid</strong></td>
					<td class="right"><%=formatNumber(sumTotIkkeUtbet,2)%></td>
				</tr>
			
			<%			
			sumTotal=0
			sumTotIkkeUtbet = 0
			end if 'skriver ut totalsum for måneden			
		loop
		if NOT rsOppdrag.EOF THEN				
			nyAvdID = rsOppdrag("avdelingID")
		Else
			exit do
		End if
	Loop 'avdelingsloop
		
Loop 'hovedloop slutt
rsOppdrag.close
set rsOppdrag = nothing
end if 'hvis det er registrert oppdrag i perioden
									%>	
								</table>
							</td>
						</tr>
					</table>
				</div>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>