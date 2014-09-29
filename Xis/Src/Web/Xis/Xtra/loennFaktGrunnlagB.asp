<%@ Language=VBScript %>
<!--#INCLUDE FILE="datafiler/Library.inc"--> 

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"
    "http://www.w3.org/TR/REC-html40/loose.dtd">
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
'********************************************************************************************
'oppretter databaseforbindelse

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")
'*******************************************************************************************
sub overtid(vikarNR, gmlUke, nyPeriode, ukeSumUtbetalt)

								
				strSQL = "Select overtidType= "&_
						" CASE Loennsartnr "&_
							" WHEN 160 THEN 50 "&_
							" WHEN 161 THEN 50 "&_
							" WHEN 162 THEN 50 "&_							
							" WHEN 163 THEN 100 "&_
							" WHEN 164 THEN 100 "&_
							" WHEN 165 THEN 100 "&_
						" END, "&_
						" OppdragID, Antall, Sats, Belop, loennperiode "&_
						" from VIKAR_UKELISTE "&_
						" where vikarid ="& vikarNR  &_
						" and Loennsartnr in(160,161,162,163,164,165)"&_
						" and Ukenr = "& gmlUke &_
						" and (isNull(loennperiode) or loennperiode="& nyPeriode &")"&_						
						" order by loennperiode, Loennsartnr "


						

				set rsOvertid = Conn.Execute(strSQL)
				Response.Write rsOvertid.source &"<br><br><br>"

					if NOT rsOvertid.EOF THEN 'hvis det er overtid registrert
					
						do until rsOvertid.EOF
							sumTotal = sumTotal + rsOvertid("Belop")
					%>
							<tr><TD colspan=2><u><i>Overtid <%=rsOvertid("overtidType")%>%</i></u></td>
							<td><%=rsOvertid("oppdragID")%></td><td>Periode: <%=rsOvertid("loennperiode")%></td><td></td><td><%=rsOvertid("Antall")%></td><td class="right"><%=rsOvertid("Sats")%></td>
							<td></td>
							<td class="right"><%=formatNumber(rsOvertid("Belop"),2)%></td>
					<%

							ukeSumUtbetalt=ukeSumUtbetalt+formatNumber(rsOvertid("Belop"),2)
							rsOvertid.MoveNext
						Loop
						
					end if  
							
				rsOvertid.close
				set rsOvertid = nothing			


end sub
sub ukesum(vikarNR, gmlUke, nyPeriode, ukeSumUtbetalt)

		strSQL = "Select  OppdragID, Antall, Sats, Belop, Loennsartnr "&_
				 " from VIKAR_UKELISTE "&_
				 " where vikarid ="& vikarNR  &_
				 " and Loennsartnr in(180,194,190,198,660,150,151,152,153,154,155,156,157,158,159)"&_
				 " and Ukenr = "& gmlUke &_
				 " and loennperiode="& nyPeriode &_
				 " order by Loennsartnr "

				set rsUkeloenn = Conn.Execute(strSQL)
'se strSQL

		if NOT rsUkeloenn.EOF THEN 'hvis det er overtid registrert

		oppdragID=0
		loennsart=0
		antall=0
		sats=0
		belop=0
		gmlLart=0
		gmlOpdrID=0
		gmlSats=0
		
			
			do until rsUkeloenn.EOF

				loennsart=rsUkeloenn("Loennsartnr")
				oppdragID=rsUkeloenn("oppdragID")
				antall=antall+rsUkeloenn("Antall")
				sats=rsUkeloenn("Sats")
				belop=belop+rsUkeloenn("Belop")
				sumTotal = sumTotal + rsUkeloenn("Belop")
				ukeSumUtbetalt=ukeSumUtbetalt+rsUkeloenn("Belop")
					
				if ( not loennsart=gmlLart and not gmlLart=0) or (not oppdragID=gmlOpdrID and not gmlOpdrID=0) or (not sats=gmlSats and not gmlSats=0)   Then 
					call skrivUtUkesum(gmlLart, gmlUke, gmlOpdrID, antall, gmlSats, belop)					
					belop=0
					antall=0
					
				end if
				'call skrivUtUkesum(loennsart, gmlUke, oppdragID, antall, sats, belop)
				gmlLart=loennsart
				gmlOpdrID=oppdragID
				gmlSats=sats
			rsUkeloenn.MoveNext
			
			Loop
			
			if belop<>0 then call skrivUtUkesum(gmlLart, gmlUke, gmlOpdrID, antall, gmlSats, belop)
			
						
		end if  
		rsUkeloenn.close
		set rsUkeloenn = nothing	
		call skrivUtUkesumTotal(gmlUke, gmlOpdrID, ukeSumUtbetalt)
		ukeSumUtbetalt=0
end sub
sub skrivUtUkesum(loennsart, gmlUke, oppdragID, antall, sats, belop)

%>
					<tr><TD colspan=2><u><strong>Sum L.art <%=loennsart%><br> uke <%=gmlUke%></strong></u></td>
					<td><%=oppdragID%></td><td></td><td></td><td><%=antall%></td><td class="right"><%=sats%></td>
					<td></td>
					<td class="right"><%=formatNumber(belop,2)%></td>
<%
end sub
sub skrivUtUkesumTotal(gmlUke, oppdragID, belop)

%>
					<tr><TD colspan=2><u><strong>Sum utbetalt<br> uke <%=gmlUke%></strong></u></td>
					<td><%=oppdragID%></td><td></td><td></td><td></td><td class="right"></td>
					<td></td>
					<td class="right"><%=formatNumber(belop,2)%></td>
<%
end sub
i=0
For each y in request("valg")
i=i+1
next

vikarid = request("valg")
'********************************************************************************************
'Henter parametere for valgt måned og vikar


periode =  request("periode")

'se periode

mnd =  mid(periode, 5)
aar = left(periode, 4)

nyPeriode=mnd
if mnd<10 then 	nyPeriode="0"&mnd
nyPeriode=aar&nyPeriode
'se nyPeriode

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


'**********************************************************************************************							
'dato for generering av rapport
Set dato = Conn.Execute("select dato=getdate()")
generertDato = dato("dato")
dato.Close
Set dato = nothing

'**********************************************************************************************	
'henter lønn-/fakturagrunnlag for perioden

strSQL = "select O.avdelingID, D.vikarID, D.firmaID, O.ansmedID, uke=CONVERT(int, CONVERT(char(4), datepart(year, D.Dato))+CONVERT(char(2), datepart(week, D.Dato))), "&_
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
	" END, "&_
	" D.AntTimer, D.TimeLonn "&_
	" from DAGSLISTE_VIKAR D, FIRMA F, OPPDRAG O"&_	
	" where D.VikarID in("& vikarid &")" &_
	" And O.avdelingID IN("&request("avdelinger")&")"&_
	" And O.oppdragID = D.oppdragID "&_
	" And NOT D.antTimer = 0 "&_
	" And D.FirmaID = F.FirmaID "&_
	" And datepart(month, D.dato)="& mnd &_
	" And datepart(year, D.dato)="& aar &_
	" order by O.avdelingID, D.vikarID, D.dato, D.oppdragid "
'Response.Write(strSQL)					
set rsOppdrag = Conn.Execute(strSQL)				

'***********************************************************************************************
					
%>

<table cellpadding='0' cellspacing='0'>
<h4>Lønn-/ fakturagrunnlag for <% =mnd2&"  "&aar %> </h4>
<H5 ><i>Rapport generert <% =generertDato %></i></h5>

<%
'******************************************************************************************************************
'looper gjennom og skriver ut oppdrags-/timelistelinjer

If rsOppdrag.EOF Then 'hvis det ikke er registrert oppdrag i perioden
	response.write "<h3> det er ikke registrert oppdrag i perioden</h3>"
else

SumAvdeling = 0
SumAlle = 0

Dim sumTotal


do until rsOppdrag.EOF 'starter looping
avdID = rsOppdrag("avdelingID")
nyAvdID = avdID

avdeling = avdID
	If      avdeling = "1" THEN
		avdeling = "Kurs"
	elseIf  avdeling = "2" THEN
		avdeling = "Data"
	elseIf  avdeling = "3" THEN
		avdeling = "Dokument"
	elseIf  avdeling = "4" THEN
		avdeling = "Intern"
	end if

%>
<tr>
<td colspan=9>
<hr>
<h4>Avdeling: <%=avdeling%></h4>
</td>
<%	do until nyAvdID<>avdID 'starter sortering på avdeling
	
		vikarNR = rsOppdrag("vikarID")
		nyVikarID = vikarNR
		dag1 = rsOppdrag("Ukedag")
		
		'********************************************************************************************
		'Henter opplysninger om vikar

		Set rsVikar = Conn.Execute("Select  v.vikarid, Navn=(v.Fornavn+' '+v.Etternavn), a.adresse, a.postnr, a.poststed "&_
					" from vikar v, adresse a "&_
					" where v.vikarid in( " &vikarNR &")" &_
							" and a.adresseRelID=v.vikarid ")
		'Response.Write(rsVikar.Source)	
		vikar = rsVikar("Navn")
		postadresse = rsVikar("adresse")
		poststed = rsVikar("postnr")&"  "&rsVikar("poststed")
		gmlUkedagNr = 1
		rsVikar.Close
		Set rsVikar = nothing						
		%>
		<tr>
		<td colspan=9>
		<hr>
		<strong><% =vikar %></strong><br>
		<% =postadresse %><br>
		<% =poststed %><br>
		<hr>
		</td>
		<table cellpadding='0' cellspacing='0'>
		<th>Ukedag</th><th>Dato</th><th>OppdNr</th><th>Kontakt</th><th>Tid</th><th>Timer</th><th>Lønn</th><th>Utbet</th><th>Sum</th>
		<%
		start=true
		ukeSumUbet=0
		gmlUkesumUbet=0
		ukeSumUtbetalt=0
		do until nyVikarID<>vikarNR OR rsOppdrag.EOF 'starter looping
			


			uke = rsOppdrag("uke")
			uke = ""&uke			
			if len(uke)= 5 Then
				uke = (trim(left(uke,4)&"0"&right(uke,1)))
			end If				
			uke = cLng(uke)			
			ukedag = rsOppdrag("Ukedag")
			ukedagNr = rsOppdrag("Ukedag")			
			utbetalt = rsOppdrag("Loennstatus")
			sumDag = (rsOppdrag("antTimer") * rsOppdrag("Timelonn"))
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
			
			'****************************************************************************************************
			'henter ut opplysninger om overtid fra uken som er ferdig skrevet ut
			


'for å kompansere for feil ukenr i kalender...
'*********************************************
aarr =  rsOppdrag("dato")
call datoKorreksjon("yyyy", aarr)
uukkee = rsOppdrag("dato")
call datoKorreksjon("ww", uukkee)
If uukkee < 10 then uukkee = "0" & uukkee
overtidUke = aarr & uukkee

uke=overtidUke

			if  ukedagNr > 1 and  not uke=gmlUke and not start=true then 'gmlUkedagNr  Then 'ny uke 
			
			ukeTest=false		      								
				call overtid(vikarNR, gmlUke, nyPeriode, ukeSumUtbetalt)
				call ukesum(vikarNR, gmlUke, nyPeriode, ukeSumUtbetalt)
				if rsOppdrag("Loennstatus")<3 Then
					gmlUkesumUbet=UkeSumUbet-sumDag					
				else 
					gmlUkesumUbet=0										
				end if
				%>
				<tr>
				<td colspan=8>Sum ikke utbetalt for uke <% =gmlUke %></td>
				<td class="right"><%=formatNumber(gmlUkesumUbet,2)%></td>

				<%
				
				ukeSumUbet=sumDag

			end IF 'ny uke					
			%>
			<TR>			
			<td><% =ukedag%></td><td><% =rsOppdrag("Dato")%></td><td><% =rsOppdrag("OppdragID")%></td><td><% =rsOppdrag("Kunde")%></td><td><% =rsOppdrag("fratidTT")&":"&rsOppdrag("fraTidMM")&"-"& rsOppdrag("tiltidTT")&":"&rsOppdrag("tiltidMM")%></td><td><% =rsOppdrag("antTimer") %></td><td class="right"><% =rsOppdrag("Timelonn") %></td><TD align="center"><%=utbet%></td><td class="right"><%=formatNumber(tempSumDag,2)%></td>
			<%			
			if ukedag = dag1 Then
				dag1 = ""
			end if
			gmlUkedagNr = ukedagNr
			gmlUkeTemp=gmlUke
			gmlUke = uke
			ukeTest=true
			start=false
	
			rsOppdrag.MoveNext			
			'gmlUkedagNr=""
			if NOT rsOppdrag.EOF THEN
				nyVikarID = rsOppdrag("vikarID")
			End if
			
			if nyVikarID<>vikarNR OR rsOppdrag.EOF AND NOT IsEmpty(sumTotal) Then 'skriver ut totalsum for måneden
			
				call overtid(vikarNR, gmlUkeTemp, nyPeriode, ukeSumUtbetalt)
				call ukesum(vikarNR, gmlUkeTemp, nyPeriode, ukeSumUtbetalt)
			
			%>
			<tr>
			<td colspan=8>	Sum ikke utbetalt for uke <% =gmlUke %></td>
			<td class="right"><%=formatNumber(gmlUkesumUbet,2)%></td>
			<tr>
			<TD colspan=8>Sum total for hele måneden inkludert eventuell overtid</td>
			<td class="right"><%=formatNumber(sumTotal,2)%></td>			
			<%			
			sumTotal=0
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
</table>
    </div>
</body>
</html>

