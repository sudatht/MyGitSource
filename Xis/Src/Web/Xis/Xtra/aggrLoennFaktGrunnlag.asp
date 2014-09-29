<%@ Language=VBScript %>
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
i=0
For each y in request("valg")
i=i+1
next

vikarid = request("valg")
'********************************************************************************************
'Henter parametere for valgt måned og vikar


periode =  request("periode")


mnd =  mid(periode, 5)
aar = left(periode, 4)

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
<h1>Lønn-/ fakturagrunnlag for <% =mnd2&"  "&aar %></h1>
<h2>Rapport generert <% =generertDato %></h2>

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
		<td colspan="9">
			<h3>Avdeling: <%=avdeling%></h3>
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
		</tr>
		<tr>
			<td colspan="9">
				<h4><% =vikar %></h4>
				<p>
					<% =postadresse %><br>
					<% =poststed %>
				</p>
			</td>
			<table cellspacing="1" cellpadding="5">
		
		<%
		do until nyVikarID<>vikarNR OR rsOppdrag.EOF 'starter looping
			
			uke = rsOppdrag("uke")
			ukedag = rsOppdrag("Ukedag")
			ukedagNr = rsOppdrag("Ukedag")
			
			utbetalt = rsOppdrag("Loennstatus")
			
				SELECT	CASE utbetalt
					CASE 3 
						utbet = "ja" 
					CASE ELSE
						utbet = "nei"
				END SELECT
			
			sumDag = (rsOppdrag("antTimer") * rsOppdrag("Timelonn"))
			sumTotal = sumTotal + sumDag
							
			'****************************************************************************************************
			'henter ut opplysninger om overtid fra uken som er ferdig skrevet ut
			
			
			if  ukedagNr < gmlUkedagNr Then 'ny uke AND NOT dag1 = 7 
			
		
				strSQL = "Select overtidType= "&_
						" CASE Loennsartnr "&_
							" WHEN 160 THEN 50 "&_
							" WHEN 163 THEN 100 "&_
						" END, "&_
						" OppdragID, Antall, Sats, Belop "&_
						" from VIKAR_UKELISTE "&_
						" where vikarid ="& vikarNR  &_
						" and Loennsartnr in(160,163)"&_
						" and Ukenr = "& gmlUke &_
						" order by Loennsartnr "
				set rsOvertid = Conn.Execute(strSQL)
				'Response.Write rsOvertid.source &"<br>"
				
					if NOT rsOvertid.EOF THEN 'hvis det er overtid registrert
					
						do until rsOvertid.EOF
							sumTotal = sumTotal + rsOvertid("Belop")
				
							rsOvertid.MoveNext
						Loop
						
					end if  'hvis det er overtid registrert
				rsOvertid.close
				set rsOvertid = nothing
			end IF 'ny uke
		
			if ukedag = dag1 Then
				dag1 = ""
			end if
			gmlUkedagNr = ukedagNr
			gmlUke = uke
			
			rsOppdrag.MoveNext			
			'gmlUkedagNr=""
			if NOT rsOppdrag.EOF THEN
				nyVikarID = rsOppdrag("vikarID")
			End if		
			
			if nyVikarID<>vikarNR OR rsOppdrag.EOF AND NOT IsEmpty(sumTotal) Then 'skriver ut totalsum for måneden
			
				if  ukedagNr > 5 Then 'ny uke
			
				strSQL = "Select overtidType= "&_
					" CASE Loennsartnr "&_
						" WHEN 160 THEN 50 "&_
						" WHEN 163 THEN 100 "&_
					" END, "&_
					" OppdragID, Antall, Sats, Belop "&_
					" from VIKAR_UKELISTE "&_
					" where vikarid ="& vikarNR &_
					" and Loennsartnr in(160,163)"&_
					" and Ukenr = "& gmlUke &_
					" order by Loennsartnr "
				set rsOvertid = Conn.Execute(strSQL)
				'Response.Write rsOvertid.source &"<br>"
			
					if NOT rsOvertid.EOF THEN 'hvis det er overtid registrert
						do until rsOvertid.EOF
							sumTotal = sumTotal + rsOvertid("Belop")
							rsOvertid.MoveNext
						Loop
					
					end if  'hvis det er overtid registrert
				rsOvertid.close
				set rsOvertid = nothing
				end IF 'ny uke
			%>
			<tr>
				<td>Sum total for hele måneden inkludert eventuell overtid</td>
				<td class="right"><%=formatNumber(sumTotal,2)%></td>
			
			<%
			SumAvdeling = SumAvdeling + sumTotal
			sumTotal=0
			end if 'skriver ut totalsum for måneden	
			
		loop
		if NOT rsOppdrag.EOF THEN				
			nyAvdID = rsOppdrag("avdelingID")
		Else
			exit do
		End if
	Loop 'avdelingsloop
	
	sumAlle = sumAlle + sumAvdeling
	%>	
			</tr>
			<tr>
				<td colspan=9><hr></td>
			</tr>
			<tr>
				<td>Sum total for avdeling <strong><% =avdeling%></strong> hele måneden inkludert eventuell overtid</td>
				<td class="right"><%=formatNumber(sumAvdeling,2)%></td>
<%
Loop 'hovedloop slutt

%>
			</tr>
			<tr>
				<td colspan="9"><hr></td>
			</tr>
			<tr>
				<td>Sum total for alle avdelinger hele måneden inkludert eventuell overtid</td>
				<td class="right"><%=formatNumber(sumAlle,2)%></td>
<%
rsOppdrag.close
set rsOppdrag = nothing

end if 'hvis det er registrert oppdrag i perioden
%>
			</tr>	
		</table>
</table>

    </div>
</body>
</html>

