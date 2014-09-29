<%@ LANGUAGE="VBSCRIPT" %>
<%
Response.Expires = 0
%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Reports.Utils.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
<%
	If (HasUserRight(ACCESS_REPORT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim Conn
	dim strSQL
	dim i
	dim y
	dim generertDato
	dim dato
	dim firmaid
	dim SumAvdeling
	dim SumAlle
	Dim sumTotal
	
	'oppretter databaseforbindelse
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

	i = 0
	For each y in request("valg")
		i = i + 1
	next

	firmaid = request("valg")

	'Henter parametere for valgt måned og vikar
	periode = request("periode")
	mnd =  mid(periode, 5)
	aar = left(periode, 4)

	SELECT CASE mnd
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
	Set dato = GetFirehoseRS("SELECT dato = getdate()", Conn)
	generertDato = dato("dato")
	dato.Close
	Set dato = nothing	

	'henter lønn-/fakturagrunnlag for perioden
	strSQL = "SELECT O.avdelingID, D.firmaID, O.ansmedID, uke = CONVERT(int, CONVERT(char(4), datepart(year, D.Dato))+CONVERT(char(2), datepart(week, D.Dato))), " &_
		" D.Loennstatus, D.splittuke, D.fakturastatus, D.fakturapris, "&_
		" ukedag = datepart(weekday , D.Dato), "&_	
		" D.Dato, D.OppdragID, vikar=(v.fornavn+' '+v.etternavn), "&_
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
		" FROM VIKAR V, DAGSLISTE_VIKAR D, FIRMA F, OPPDRAG O "&_
		" WHERE  f.firmaID in ("& firmaid &")"&_
		" AND O.avdelingID IN("&request("avdelinger")&")"&_
		" AND O.oppdragID = D.oppdragID "&_
		" AND NOT D.antTimer = 0 "&_
		" AND D.FirmaID = F.FirmaID "&_
		" AND V.vikarid = D.vikarid "&_
		" AND datepart(month, D.dato)="& mnd &_
		" AND datepart(year, D.dato)="& aar &_
		" ORDER BY O.avdelingID, D.firmaID, D.dato, D.oppdragid "

	set rsOppdrag = GetFirehoseRS(strSQL, Conn)
	%>
	<h1>Fakturagrunnlag for <% =mnd2 & "  " & aar %></h1>
	<h2>Rapport generert <%=generertDato%></h2>
	<table>
	<%
	'looper gjennom og skriver ut oppdrags-/timelistelinjer
	if (HasRows(rsOppdrag) = false) Then 'hvis det ikke er registrert oppdrag i perioden
		response.write "<p class='warning'> det er ikke registrert oppdrag i perioden</p>"
	else
		SumAvdeling = 0
		SumAlle = 0
		dag1 = rsOppdrag("Ukedag")

		do until rsOppdrag.EOF 'starter looping
			avdID = rsOppdrag("avdelingID")
			nyAvdID = avdID

			avdeling = avdID
			If  avdeling = "1" THEN
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
				<td colspan="2">
					<h3>Avdeling: <%=avdeling%></h3>
				</td>
				<%
				do until nyAvdID <> avdID 'starter sortering på avdeling
					firmaNR = rsOppdrag("firmaID")
					nyFirmaID = firmaNR
		
					'Henter opplysninger om firma
				
					strSQL = "SELECT  f.firmaid, f.firma, a.adresse, a.postnr, a.poststed "&_
						" FROM firma f, adresse a " &_
						" WHERE  f.firmaid in (" & firmaNR & ")" &_
						" AND a.adresseRelID = f.firmaid "

					Set rsFirma = GetFirehoseRS(strSQL, Conn)

					firma = rsFirma("firma")
					postadresse = rsFirma("adresse")
					poststed = rsFirma("postnr") & "  " & rsFirma("poststed")
					rsFirma.Close
					Set rsFirma = nothing			
					%>
				</tr>
				<tr>
				<td colspan=2>
					<h3><%=firma%></h3>
					<p>
						<%=postadresse %><br>
						<%=poststed %>
					</p>
				</td>
				<table>
				<%
				do until nyFirmaID <> firmaNR OR rsOppdrag.EOF
				
					uke = rsOppdrag("uke")
					ukedag = rsOppdrag("Ukedag")
					ukedagNr = rsOppdrag("Ukedag")
					fakt = rsOppdrag("fakturastatus")

						SELECT	CASE fakt
							CASE 3 
								fakt = "ja" 
							CASE ELSE
								fakt = "nei"
						END SELECT

					sumDag = (rsOppdrag("antTimer") * rsOppdrag("fakturapris"))
					sumTotal = sumTotal + sumDag
				
					'henter ut opplysninger om overtid fra uken 
					if (ukedagNr < gmlUkedagNr) Then 'ny uke					
						strSQL = "SELECT overtidType= " &_
								" CASE Loennsartnr "&_
									" WHEN 160 THEN 35 "&_
									" WHEN 163 THEN 70 "&_
								" END, "&_
								" OppdragID, fakturatimer, fakturapris, fakturabeloep "&_
								" FROM VIKAR_UKELISTE "&_
								" WHERE  vikarid in("& firmaNR &")" &_
								" AND Loennsartnr in(160,163)"&_
								" AND Ukenr = "& gmlUke &_
								" ORDER BY Loennsartnr "
								
						set rsOvertid = GetFirehoseRS(strSQL, Conn)
			
						if HasRows(rsOvertid) THEN 'hvis det er overtid registrert
							do until rsOvertid.EOF
								sumTotal = sumTotal + rsOvertid("fakturaBelop")
								rsOvertid.MoveNext
							Loop					
							rsOvertid.close
						end if  'hvis det er overtid registrert
						set rsOvertid = nothing
					end IF 'ny uke
					if ukedag = dag1 Then
						dag1 = ""
					end if
					gmlUkedagNr = ukedagNr
					sluttUke = gmlUke
					gmlUke = uke
					rsOppdrag.MoveNext

					if NOT rsOppdrag.EOF THEN
						nyFirmaID = rsOppdrag("firmaID")
					End if		
					if nyFirmaID <> firmaNR OR rsOppdrag.EOF AND NOT IsEmpty(sumTotal) Then 'skriver ut totalsum for måneden 

						if  ukedagNr > 5 Then 'ny uke
							
							strSQL = "SELECT overtidType= "&_
								" CASE Loennsartnr "&_
									" WHEN 160 THEN 35 "&_
									" WHEN 163 THEN 70 "&_
								" END, "&_
								" OppdragID, fakturatimer, fakturapris, fakturabeloep "&_
								" FROM VIKAR_UKELISTE "&_
								" WHERE  firmaid="& firmaNR &_
								" AND Loennsartnr in(160,163)"&_
								" AND Ukenr = "& sluttUke &_
								" ORDER BY Loennsartnr "
							set rsOvertid = GetFirehoseRS(strSQL, Conn)
						
							if HasRows(rsOvertid) THEN 'hvis det er overtid registrert
								do until rsOvertid.EOF
									sumTotal = sumTotal + rsOvertid("fakturabeloep")
									rsOvertid.MoveNext
								Loop								
								rsOvertid.close
							end if  'hvis det er overtid registrert
							set rsOvertid = nothing
						end IF 'ny uke
						%>
						<tr>
							<td>Sum total for hele måneden inkludert eventuell overtid</td>
							<td class="right"><%=formatNumber(sumTotal, 2)%></td>
						<%
						SumAvdeling = SumAvdeling + sumTotal
						sumTotal = 0
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
						<tr>
							<td colspan="2"><hr></td>
						</tr>
						<tr>
							<td>Sum total for avdeling <strong><% =avdeling%></strong> hele måneden inkludert eventuell overtid</td>
							<td class="right"><%=formatNumber(sumAvdeling,2)%></td>
							<%
						Loop 'hovedloop slutt
						%>
						</tr>
						<tr>
							<td>Sum total for alle avdelinger hele måneden inkludert eventuell overtid</td>
							<td class="right"><%=formatNumber(sumAlle,2)%></td>
						</tr>
						<%
						rsOppdrag.close
						set rsOppdrag = nothing
					end if 'hvis det er registrert oppdrag i perioden
					%>
				</table>
			</table>
		</div>
	</body>
</html>

