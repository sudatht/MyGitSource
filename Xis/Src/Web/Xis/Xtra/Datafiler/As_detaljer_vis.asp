<html>
<head>
	<title>AS-detaljer</title>
	<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
</head>
<body>
	<div class="pageContainer" id="pageContainer">
<!--#INCLUDE FILE="../includes/Library.inc"--> 
<%

'se session("ASTildato")
'--------------------------------------------------------------------------------------------------
' prosessing parameters
'--------------------------------------------------------------------------------------------------

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")


strVikarID = Request("VikarID")
If Request("tildato") = "" Then
	tildato = session("AStildato")
Else
	tildato = Request("tildato")
End If

minOppdateringsDato = Request("minOppdateringsDato")
If minOppdateringsdato = "" Then minOppdateringsDato = tildato
'--------------------------------------------------------------------------------------------------
' søkeresultat
'--------------------------------------------------------------------------------------------------

	SQL = "select D.dato, D.VikarID, D.OppdragID, V.Etternavn, V.Fornavn, D.Fakturastatus, D.Loennstatus" &_
	", D.FirmaID, stat=D.TimelisteVikarStatus, F.Firma, D.Fakturadato" &_
	", D.Timelonn, D.AntTimer, Sum1=(D.Timelonn * D.AntTimer), splittuke" &_ 
	" from  DAGSLISTE_VIKAR D, VIKAR V, FIRMA F" &_
	" where D.TimelisteVikarStatus < 6"  &_
	" and D.VikarID = " & strVikarID &_
	" and D.Loennstatus < 3 " &_
	" and V.TypeID = 3" &_
	" and D.VikarID = V.VikarID" &_
	" and D.Dato < " & DbDate(tildato)  &_
	" and D.FirmaID = F.FirmaID" &_
	" order by V.Etternavn"

'se SQL

set rsAS = conn.execute(sql)

If Not rsAS.EOF Then

'--------------------------------------------------------------------------------------------------
' kolonner
'--------------------------------------------------------------------------------------------------
%>
<% =rsAS("VikarID") %>&nbsp;&nbsp;<% =rsAs("Etternavn") & " " & rsAS("Fornavn") %></A>
<table cellpadding='0' cellspacing='0'>
	<tr>
		<th>OppdragID</th>
		<th>Arbeidssted</th>
		<th>Dato</th>
		<th>Timer</th>
		<th>Sats</th>
		<th>Lønn</th>
		<th>T.stat</th>
		<th>F.stat</th>
		<th>F.dato</th>
		<th>Ukenr</th>
		<th>L.stat</th>
	</tr>
<%
'--------------------------------------------------------------------------------------------------
' søkeresultat
'--------------------------------------------------------------------------------------------------
FID = ""
tildato2 = tildato
tilogmedDatoTreff = False
antIkketreff = 0 'telle oppdager etter mino
do while not rsAS.EOF %>

<!------------------Siste felt og datogrense------------------------->
	<% 
	ukeNrNeste = Datepart("ww", rsAs("Dato"),2)
	ukeSplittNeste = rsAs("splittuke")
	datoNeste = rsAS("Dato")

	'prøver å finne oppdateringsgrense...
	'If FormatDateTime(datoDenne,2) < FormatDateTime(minOppdateringsDato,2) Then
	If fStatusDenne = 3 And CDate(datoDenne) <= CDate(minOppdateringsDato) Then
		color ="#00FFFF"
		treff = True
	Else
		color="white"
		treff = False
	End If 

	'Hvis vi har nådd ukegrense (enten splittet eller ny uke) 
	'skal datoDenne ("tilogmeddato") beholdes og kun skiftes ut med ny hvis hele neste
	'uke(del) har fakturastatus=3 (pga. ukesvis lagring i vikar_ukeliste)
	If (ukeSplittNeste = 1 Or ukeNrNeste <> ukeNrDenne) And fStatusDenne = 3 And Treff Then
		oppdateringsDato = datoDenne
		If splittuke = 1 Then 
			splittuke = 2
		Else
			If ukeSplittNeste = 1 Then splittuke = 1 Else splittuke = Null
		End If
		'skiver ut splittuke (gjelder også for dagene før...)
		If Not IsNull(splittuke) Then se0 "-" & splittuke 
		If CDate(datoNeste) > CDate(minOppdateringsDato) Then
			tilogmedDatoTreff = True
			infokode = 1
			ukenr = ukeNrDenne  
		Else
			tilogmedDatoTreff = False		
			infokode = 2
		End If	
	End If

	If FID <> "" Then 'ikke første gangen %>
		<TD BGCOLOR=<% =color %>><% =lStatusDenne %>
	<% End If %>

<!-------------------------Start på listing-------------------------->
	<% 'Skriver ut først og for hvert firma...
	If FID <> rsAs("FirmaID") Then 
		FID = rsAs("FirmaID") %>
		<tr>
		<TD><% =rsAs("OppdragID") %>
		<TD><% =rsAs("Firma") %>
	<% Else %>
		<tr>
		<TD colspan=2>
	<% End If %>		

	<TD><% =rsAs("Dato") %>
	<TD><% =rsAs("AntTimer") %>
	<TD><% =rsAs("timelonn") %>
	<TD class=right><% =rsAs("Sum1") %>
	<TD><% =rsAs("stat") %>
	<TD><% =rsAs("Fakturastatus") %>
	<TD><% =rsAs("Fakturadato") %>
	<%
	fStatusDenne = rsAs("Fakturastatus")
	lStatusDenne = rsAs("Loennstatus")
	datoDenne = rsAs("Dato")
	ukeNrDenne = Datepart("ww", rsAs("Dato"),2)
	sum = sum + rsAs("Sum1")
	If rsAs("Fakturastatus") = 3 Then
		tildato2 = rsAS("dato")
	End If %>
	<TD><% =ukeNrDenne %>
	<% rsAs.MoveNext
loop
rsAS.Close: Set rsAS = Nothing

'utskrift av siste felt...
If fStatusDenne = 3 And CDate(datoDenne) <= CDate(minOppdateringsDato) Then
	color ="#00FFFF"
	If (CDate(session("AStildato")) - CDate(datoDenne)) > 6 Then
		ukenr = ukeNrDenne  
		oppdateringsDato = datoDenne
		tilogmedDatoTreff = True
	End If
Else
	color="white"
End If %>
<TD BGCOLOR=<% =color %>><% =lStatusDenne %>

<!--------------------------Lønnssum-------------------------->
<tr>
<th colspan=5><th><% =sum %>
</table>
<!--------------------------Dato resultat-------------------------->
<%
Select Case infokode
	Case 1
		info = "Oppdateringsdato samsvarer med ukeslutt/delukeslutt og fakturastatus = 3."
	Case 2 
		info = "Det er ikke samsvar mellom foreslått oppdateringsdato, fakturastatus=3 og ukeslutt. Trykk 'Juster oppdateringsdato'."
	Case Else
End Select %>	
Status:&nbsp;&nbsp;<% =info %>
<br><br>
<!--------------------------Juster tildato-------------------------->
<table cellpadding='0' cellspacing='0'>
<tr>
<FORM ACTION="AS_detaljer_vis.asp?VikarID=<% =strVikarID %>&tildato=<% =tildato %>" METHOD=POST>
<th>Forslått oppdateringsdato:
<TD><% =minOppdateringsDato %>
<th>Juster oppdateringsdato:
<TD><INPUT SIZE=6 TYPE=TEXT NAME=minOppdateringsDato VALUE="<% =oppdateringsDato %>">
<TD><INPUT TYPE=SUBMIT VALUE="     Hent    "></TD>
</table>
</form>
<% If tilogmedDatotreff Then %>
<!--------------------------Oppdater statuser------------------------->
<br>
<table cellpadding='0' cellspacing='0'>
<tr>
<FORM ACTION="AS_oppdater_db.asp?VikarID=<% =strVikarID %>&splittuke=<% =splittuke %>&ukenr=<% =ukenr %>" METHOD=POST>
<th>Sett lønnsstatus = 3 (OBS! Kan ikke angres!) tilogmeddato:
<TD><INPUT SIZE=6 TYPE=TEXT NAME=oppdateringsDato VALUE="<% =oppdateringsDato %>" READONLY>
<TD><INPUT TYPE=SUBMIT VALUE="      Utfør     "></TD>
</table>
</form><!--------------------------Bunntekst------------------------->
Kontroller at faktura er mottat og at den er utbetalt!
<br>
<%
End If 'tilogmedDatoTreff

Else 'ingen rader
	se "Obs! Han er vekk!"
End If 'rader
%>
<br>|
<A HREF="AS_vis.asp?ASTildato=<% =session("ASTildato") %>" >Hent nytt AS</A>
|
    </div>
</body>
</html>





