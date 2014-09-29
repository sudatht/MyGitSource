<html>
<head>
	<!--#INCLUDE FILE="../includes/Library.inc"-->
	<title>Timeliste</title>
	<link rel="stylesheet" href="/xtra/css/
main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/
print.css" type="text/css" title="xtra intranett stylesheet" media="print">
</head>
<body>
	<div class="pageContainer" id="pageContainer">
<%
Dim Avd

Function fjernEtterKomma(tall)

  Pos1 = InStr(1, tall,",")
  If pos1 > 0 Then
	tall = Left(tall, pos1 - 1)
  End If

  fjernEtterKomma = tall

End Function

'-----------------------------------------------------------------------------------------------
' Connect to database
'-----------------------------------------------------------------------------------------------

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

'-----------------------------------------------------------------------------------------------
' prosessing parameters
'-----------------------------------------------------------------------------------------------

fradato = Request("fradato")
tildato = Request("tildato")
sammendrag = Request("sammendrag")

'-----------------------------------------------------------------------------------------------
' sql to find data
'-----------------------------------------------------------------------------------------------

strSQL = "select D.VikarID, OPPDRAG.AvdelingID, " &_
	"Lønn=Sum(Timelonn * AntTimer), " &_
	"Faktgr=Sum(Fakturatimer * Fakturapris) " &_
	"from DAGSLISTE_VIKAR D, OPPDRAG " &_
	"where OPPDRAG.OppdragID = D.OppdragID " &_
	"and D.Dato >= " & dbDate(fradato) &_
	" and D.Dato <= " & dbDate(tildato) &_
	" group by OPPDRAG.AvdelingID, D.VikarID"
	

Set rsOmsetning = conn.Execute(strSQL)

If rsOmsetning.EOF Then
	Response.write "Noe er galt!!"
Else


'-----------------------------------------------------------------------------------------------
' show data
'-----------------------------------------------------------------------------------------------
Headding = "<H4>Omsetningstall for perioden " & fradato & " - " & tildato & "</H4>"
Response.write Headding

AVD = 0
sumLoen = 0
sumFakt = 0
sumsumLoen = 0
sumsumFakt = 0

%>


<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If sammendrag = "ja" Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>
<table cellpadding='0' cellspacing='0'>
<tr><th>Avdeling<th>Lønnsutbetaling<th>Faktureringsgrunnlag</th>
<%
do while Not rsOmsetning.EOF

loen = fjernEtterKomma(rsOmsetning("Lønn"))
fakt = fjernEtterKomma(rsOmsetning("Faktgr"))

AvdID = rsOmsetning("AvdelingID")

If AvdID <> AVD Then
	AVD = AvdID

	If sumLoen <> 0 Then
	%>
	<tr><th><% =avdnavn %><th class="right"><% =sumLoen %><th class="right"><% =sumFakt %></th>
	<%
	sumsumLoen = sumsumLoen + sumLoen
	sumsumFakt = sumsumFakt + sumFakt
	sumLoen = 0
	sumFakt = 0

	End If 'sumLoen <> 0 

	avdnavn = UCAse(HentAvdNavn(Avd))
	'Select Case AVD
		 'Case 1 
	'avdnavn = "KURS"
		 'Case 2 
	'avdnavn = "DATA" 
	'	Case 3 
	'avdnavn = "DOKUMENT" 
	'	Case Else 
	'End Select 
	
End If 'AVD <> Avdeling

sumLoen = sumLoen + loen
sumFakt = sumFakt + fakt


rsOmsetning.MoveNext
loop

sumsumLoen = sumsumLoen + sumLoen
sumsumFakt = sumsumFakt + sumFakt

%>
<tr><th><% =avdnavn %><th class="right"><% =sumLoen %><th class="right"><% =sumFakt %></th>
<tr><tr>
<tr><th>Sum alle avd.<th class="right"><% =sumsumLoen %><th class="right"><% =sumsumFakt %></th>
<tr><tr>
<th>Differanse<th class="right"><% =sumsumFakt - sumsumLoen %>


</table>
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Else 'sammendrag <> ja
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

do while Not rsOmsetning.EOF

loen = fjernEtterKomma(rsOmsetning("Lønn"))
fakt = fjernEtterKomma(rsOmsetning("Faktgr"))

AvdID = rsOmsetning("AvdelingID")

If AvdID <> AVD Then
AVD = AvdID

	If sumLoen <> 0 Then
	%>
	<tr><th colspan=2 ><% =avdnavn %><th class="right"><% =sumLoen %><th class="right"><% =sumFakt %></th>
	<%
	sumsumLoen = sumsumLoen + sumLoen
	sumsumFakt = sumsumFakt + sumFakt
	sumLoen = 0
	sumFakt = 0

	End If 'sumLoen <> 0 

	avdnavn = HentAvdNavn(Avd)%>
	<table cellpadding='0' cellspacing='0'><tr><tr><th colspan=4>
	Avdeling: <% = UCase(Avdnavn) %>

	<tr><th>Avdeling<th>Vikarnr<th>Lønnsutbetaling<th>Fakt.gr</th>
	<%
End If 'AVD <> Avdeling

sumLoen = sumLoen + loen
sumFakt = sumFakt + fakt


%>
<tr>
<TD><% =rsOmsetning("AvdelingID") %></TD>
<TD><% =rsOmsetning("VikarID") %></TD>
<TD class=right><% =loen %></TD>
<TD class=right><% =fakt %></TD>
<%

rsOmsetning.MoveNext
loop

SumsumLoen = sumsumLoen + sumLoen
SumsumFakt = sumsumFakt + sumFakt

%>
<tr><th colspan=2 ><% =avdnavn %><th class="right"><% =sumLoen %><th class="right"><% =sumFakt %></th>
<tr><tr>
<tr><th colspan=2 >Sum alle avd.<th class="right"><% =sumsumLoen %><th class="right"><% =sumsumFakt %></th>
<tr><tr>
<th colspan=2 >Differanse<th class="right"><% =sumsumFakt - sumsumLoen %>


</table>

<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If 'sammendrag = ja
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If 'no rows
rsOmsetning.Close
Set rsOmsetning = Nothing
%> 

</body>
</html>