<html>
<head>
	<title>Faktura</title>
	<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	<!--#INCLUDE FILE="../Includes/Library.inc"--> 
</head>
<body>
	<div class="pageContainer" id="pageContainer">
<%
'seVar
'--------------------------------------------------------------------------------------------------
' parametere
'--------------------------------------------------------------------------------------------------

'seVar

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

'--------------------------------------------------------------------------------------------------
' SQL for alle fakturaer på denn kontakten (etter dato)
'--------------------------------------------------------------------------------------------------

sql = "select Distinct Fakturadato from FAKTURAGRUNNLAG " &_
	"where Kontakt = " & Request("Kontakt")  &_
	" and Avdeling = " & Request("AvdelingID") &_
	" and Status = 3" &_
	" order by Fakturadato desc"
'se sql

set rsDato = conn.execute(sql)

If Request("Fakturadato") = "" Then Fakturadato = rsDato("Fakturadato") Else Fakturadato = Request("Fakturadato")


'--------------------------------------------------------------------------------------------------
' SQL     
'--------------------------------------------------------------------------------------------------
strSQL = "select Linje=FakturaLinjeID, Kontakt, FirmaID, OppdragID, ArtikkelNr, VikarID, " &_
	"Tekst, Antall, Status, Enhetspris, LinjeSum, LinjeNr, Split, NyLinje " &_
	", Fakturanr, Fakturadato " &_
	"from FAKTURAGRUNNLAG " &_
	"where Kontakt = " & Request("Kontakt")  &_
	" and Avdeling = " & Request("AvdelingID") &_
	" and Fakturadato = " & dbDate(fakturadato) &_
	" order by VikarID, LinjeNr"

'se strSQL & "<br>"

Set rsFaktLinjer = conn.Execute(strSQL)

If rsFaktLinjer.EOF Or rsFaktLinjer.BOF Then
	se "Ingen fakturalinjer for " & strFirmaNavn & ",   " & strKontaktNavn
Else

'--------------------------------------------------------------------------------------------------
' SQL for finding firma and kontakt
'--------------------------------------------------------------------------------------------------

strSQL = "select Navn=(Fornavn + ' ' + Etternavn) from KONTAKT where KontaktID = " & Request("Kontakt")
'se strSQL & "<br>"
Set rsNavn = Conn.Execute(strSQL)
strKontaktNavn = rsNavn("Navn")
rsNavn.Close: Set rsNavn = Nothing

strSQL = "select Firma from Firma where FirmaID = " & Request("FirmaID")
'se strSQL & "<br>"
Set rsNavn = Conn.Execute(strSQL)
strFirmaNavn = rsNavn("Firma")
rsNavn.Close : Set rsNavn = Nothing


'--------------------------------------------------------------------------------------------------
' Display data
'--------------------------------------------------------------------------------------------------

status = rsFaktLinjer("Status")

%>
<table cellpadding='0' cellspacing='0'>
<tr><th colspan=6><% =Fakturadato %> - <% =strFirmanavn %>, <% = strKontaktNavn %> 
<tr>
<th>Artikkelnr
<th>Tekst
<th>Antall
<th>Pris
<th>Sum
<%

VID = rsFaktLinjer("VikarID")
VikarerHer = VID
splitOption = False
faktNr = rsFaktLinjer("Fakturanr")
sumsum = 0
'--------------------------------------------------------------------------------------------------------------------------------
do while NOT rsFaktLinjer.EOF
'--------------------------------------------------------------------------------------------------------------------------------


If rsFaktLinjer("VikarID") <> VID Then
	splitOption = True
	VID = rsFaktLinjer("VikarID")
	VikarerHer = VikarerHer & "," & VID
	If rsFaktLinjer("split") = 1 Then %>
	<tr><TD><TD><th class="right">Sum<TD><TD><th class="right"><% =sumsum %><% sumsum = 0 %>
	<tr><tr><tr>
	<tr><th colspan=8><% =strFirmanavn %>, <% = strKontaktNavn %> 
	<tr><th>Endre<th>Artikkelnr<th>Tekst<th>Antall<th>Pris<th>Sum<th>Slett
	<TD><A HREF="Faktura_vis.asp?Endre=Insert&Kontakt=<% =strKontaktID %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&LinjeNr=0&VikarID=<% =rsFaktLinjer("VikarID") %>" >Settinn </A>
<% 'sumsum = rsFaktLinjer("LinjeSum")	
	End If
End If

If Not IsNull(rsFaktLinjer("LinjeSum")) Then
	sumsum = sumsum + rsFaktLinjer("LinjeSum")
End If

%>


<tr>
<TD><% =rsFaktLinjer("ArtikkelNr") %>
<TD><% =rsFaktLinjer("Tekst") %>
<TD class=right><% =rsFaktLinjer("Antall") %>
<TD class=right><% =rsFaktLinjer("Enhetspris") %>
<TD class=right><% =rsFaktLinjer("LinjeSum") %>
<TD>
<%

delt = rsFaktLinjer("split")

rsFaktLinjer.MoveNext

'--------------------------------------------------------------------------------------------------------------------------------
loop
'--------------------------------------------------------------------------------------------------------------------------------
%>
</table><br>
<%
'--------------------------------------------------------------------------------------------------------------------------------
' andre fakturaer
'--------------------------------------------------------------------------------------------------------------------------------

teller = 0
do while not rsDato.EOF

	teller = teller + 1
	If teller = 6 Then se " |": teller = 1
	If CStr(rsDato("Fakturadato")) <> CStr(Fakturadato) Then
		link = "Faktura_vis_gml.asp" &_
			"?Fakturadato=" & rsDato("Fakturadato") &_
			"&AvdelingID=" & Request("AvdelingID") &_ 
			"&Kontakt=" & Request("Kontakt") &_
			"&FirmaID=" & Request("FirmaID") 
		se0 "| <A HREF=" & link & ">" & rsDato("Fakturadato")& "</A>"
	Else
		se0 "| " & rsDato("Fakturadato")
	End If
	rsDato.MoveNext
loop
se " |"
End If ' ingen rader

%>
    </div>
</body>
</html>
