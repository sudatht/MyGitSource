<!--#INCLUDE FILE="../Includes/Library.inc"--> 
<html>
<head>
	<title>Lønn</title>
	<link rel="stylesheet" href="/xtra/css/
main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/
print.css" type="text/css" title="xtra intranett stylesheet" media="print">
</head>
<body>
	<div class="pageContainer" id="pageContainer">
<%
'seVar
'--------------------------------------------------------------------------------------------------
' Connect to database
'--------------------------------------------------------------------------------------------------

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

'--------------------------------------------------------------------------------------------------
' prosessing parameters
'--------------------------------------------------------------------------------------------------

VikarID = Request("VikarID")
viskode = Request("viskode")
OppdragID = Request("OppdragID")
frakode = Request("frakode")


'--------------------------------------------------------------------------------------------------
' SQL for finding name
'--------------------------------------------------------------------------------------------------
strSQL = "select Navn=(Fornavn + ' ' + Etternavn) from VIKAR where Vikarid = " & VikarID

'se strSQL & "<br>"

Set rsNavn = Conn.Execute(strSQL)

Navn = rsNavn("Navn")

rsNavn.Close

'--------------------------------------------------------------------------------------------------
' liste over alle lønnsdatoer
'--------------------------------------------------------------------------------------------------
sql = "select distinct LoennDato from VIKAR_LOEN_VARIABLE " &_
	"where VikarID = " & VikarID &_
	" and NOT LoennDato IS NULL " &_
	"order by LoennDato desc"
'se sql

set rsLoenndato = conn.execute(sql)

If rsLoenndato.EOF Then
	se Navn & " Har ikke fått utbetalt lønn!"
Else

	If Request("loennDato") = "" Then
		loennDato = rsLoenndato("LoennDato")
	Else
		loennDato = Request("LoennDato")
	End If 

'--------------------------------------------------------------------------------------------------
' SQL 
'--------------------------------------------------------------------------------------------------

strSQL = "select VLV.Id, VLV.Dato, VLV.Prosjektnr, VLV.Loennsartnr, " &_
	"HL.Loennsart, VLV.Antall, VLV.Sats, VLV.Beloep, VLV.LoennDato, A.Avdeling" &_ 
	" from VIKAR_LOEN_VARIABLE VLV, H_LOENNSART HL, AVDELING A" &_
	" where VikarID = " & VikarID &_
	" and VLV.Loennsartnr = HL.Loennsartnr" &_
	" and A.avdelingID = VLV.Avdeling" &_
	" and LoennDato = " & dbDate(loennDato)



Set rsVikar = Conn.Execute(strSQL)


'--------------------------------------------------------------------------------------------------
' SQL for displaying data
'--------------------------------------------------------------------------------------------------

%>

<table cellpadding='0' cellspacing='0'>
<tr><th  colspan=7 >Variable lønn - <% =LoennDato %> - <% =VikarID %> - <% =Navn %>
<TR class=right>
<th WIDTH=20>Artnr
<th WIDTH=20>Artnavn
<th WIDTH=30>Antall
<th>Sats
<th WIDTH=50>Beløp
<th WIDTH=50>Avdeling

<% do while not rsVikar.EOF %>

	<TR class=right>
	<TD><% =rsVikar("Loennsartnr")%></TD>
	<TD><% =rsVikar("Loennsart")%></TD>
	<TD><% =rsVikar("Antall")%></TD>
	<TD><% =rsVikar("Sats")%></TD>
	<TD><% =rsVikar("Beloep")%></TD>
	<TD><% =rsVikar("Avdeling")%></TD>

<% 
rsVikar.MoveNext
loop
rsVikar.Close: Set rsVikar = Nothing
%>
</table>
<br>
<%
'--------------------------------------------------------------------------------------------------
' liste opp andre lønnsdatoer
'--------------------------------------------------------------------------------------------------
teller = 0
do while not rsLoennDato.EOF

	If CStr(rsLoennDato("LoennDato")) <> CStr(loennDato) Then %>

		|<A HREF="Vikar_varl_gml3.asp?VikarID=<% =VikarID %>&LoennDato=<% =rsLoennDato("LoennDato") %>&OppdragID=<% =OppdragID %>&viskode=<% =viskode %>&frakode=<% =frakode %>"><% =rsLoennDato("LoennDato") %></A>

	<% Else %>

		|<% =rsLoennDato("LoennDato") %>

	<% End If 
 	rsLoennDato.MoveNext
	teller = teller + 1
	If teller = 6 Then Response.write "<br>"
loop
rsLoennDato.Close: Set rsLoennDato = Nothing
%>
|<br><br>
<% 
'--------------------------------------------------------------------------------------------------
' Tilbake
'--------------------------------------------------------------------------------------------------
 'If frakode = 2 Then %>
	<!--TABLE BORDER=1><TR-->
	<!--FORM ACTION="Vikar_timeliste_list3.asp?Viskode=<% =viskode %>" METHOD=POST -->
	<!--TH><INPUT TYPE=SUBMIT VALUE="                  Tilbake                        "></th>
	</form>
	</table><BR!-->
<% 'End If 

'--------------------------------------------------------------------------------------------------
' Beskjeder
'--------------------------------------------------------------------------------------------------
If frakode < 4 Then 
strSQL = "select distinct NotatOkonomi " &_
	"from OPPDRAG " &_
	"where oppdragID = " & OppdragID

'se strSQl

Set rsBeskjed = conn.Execute(strSQL)
If Not rsBeskjed.EOF Then
	se "<br>Beskjeder:<br>"
	do while Not rsBeskjed.EOF
%>
		<% =rsBeskjed("NotatOkonomi") %><br>
<%
		rsBeskjed.MoveNext
	loop
	rsBeskjed.Close
	Set rsBeskjed = Nothing
End If 'no rows (beskjeder)
End If 'frakode < 4

End If 'ingen lønn

%>
</body>
</html>

