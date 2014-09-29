<html>
<head>
	<title>Arbeid med datafiler</title>
	<!--#INCLUDE FILE="../includes/Library.inc"-->
	<link rel="stylesheet" href="/xtra/css/
main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/
print.css" type="text/css" title="xtra intranett stylesheet" media="print">
</head>
<body>
	<div class="pageContainer" id="pageContainer">

<% 
kkk= Request.QueryString("Notat") 
   	p = Len(kkk) - 2
If p > 1  Then 
	kk = Mid(kkk,2, p)
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

If Request.QueryString("VikarID") = "" Then
	strVikarID = Request.Form("VikarID")
	strNavn = Request.Form("Navn")
	strOppdragID = Request.Form("OppdragID")
	strLinje = Request.Form("Linje")
	strFirmaId = Request.Form("FirmaID")
'	tilgang = Request.Form("tilgang")
Else
	strVikarID = Request.QueryString("VikarID")
	strNavn = Request.QueryString("Navn")
	strOppdragID = Request.QueryString("OppdragID")
	strLinje = Request.QueryString("Linje")
	strFirmaId = Request.QueryString("FirmaID")
'	tilgang = Request.QueryString("tilgang")
End IF
	strStartDato = Request.QueryString("StartDato")
	strAntTimer = Request.QueryString("AntTimer")	
status_kode = Request.QueryString("kode")

'tilgang = CInt(tilgang)
tilgang = Session("tilgang")

strTimeLonn = Request.QueryString("Timelonn")
strFakturapris = Request.QueryString("Fakturapris")
strNotat = Request.QueryString("Notat")
strBestilltav = Request.QueryString("Bestilltav")

strOppdragID =  Request.QueryString("OppdragID")

strUkeNr = Datepart("ww", session("limitDato"),2)

'--------------------------------------------------------------------------------------------------
' check parameters
'--------------------------------------------------------------------------------------------------
'Response.write DateValue(Date) & "<br>"
'Response.write OnlyDigits(Date) & "<br>"
'Response.write strVikarID & "<br>"
'Response.write strNavn & "<br>"
'Response.write strOppdragID & "<br>"
'Response.write strFirmaID & " " & strFirma & "<br>"
'Response.write strLinje & "<br>"
'Response.write strStartDato & "<br>"
'Response.write strAntTimer & "<br>"
'Response.write Request.QueryString("Dato") & "<br>"
'Response.write strTimelonn & "<br>"
'Response.write status_kode & "<br>"
'Response.write tilgang & "t<br>"
'Response.write Request.QueryString("tilgang") & "q<br>"
'Response.write Request.Form("tilgang") & "f<br>"
'Response.write Request.Querystring("Endre")
'Response.write strTimelonn & "<br>"
'Response.write strFakturapris & "<br>"
'Response.write strNotat & "<br>"
'Response.write strBestilltav & "<br>"
'Response.write session("limitdato") & "<br>"
'Response.write strUkeNr & "<br>"


'--------------------------------------------------------------------------------------------------
' SQL for displaying data
'--------------------------------------------------------------------------------------------------
'status_kode = session("status_kode")


strSQL = "select ID, V.VikarId, VNavn=V.Fornavn + ' ' + V.Etternavn, OppdragID, FirmaID, UkeNr, Loennsartnr, " &_
	"Antall, Sats, Belop, U.Notat, Dato, U.StatusID, BestilltAv, Fakturapris, Fakturabeloep " &_ 
	 "from VIKAR_UKELISTE U, VIKAR V " &_
	"where U.VikarID = V.VikarID " &_
	" and BestilltAv = " & strBestilltAv &_
	" and U.StatusID > 3 " &_
	" order by VNavn, UkeNr"


'Response.write strSQL & "<br>"

Set rsTimeliste = Conn.Execute(strSQL)


'--------------------------------------------------------------------------------------------------
' If no record exsists
'--------------------------------------------------------------------------------------------------

If rsTimeliste.BOF= True And rsTimeliste.EOF = True Then 


	Response.write "<H4><br>Ingen klargjorte!<br></H4>"


Else	'if no record 

'--------------------------------------------------------------------------------------------------
' SQL for finding name
'--------------------------------------------------------------------------------------------------

strSQL = "select Kontaktnavn=(Fornavn + ' ' + Etternavn) from Kontakt where KontaktID = " & rsTimeliste("Bestilltav")

'Response.write strSQL & "<br>"

Set rsNavn = Conn.Execute(strSQL)

strKontaktNavn = rsNavn("Kontaktnavn")

rsNavn.Close


'--------------------------------------------------------------------------------------------------
' SQL for firma name
'--------------------------------------------------------------------------------------------------
strFirmaID = rsTimeliste("FirmaID")
'Response.write strFirmaID

strSQL = "select Firma from Firma where FirmaID = " & rsTimeliste("FirmaID")

Set rsFirma = Conn.Execute(strSQL)

strFirma = rsFirma("Firma")

rsFirma.Close

'--------------------------------------------------------------------------------------------------
' Find week number
'--------------------------------------------------------------------------------------------------
If strStartDato = "" Then
	strStartDato = rsTimeliste("Dato")	
End If
strWeekNumber = Datepart("ww", strStartDato,2)


'--------------------------------------------------------------------------------------------------
' Form to register and change
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' If edit find the right row to display
'--------------------------------------------------------------------------------------------------


If strLinje <> "" Then ' edit

 strSQL = "select Dato, Starttid, Sluttid " &_ 
	"from DAGSLISTE_VIKAR " &_
	 "where TimelisteVikarID = " & strLinje

	Set rsLinje = conn.Execute(strSQL)
strStarttid = Left(TimeValue(rsLinje("Starttid")),5)
strSluttid = Left(TimeValue(rsLinje("Sluttid")),5)
strUkedag = WeekDayName(Weekday(rsLinje("Dato"),2),,2)

End If 	

'--------------------------------------------------------------------------------------------------
' Form to register and change
'--------------------------------------------------------------------------------------------------
Response.write "<br><H4>" & strFirma & ", " & strKontaktnavn & "</H4>"

If Request.QueryString("Endre") <> "" Then

 kk = strStartDato
 pp = Request.QueryString("Dato")
'Response.write kk & "s<br>"
'Response.write pp & "d<br>"

 If kk = pp Then
	kk = ""
End If

%>

<table cellpadding='0' cellspacing='0'>

<FORM NAVN="OPPDRAG" ACTION="Vikar_timeliste_oppd_db2.asp" METHOD=POST>
<input name="VikarID" TYPE=HIDDEN VALUE=<% =strVikarID %> >
<input name="OppdragID" TYPE=HIDDEN VALUE="<% =strOppdragID %>" >
<input name="FirmaID" TYPE=HIDDEN VALUE="<% =strFirmaID %>" >
<input name="StartDato" TYPE=HIDDEN VALUE="<% =kk %>" >
<input name="Endre" TYPE=HIDDEN VALUE=<% =Request.Querystring("Endre") %>>
<input name="Linje" TYPE=HIDDEN VALUE="<% =strLinje %>">
<input name="Bestilltav" TYPE=HIDDEN VALUE="<% =strBestilltav %>">

<tr><th>Ukedag<th>Dato<th>Fratid<th>Tiltid<th>Timelønn<th>Fakt.pris</th>

<TR	>
<th><% If strLinje <> "" Then Response.write strUkedag Else Response.write "" %></th>
<th><input type=text size=8 NAME=Dato <% If strLinje <> "" Then Response.write "VALUE=" & Request.QueryString("Dato") %> ></th>
<th><input type=text size=6 NAME=Starttid <% If strLinje <> "" Then Response.write "VALUE=" & strStarttid %> ></th>
<th><input type=text size=6 NAME=Sluttid <% If strLinje <> "" Then Response.write "VALUE=" & strSluttid %> ></th>
<th><input type=text size=6 NAME=Timelonn <% If strLinje <> "" Then Response.write "VALUE=" & strTimelonn %> ></th>
<th><input type=text size=6 NAME=Fakturapris <% If strLinje <> "" Then Response.write "VALUE=" & strFakturapris %> ></th>
<tr><th colspan=7 ><input type=text size=53 NAME=Notat <% If strLinje <> "" Then Response.write "VALUE=" & strNotat %> ></th>
<tr>
<th colspan=7 ><INPUT TYPE=SUBMIT	VALUE="                           Lagre                                   "><INPUT TYPE=RESET ></th>
</form></table>

<% End If  'endre eller ny %>
<% If strLinje <> "" Then rsLinje.Close %>

<%
'--------------------------------------------------------------------------------------------------
' Display data
'--------------------------------------------------------------------------------------------------
%>

<table cellpadding='0' cellspacing='0'>

<TR class=right>
<th>ID</th>
<th>Ukenr</th>
<th>Lønnsartnr</th>
<th>Timel</th>
<th>Antall</th>
<th>Sumlønn</th>
<th>lønnStatus</th>
<th>Fakt.gr.</th>
<th>Faktsum</th>
<th>Vikar</th>

<%
nnn = rsTimeliste("VNavn")


 do while not rsTimeliste.EOF

If nnn <> rsTimeliste("VNavn") Then
	nnn = rsTimeliste("VNavn")
%>
<tr><TD><TD><TD><TD><TD><th><% =sumLoenn %>
<TD><TD><th><% =sumFakt %>
<%
	sumsumLoenn = sumsumLoenn + sumLoenn
	sumsumFakt = sumsumFakt + sumFakt
	sumLoenn = rsTimeliste("Belop")
	sumFakt = rsTimeliste("Fakturabeloep")
Else
	sumLoenn = sumLoenn + rsTimeliste("Belop")
	sumFakt = sumFakt + rsTimeliste("Fakturabeloep")

End If

%>

<tr>
<TD><% =rsTimeliste("ID") %></TD>
<TD><% =rsTimeliste("Ukenr") %>
<TD><% =rsTimeliste("Loennsartnr") %>
<TD><% =rsTimeliste("Sats") %>
<TD><% =rsTimeliste("Antall") %>
<TD><% =rsTimeliste("Belop") %>
<th><% If rstimeliste("StatusID") = "5" Then %><FONT COLOR=GREEN><%  ElseIf rsTimeliste("StatusID") = "1" Then %> <FONT COLOR=RED><% Else %><FONT COLOR=YELLOW><% End If %>
(<% =rsTimeliste("StatusID") %>)
<TD><% =rsTimeliste("Fakturapris") %>
<TD><% =rsTimeliste("Fakturabeloep") %>
<TD><% =rsTimeliste("VNavn") %>

<% 
rsTimeliste.MoveNext
loop

%>
<tr><TD><TD><TD><TD><TD><th><% =sumLoenn %>
<TD><TD><th><% =sumFakt %>
<tr>
<% sumsumLoenn = sumsumLoenn + sumLoenn %>
<tr><TD><TD><TD><TD><TD><th><% =sumsumLoenn %>
<TD><TD><th><% =sumsumFakt %>

</table>
<%
'--------------------------------------------------------------------------------------------------
' Prepare for Loop
'--------------------------------------------------------------------------------------------------
rsTimeliste.MoveFirst

antRecord = 0
antLoop = 0
neste = False

If Request.QueryString("startDato") = "" Then
 	strStartDato = rsTimeliste("Dato") 
Else
 	strStartDato = Request.QueryString("startDato")
End If

ukeTeller = 0
nowWeek =  0 

'--------------------------------------------------------------------------------------------------
' Loop through recordset
'--------------------------------------------------------------------------------------------------
 do while not rsTimeliste.EOF
	antRecord = antRecord + 1
	rstimeliste.MoveNext	
 loop	
 rsTimeliste.MoveFirst
	

' do while not rsTimeliste.EOF 

	antLoop = antLoop + 1

'	If CStr(rsTimeliste("Dato")) = CStr(strStartDato) Then
'		neste = True
'		displayWeek  = Datepart("ww", rsTimeliste("Dato"),2)
'	End If
'	dag = WeekDayName(Weekday(rsTimeliste("Dato"),2),,2)

'--------------------------------------------------------------------------------------------------
' Determin when to display rows (weeks before are collected at the bottom of loop)
' Display until sunday else (create buttons for stored weeks ) and create buttons for next weeks 
'--------------------------------------------------------------------------------------------------
If neste Then

     If displayWeek = Datepart("ww", rsTimeliste("Dato"),2) Then
	    strStarttid = Left(TimeValue(rsTimeliste("Starttid")),5)
		strSluttid = Left(TimeValue(rsTimeliste("Sluttid")),5)
		strAnttimer = rsTimeliste("AntTimer")

%>

<TR class=right>

<!---------------Link to UPDATE---------------------->
<% update="Endre=Ja&VikarID=" & rsTimeliste("VikarID") & "&OppdragID=" & strOppdragID & "&Linje=" & rsTimeliste("Linje") &_
                       "&FirmaID=" & strFirmaID & "&Fratid=" & strStarttid & "&Tiltid=" & strSluttid &_
	     "&Dato=" & rsTimeliste("Dato") & "&AntTimer=" & strAnttimer & "&StartDato=" & strStartDato &_
                       "&kode=" & status_kode & "&Timelonn=" & rsTimeliste("Timelonn") &_
	     "&Fakturapris="  & rsTimeliste("FakturaPris") & "&Notat='" & rsTimeliste("Notat") & "'&Bestilltav=" & strBestilltav
%>

<th><A HREF="Vikar_timeliste_fakt_vis.asp?<% =update %>" >Endre</A></th>

<th><% =dag %></th>
<th><%=DateValue(rsTimeliste("Dato"))%></th>
<th><%=strStarttid %></th>
<th><%=strSluttid %></th>
<th><%=strAnttimer %></th>
<% name=CStr(OnlyDigits(rsTimeliste("Dato"))) %>
<th><% =rsTimeliste("Timelonn") %>
<th><% =rsTimeliste("FakturaPris") %>
<th><% =rsTimeliste("Status2") %>
<th><% If rstimeliste("Status1") = "5" Then %><FONT COLOR=GREEN><%  Else %> <FONT COLOR=RED><% End If %>
(|)

<!------------------Link to DELETE-------------------->
<th><A HREF=Vikar_timeliste_fakt_db.asp?Linje=<%=rsTimeliste("Linje") %>&Slett=Ja&VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&StartDato=<% =strStartDato %>&FirmaID=<% =strFirmaID %>&kode=<% =status_kode %>&Dato=<%=DateValue(rsTimeliste("Dato"))%>&Bestilltav=<% =strBestilltav %> >Slett</A>

<th><% =rsTimeliste("Vikarnavn") %>
<th><% =rsTimeliste("Notat") %>

<%

	Else	
		neste = False
 	End If 

 End If	 'første dato er startDato, visning av uke

  rsTimeliste.MoveNext

'loop 

'--------------------------------------------------------------------------------------------------
' End of loop 
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Display more buttons 
'--------------------------------------------------------------------------------------------------
%>
<br>
<table cellpadding='0' cellspacing='0'>
<FORM ACTION="Faktura_vis.asp?OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&Kontakt=<% =strBestilltAv %>" METHOD=POST>
<tr><th>
<INPUT TYPE=SUBMIT VALUE="            Fakturalinjer            "></th>
</table>

<%  End If 	'no rows %>

<table cellpadding='0' cellspacing='0'>
</form>
<FORM ACTION="Vikar_timeliste_fakt_list.asp?viskode=<% =session("viskode") %>&dato1=<% =session("limitDato") %>" METHOD=POST>
<tr><th>
<INPUT TYPE=SUBMIT VALUE="                Tilbake                 "></th>
</table>
</form><%

 rsTimeliste.Close


session("sluttdato") = sluttdato 
session("startdato") = strStartdato

%>



</body>
</html>


