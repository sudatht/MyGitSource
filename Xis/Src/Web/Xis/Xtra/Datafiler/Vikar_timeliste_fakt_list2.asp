<html>
<head>
	<!--#INCLUDE FILE="../includes/Library.inc"-->
	<script type="text/javascript" language="javascript" src="../js/javascript.js"></script>
	<title>Arbeid med datafiler</title>
	<link rel="stylesheet" href="/xtra/css/
main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/
print.css" type="text/css" title="xtra intranett stylesheet" media="print">
</head>
<body>
	<div class="pageContainer" id="pageContainer">
<%
'--------------------------------------------------------------------------------------------------
' PARAMETERS
'--------------------------------------------------------------------------------------------------

Session("tilgang") = 3

If Request.QueryString("viskode") = "" Then
  viskode = session("viskode")
Else
  viskode = Request.QueryString("viskode")
  session("viskode") = viskode
End IF



mm = Datepart("m", Date): mm = mm + 1
yy = Datepart("yyyy", DateValue( Date))
If mm = 13 Then yy = yy + 1: mm= 1
If mm < 10 Then dd = "01.0" Else dd="01."
yy = Right(CStr(yy),2)
dato1 = dd & mm & "." & yy
dato2 = (Date - 60)

If Request.QueryString("Dato1") <> "" Then
	dato1 = Request.QueryString("Dato1")
End If
If Request.Form("Dato1") <> "" Then
	dato1 = Request.Form("Dato1")
End If
session("limitDato") = dato1

If Request.QueryString("Dato2") <> "" Then
 	dato2 = Request.QueryString("Dato2")
End If
If Request.Form("Dato2") <> "" Then
 	dato2 = Request.Form("Dato2")
End If
session("oldStartDate") = dato2

If Request.QueryString("OppdragID") = "" Then
	strOppdragID = Request.Form("OppdragID")
Else
	strOppdragID = Request.QueryString("OppdragID")
End If


'Response.write strOppdragID & "<br>"
'Response.write viskode & "<br>"
'Response.write Request.QueryString("tilgang") & "<br>"
'Response.write dato1 & "<br>"
'Response.write dato2 & "<br>"

'--------------------------------------------------------------------------------------------------
' BUTTONS
'--------------------------------------------------------------------------------------------------
If viskode =  1 Then k = "-> " Else k = "     "
If viskode =  2 Then kk = "-> " Else kk = "     "
If viskode =  3 Then kkk = "-> " Else kkk = "     "


%>

<FORM ACTION="Vikar_timeliste_fakt_list2.asp?viskode=1&dato2=<% =dato2 %>" METHOD=POST>
	<INPUT TYPE=SUBMIT VALUE="<% =k %>Ny fakturaer"><br>
	<input type=text size=7 NAME=DATO1 VALUE="<% =dato1 %>" ONBLUR="dateCheck(this.form, this.name)" >
</form>

<FORM ACTION="Vikar_timeliste_fakt_list2.asp?viskode=3&dato1=<% =dato1 %>" METHOD=POST>
	<INPUT TYPE=SUBMIT VALUE="<% =kkk %>Gamle"><br>
	<input type=text size=7 NAME=DATO2 VALUE=<% =dato2 %> ONBLUR="dateCheck(this.form, this.name)">
</form>
<% If viskode < 2 Or viskode = 3 Then %>

<FORM ACTION="Vikar_timeliste_ny3.asp?frakode=2" METHOD=POST >
	<p>Vikarnr: <input type="text" name=VIKARID SIZE=4 ></p>
	<p>Oppdragnr: <input type="text" name=OPPDRAGID SIZE=4 ></p>
	<INPUT TYPE=SUBMIT VALUE="Søk">
</form>
<% End If %>

<FORM ACTION="Vikar_timeliste_fakt_list2.asp?viskode=2&dato1=<% =dato1 %>&dato2=<% =dato2 %>" METHOD=POST>
	<INPUT TYPE=SUBMIT VALUE="<% =kk %> Klargjorte - overføring">
</form> 

<%
'--------------------------------------------------------------------------------------------------
' Connect to database
'--------------------------------------------------------------------------------------------------

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

'--------------------------------------------------------------------------------------------------
' SQL
'--------------------------------------------------------------------------------------------------
If viskode = 1 Then

strSQL = "select D.FirmaID, Firma, Navn=(Etternavn + ' ' + Fornavn), " &_
			"D.BestilltAv, stat1=Fakturastatus, maks=Max(TimelisteVikarStatus), " &_
			"AvdelingID " &_
			"from Dagsliste_vikar D, KONTAKT, FIRMA, OPPDRAG " &_
			"where D.FakturaStatus < 3 " &_
			"and D.FirmaID = FIRMA.FirmaID " &_
			"and D.BestilltAv = KONTAKT.KontaktID " &_
			"and D.OppdragID = OPPDRAG.OppdragID " &_
			"and D.Dato < " & dbDate(session("limitDato")) &_
			" group by D.FirmaID, Firma, Etternavn, Fornavn, " &_
			"D.BestilltAv, Fakturastatus, AvdelingID " &_
			"order by Firma"


'Response.write strSQL & "<br>"

Set rsVikar = Conn.Execute(strSQL)


ElseIF viskode = 2 Then

strSQL = "select distinct D.FirmaID, Firma, Navn=(Etternavn + ' ' + Fornavn), " &_
			"BestilltAV=D.Kontakt, stat1=Status, FakturaDato, FakturaNr, " &_
			"AvdelingID=Avdeling " &_
			"from FAKTURAGRUNNLAG D, KONTAKT, FIRMA " &_
			"where D.Status = 2 " &_
			"and D.FirmaID = FIRMA.FirmaID " &_
			"and D.Kontakt = KONTAKT.KontaktID " &_
			" order by FakturaNr"



'Response.write strSQL & "<br>"

Set rsVikar = Conn.Execute(strSQL)

ElseIf viskode = 3 Then

strSQL = "select distinct D.OppdragID, D.FirmaID, Firma, Navn=(Etternavn + ' ' + Fornavn), " &_
			"BestilltAV=D.Kontakt, stat1=Status, D.FakturaDato, FakturaNr, " &_
			"AvdelingID=Avdeling " &_
			"from FAKTURAGRUNNLAG D, KONTAKT, FIRMA " &_
			"where D.Status = 3 " &_
			"and D.FirmaID = FIRMA.FirmaID " &_
			"and D.Kontakt = KONTAKT.KontaktID " &_
			"and D.FakturaDato >= "  & dbDate(session("oldStartDate")) &_
			" order by FakturaNr"


'Response.write strSQL & "<br>"

Set rsVikar = Conn.Execute(strSQL)

End If  'viskode

If viskode > 0 Then

If rsVikar.BOF = True And rsVikar.EOF = True Then 

	Response.write "<br>Ingen fakturaer med status " & viskode & "!"
Else


'--------------------------------------------------------------------------------------------------
' Display table headdings
'--------------------------------------------------------------------------------------------------
If viskode = 1 Or viskode = 3 Then %>
<A HREF=#A>A</A>
<A HREF=#B>B</A>
<A HREF=#C>C</A>
<A HREF=#D>D</A>
<A HREF=#E>E</A>
<A HREF=#F>F</A>
<A HREF=#G>G</A>
<A HREF=#H>H</A>
<A HREF=#I>I</A>
<A HREF=#J>J</A>
<A HREF=#K>K</A>
<A HREF=#L>L</A>
<A HREF=#M>M</A>
<A HREF=#N>N</A>
<A HREF=#O>O</A>
<A HREF=#P>P</A>
<A HREF=#R>R</A>
<A HREF=#S>S</A>
<A HREF=#T>T</A>
<A HREF=#U>U</A>
<A HREF=#V>V</A>
<A HREF=#Ø>Ø</A>
<A HREF=#Å>Å</A>
<% End If %>

<table cellpadding='0' cellspacing='0'>
	<tr>
		<!--TH>OppdID</TH-->
		<% If viskode > 1 Then %>
		<th>Fakt.nr</th>
		<% End If %>
		<th>Kontakt</th>
		<th>Kontaktperson</th>
		<th>Avd</th>
		<th>FStat</th>
		<% If viskode = 1 Then %>
		<th>Vikar</th>
		<th>TStat</th>
		<th>LStat</th>
		<% End If %>
		<th>Fakt</th>
		<% If viskode = 2 Then %>
		<th>Nedgrad</th>
	</tr>
	<% 
	End If
'--------------------------------------------------------------------------------------------------
' Show data
'--------------------------------------------------------------------------------------------------
Bok = Left(rsVikar("Firma"),1)

Do Until rsVikar.EOF


%>

	<tr>
		<!--TD> <% '=rsVikar("OppdragID") %></TD-->

<% If viskode > 1 Then %>
		<TD><% =rsVikar("Fakturanr") %></TD>
<% End If %>

<% If viskode = 1 Then %>
	<% If rsVikar("maks") = 5  Then %>
	<TD> <A HREF="Vikar_timeliste_fakt_vis2.asp?viskode=<% =viskode %>&FirmaID=<% =rsVikar("FirmaID") %>&BestilltAv=<% =rsVikar("BestilltAv") %>" ><% =rsVikar("Firma") %></A>
<%   If Bok <> Left(rsVikar("Firma"),1) Then
	Bok = Left(rsVikar("Firma"),1)	
	Response.write "<A NAME=" & Bok & ">"
   End If %>
	<% Else %>
	<TD> <% =rsVikar("Firma") %>
	<% End If %>
<% End If %>

<% If viskode = 2 Then %>
	<TD><A HREF="Faktura_vis.asp?Kontakt=<% =rsVikar("BestilltAv") %>&OppdragID=<% '=rsVikar("OppdragID") %>&FirmaID=<% =rsVikar("FirmaID") %>"   ><% =rsVikar("Firma") %></A></TD>
<% End If %>

<% If Viskode = 3 Then %>
	<TD> <% =rsVikar("Firma") %>
<%   If Bok <> Left(rsVikar("Firma"),1) Then
	Bok = Left(rsVikar("Firma"),1)	
	Response.write "<A NAME=" & Bok & ">"
   End If %>
<% End If %>

<TD><% = rsVikar("BestilltAV") %>-<% =rsVikar("Navn") %>
<TD><% =rsVikar("AvdelingID") %>
<TD>
<% If rsVikar("stat1") = 3 Then %>
<FONT COLOR=BLACK >
<% ElseIf rsVikar("stat1") = 2 Then %>
<FONT COLOR=GREEN >
<% Else %>
<FONT COLOR=#CB1700>
<% End If %>
(<% =rsVikar("stat1") %>)

<% If viskode = 1 Then %>
<TD><TD><TD>
<% End If %>

<TD>
<% If viskode = 1 Then %>
<% If rsVikar("maks") = 5 Then %>
<A HREF="Faktura_vis.asp?Kontakt=<% =rsVikar("BestilltAv") %>&FirmaID=<% =rsVikar("FirmaID") %>&Avdeling=<% =rsVikar("AvdelingID") %>" TARGET=_new  >Fakt</A></TD>
<% End If %><% End If %>
<% If viskode = 2 Then %>
<TD><A HREF="Faktura_lagre.asp?graderingskode=nedgrad&Kontakt=<% =rsVikar("BestilltAv") %>&OppdragID=<% '=rsVikar("OppdragID") %>&FirmaID=<% =rsVikar("FirmaID") %>&Fakturadato=<% =rsVikar("Fakturadato") %>"   >Nedgrad</A></TD>
<% end If %>
<% If viskode = 3 Then %>
<TD><A HREF="Faktura_lagre.asp?graderingskode=nedgrad&Kontakt=<% =rsVikar("BestilltAv") %>&OppdragID=<% =rsVikar("OppdragID") %>&FirmaID=<% =rsVikar("FirmaID") %>&Fakturadato=<% =rsVikar("Fakturadato") %>&frakode=3&Fakturanr=<% =rsVikar("Fakturanr") %>"   >Nedgrad</A></TD>
<% end If %>
<%
If viskode = 1 Then

strSQL = "select distinct D.VikarID, D.OppdragId, Navn=(Etternavn + ' ' + Fornavn), " &_
			"stat1=Fakturastatus, stat2=TimelisteVikarStatus, " &_
			"D.BestilltAv, Loennstatus, AvdelingID " &_
			"from Dagsliste_vikar D, VIKAR, OPPDRAG " &_
			"where D.FakturaStatus < 3 " &_
			"and D.VikarID = VIKAR.VikarID " &_
			"and D.OppdragID = OPPDRAG.OppdragID " &_
			"and D.BestilltAv = " & rsVikar("BestilltAv") &_
			"and OPPDRAG.AvdelingID = " & rsVikar("AvdelingID") &_
			" and D.Dato < " & dbDate(session("limitDato")) &_
			" order by D.BestilltAv"



'Response.write strSQL & "<br>"


Set rsVikar2 = Conn.Execute(strSQL)

Do Until rsVikar2.EOF


%>
<tr><TD><TD><TD><% =rsVikar2("AvdelingID") %>
<TD><TD><% =rsVikar2("VikarID") %>
<A HREF="Vikar_timeliste_vis3.asp?VikarID=<% =rsVikar2("VikarID") %>&OppdragID=<% =rsVikar2("OppdragID") %>&frakode=3" TARGET=_new ><% =rsVikar2("Navn") %></A>

<TD>
<% If rsVikar2("stat2") = 5 Then %>
<FONT COLOR=GREEN >
<% ElseIf rsVikar2("stat2") = 4 Then %>
<FONT COLOR=YELLOW >
<% Else %>
<FONT COLOR=#CB1700>
<% End If %>
(<% =rsVikar2("stat2") %>)
</TD>
<TD>
<% If rsVikar2("Loennstatus") = 2 Then %>
<FONT COLOR=GREEN >
<% ElseIf rsVikar2("Loennstatus") = 3 Then %>
<FONT COLOR=BLACK >
<% Else %>
<FONT COLOR=#CB1700>
<% End If %>
(<% =rsVikar2("Loennstatus") %>)
</TD>
<%
rsVikar2.MoveNext


Loop
rsVikar2.Close
End If  'kode = 1

rsVikar.MoveNext
Loop

rsVikar.Close
%>
</table>
<%

 If viskode = 2 Then %>
<br>
<table cellpadding='0' cellspacing='0'>
<tr>
<FORM ACTION="Faktura_overf_ordrehode.asp" METHOD=POST >
<th><input name=btnOverfoer TYPE=SUBMIT  VALUE="           Lag fil til Rubicon              "></th>
<th>Siste ordrenr:
<th><input name=ONR TYPE=TEXT SIZE=5 ></th>
</form></table>

<% End If %>
<% If viskode = 1 Or viskode = 3 Then %>
<A HREF=#A>A</A>
<A HREF=#B>B</A>
<A HREF=#C>C</A>
<A HREF=#D>D</A>
<A HREF=#E>E</A>
<A HREF=#F>F</A>
<A HREF=#G>G</A>
<A HREF=#H>H</A>
<A HREF=#I>I</A>
<A HREF=#J>J</A>
<A HREF=#K>K</A>
<A HREF=#L>L</A>
<A HREF=#M>M</A>
<A HREF=#N>N</A>
<A HREF=#O>O</A>
<A HREF=#P>P</A>
<A HREF=#R>R</A>
<A HREF=#S>S</A>
<A HREF=#T>T</A>
<A HREF=#U>U</A>
<A HREF=#V>V</A>
<A HREF=#Ø>Ø</A>
<A HREF=#Å>Å</A>
<% End If %>

<% End If 'ingen treff %>

<% 'Response.Write "<br>" & Session("tilgang") %>
<% End If 'viskode > 0 %>
</body>
</html>
