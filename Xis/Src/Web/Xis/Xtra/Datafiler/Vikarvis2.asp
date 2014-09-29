<html>
<head>
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
' Connect to database
'--------------------------------------------------------------------------------------------------

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

If Request.QueryString("VikarID") = "" Then
   strVikarID = Request.Form("VikarID")
   strOppdragID = Request.Form("OppdragID")
   strFirmaID = Request.Form("FirmaID")
   kode = Request.Form("kode")
   tilgang = Request.Form("tilgang")
Else
   strVikarID = Request.QueryString("VikarID")
   strOppdragID = Request.QueryString("OppdragID")
   strFirmaID = Request.QueryString("FirmaID")
   kode = Request.QueryString("kode")
   tilgang = Request.QueryString("tilgang")
End If
If tilgang = "" Then
	tilgang = 1
End If

'Response.write strVikarID & "<br>"
'Response.write strOppdragId & "<br>"
'Response.write strFirmaID & "<br>"
'Response.write kode & "<br>"
'Response.write tilgang & "<br>"



'--------------------------------------------------------------------------------------------------
' SQL to get data
'--------------------------------------------------------------------------------------------------




'	Set rsVikar = Conn.Execute("select VikarId, Navn=(Fornavn + ' ' + Etternavn), Adresse, " &_ 
'			"ASted=(ADRESSE.Postnr + ' ' + ADRESSE.PostSted) , Ansattdato, " &_
'			"Foedselsdato, Personnummer , Bankkontonr, " &_
'			"Kommunenr, TypeID, Skattetabellnr, Skatteprosent, " &_
'			"Loenn1 " &_
'			"from VIKAR, ADRESSE " &_
'			"where VIKAR.VikarID = ADRESSE.AdresseRelID " &_
'			"and ADRESSE.AdresseType = 1 " &_
'			"and ADRESSE.Adresserelasjon = 2 " &_
'			"and VikarID = " & strVikarID )

	strVikarSQL = "SELECT " & _
					"VIKAR.VikarId, " & _
					"VIKAR.Foedselsdato, " & _
					"VIKAR.Personnummer, " & _
					"VIKAR.Ansattdato, " & _
					"VIKAR.Bankkontonr, " & _
					"VIKAR.Kommunenr, " & _
					"VIKAR.TypeID, " & _
					"VIKAR.Skattetabellnr, " & _
					"VIKAR.Skatteprosent, " & _
					"VIKAR.Loenn1, " & _
					"Navn = (VIKAR.Fornavn + ' ' + VIKAR.Etternavn), " & _
					"ADRESSE.Adresse, " & _
					"ASted = (ADRESSE.Postnr + ' ' + ADRESSE.PostSted), " & _
					"VIKAR_ANSATTNUMMER.ansattnummer " & _
					"FROM VIKAR " & _
					"LEFT OUTER JOIN ADRESSE ON VIKAR.Vikarid = ADRESSE.Adresserelid " & _
					"LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " & _
					"WHERE ADRESSE.Adressetype = '1' " & _
					"AND ADRESSE.Adresserelasjon = '2' " & _
					"AND VIKAR.Vikarid = '" & strVikarID & "' "

	Set rsVikar = Conn.Execute(strVikarSQL)

   fNr =  rsVikar("Foedselsdato") & " " & rsVikar("Personnummer")

'--------------------------------------------------------------------------------------------------
' Display data
'--------------------------------------------------------------------------------------------------
%>
<H3><% =rsVikar("Navn") %></H3>


<table cellpadding='0' cellspacing='0'>
<tr><th class="right">Ansattnummer:<th><th><%=rsVikar("ansattnummer").Value%>
<tr><th class="right">Navn		:<th><th><%=rsVikar("Navn").Value%>
<tr><th class="right">Adresse		:<th><th><%=rsVikar("Adresse").Value%>
<tr><th class="right">Sted		:<th><th><%=rsVikar("ASted").Value%>
<tr><th class="right">Fødselsnr		:<th><th><%=fNr%>
<tr><th class="right">Ansattdato		:<th><th><%=rsVikar("Ansattdato").Value%>
<tr><th class="right">Bankkontonr		:<th><th><%=rsVikar("Bankkontonr").Value%>
<tr><th class="right">Kommunenr		:<th><th><%=rsVikar("Kommunenr").Value%>
<tr><th class="right">Timelønn		:<th><th><%=rsVikar("Loenn1").Value%>
<tr><th class="right">Skattetabellnr	:<th><th><%=rsVikar("Skattetabellnr").Value%>
<tr><th class="right">Skatteprosent	:<th><th><%=rsVikar("Skatteprosent").Value%>
<tr><th class="right">A/S			:<th><th><% If rsVikar("TypeID")= 3 Then Response.write "Ja" %>
</table>

<table cellpadding='0' cellspacing='0'>

<% If tilgang = 1 Then  'synlig for vikar %>


<FORM ACTION="Vikarny2.asp?VikarID=<% =strVikarID %>&kode=<% =kode %>&tilgang=<% =tilgang %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>" TARGET=RIGHT_WINDOW METHOD=POST>
<tr><th>
<tr><th>
<input name=btnOverfoer TYPE=SUBMIT VALUE=" Endre personopplysninger ">
</th><tr>
</form>

<FORM ACTION="Vikar_timeliste_vis.asp?VikarID=<% =strVikarID %>&kode=<% =kode %>&tilgang=<% =tilgang %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&Navn=<% =rsVikar("Navn") %>" METHOD=POST TARGET=RIGHT_WINDOW >
<tr><th>
<input name=btnOverfoer TYPE=SUBMIT  VALUE="                 Timelister                  ">
</th><tr>
<input name="VikarID" TYPE=HIDDEN VALUE=<% =rsVikar("VikarID") %> >
<input name="Navn" TYPE=HIDDEN VALUE="<% =rsVikar("Navn") %>" >
<input name="OppdragID" TYPE=HIDDEN VALUE="<% =Request.QueryString("OppdragID") %>" >
<input name="FirmaID" TYPE=HIDDEN VALUE="<% =Request.QueryString("FirmaID") %>" >
</form>

<FORM ACTION="Hult_Lill_05.asp?kode=<% =kode %>&tilgang=<% =tilgang %>&VikarID=<% =strVikarID %>" TARGET=_parent METHOD=POST>
<tr><th>
<input name=btnOverfoer TYPE=SUBMIT  VALUE="                Velg ny                        ">
</th>
</form> 
<tr></table>

<% End If 'synlig for vikar %>


<% If tilgang = 2 Then 'synlig for personalkonsulent %>

<FORM ACTION="Vikarny2.asp?VikarID=<% =strVikarID %>&kode=<% =kode %>&tilgang=<% =tilgang %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>" TARGET=RIGHT_WINDOW METHOD=POST>
<tr><th>
<tr><th>
<input name=btnOverfoer TYPE=SUBMIT  VALUE=" Endre personopplysninger ">
</th><tr>
</form>

<FORM ACTION="Vikar_timeliste_vis.asp?VikarID=<% =strVikarID %>&kode=<% =kode %>&tilgang=<% =tilgang %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&Navn=<% =rsVikar("Navn") %>" METHOD=POST TARGET=RIGHT_WINDOW >
<tr><th>
<input name=btnOverfoer TYPE=SUBMIT  VALUE="                 Timelister                  ">
</th><tr>
<!--INPUT NAME="VikarID" TYPE=HIDDEN VALUE=<% =rsVikar("VikarID") %> >
<input name="Navn" TYPE=HIDDEN VALUE="<% =rsVikar("Navn") %>" >
<input name="OppdragID" TYPE=HIDDEN VALUE="<% =Request.QueryString("OppdragID") %>" >
<input name="FirmaID" TYPE=HIDDEN VALUE="<% =Request.QueryString("FirmaID") %>" -->
</form>

<FORM ACTION="Hult_Lill_02.asp?kode=<% =kode %>&tilgang=<% =tilgang %>" TARGET=_parent METHOD=POST>
<tr><th>
<input name=btnOverfoer TYPE=SUBMIT  VALUE="                Velg ny                        ">
</th>
</form> 
<tr>

<% End If 'synlig for personalkonsulent %>


<% If tilgang = 3 Then 'synlig for lønnskonsulent %>

<FORM ACTION="Vikarny2.asp?VikarID=<% =strVikarID %>&kode=<% =kode %>&tilgang=<% =tilgang %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>" TARGET=RIGHT_WINDOW METHOD=POST>
<tr><th>
<tr><th>
<input name=btnOverfoer TYPE=SUBMIT  VALUE=" Endre personopplysninger ">
</th><tr>
</form>

<FORM ACTION="Vikar_fastl_vis.asp?kode=<% =kode %>&tilgang=<% =tilgang %>" METHOD=POST TARGET=RIGHT_WINDOW >
<tr><th>
<input name=btnOverfoer TYPE=SUBMIT  VALUE="                  Fast lønn                   ">
</th><tr>
<input name="VikarID" TYPE=HIDDEN VALUE=<% =rsVikar("VikarID") %> >
<input name="Navn" TYPE=HIDDEN VALUE="<% =rsVikar("Navn") %>" >
</form>




<% If kode = 0 Then %>

<FORM ACTION="Hult_Lill_03.asp?kode=<% =kode %>&tilgang=<% =tilgang %>" TARGET=_parent METHOD=POST>
<tr><th>
<input name=btnOverfoer TYPE=SUBMIT  VALUE="                Velg ny                        ">
</th>
</form> 
<tr></table>

<% Else %>

<% If kode = 3 then %>

<FORM ACTION="../vikarvis.asp?vikarID=<% =rsVikar("VikarID") %>&kode=<% =kode %>&tilgang=<% =tilgang %>" TARGET=_parent METHOD=POST>
<tr><th>
<input name=btnOverfoer TYPE=SUBMIT  VALUE="                Tilbake                        ">
</th>
</form> 
<tr></table>

<% ElseIf kode = 5 Then %>

<FORM ACTION="Hult_Lill_06.asp?kode=<% =kode %>&tilgang=<% =tilgang %>" TARGET=_parent METHOD=POST>
<tr><th>
<input name=btnOverfoer TYPE=SUBMIT  VALUE="                Velg ny                        ">
</th>
</form> 
<tr></table>

<% Else 'kode ikke er 3 %> 

<FORM ACTION="Hult_Lill_04.asp?kode=<% =kode %>&tilgang=<% =tilgang %>" TARGET=_parent METHOD=POST>
<tr><th>
<input name=btnOverfoer TYPE=SUBMIT  VALUE="                Velg ny                        ">
</th>
</form> 
<tr></table>


<% End If %>
<% End If %>

<% End If  'synlig for lønnskonsulent %>

<%
   rsVikar.Close
%>

</body>
</html>
