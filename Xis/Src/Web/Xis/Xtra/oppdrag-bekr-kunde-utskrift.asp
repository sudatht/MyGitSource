<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"--> 
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes\Xis.Settings.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="includes\CRM.Integration.Contact.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim lOppdragID
	dim lOppdragVikarID
	dim RSOppdrag
	dim RSAvdKontor
	dim strAvdKontorFax
	dim strTitle
	dim Conn
	dim strSQL
	dim strAdresse : strAdresse = ""
	dim strPostadresse : strPostadresse =""
	dim strKontaktperson
	dim vikar	
	dim aXmlHTTP

	Response.Clear
	Response.ContentType = "application/html"
	vikar = Request("vikarname")
	Response.AddHeader "Content-Disposition", "attachment;filename=Oppdragsbekreftelse-"& vikar &".htm"
	
	' Move parameter OppdragsID to variable
	If Request("OppdragID") <> "" Then
	   lOppdragID = CLng(Request("OppdragID"))
	Else
		AddErrorMessage("Feil:Parameter for oppdragid mangler!")
	End If

	' Move parameter VikarID to variable
	If Request("OppdragVikarID") <> "" Then
	   lOppdragVikarID = CLng( Request("OppdragVikarID") )
	Else
	   AddErrorMessage("Feil:Parameter for OppdragVikarID mangler!")
	End If

	if(HasError() = true) then
		call RenderErrorMessage()
	end if

	' Get a database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	' Get information from database
	' Get oppdrags data
	Set rsOppdrag  = GetFirehoseRS("SELECT F.Firma, F.SOCuID, F.Fax, O.FirmaID, O.SOPeID, Ansvarlig = M.Fornavn + ' ' + M.Etternavn, KontaktPerson = K.Fornavn + ' ' + K.Etternavn, M.dirtlf, K.Fax, " &_
	" ISNULL(O.AnsMedID, 0) AS AnsMedID, " & _	
	"O.Beskrivelse, O.ArbAdresse, " &_	
	"O.Notatkunde " &_
	"FROM Oppdrag O, Firma F, Medarbeider M, KONTAKT K, ADRESSE A " &_
	" WHERE O.OppdragID = " & lOppdragID & _
	" and O.AnsMedID *= M.MedID " &_
	" and O.BestilltAv *= K.KontaktID " &_
	" and O.FirmaID = F.FirmaID ", Conn)

	'Dim cts 
	'Set cts = server.CreateObject("Integration.SuperOffice")

	strKontaktperson = rsOppdrag("Kontaktperson")
	
	if(isnull(strKontaktperson)) then
		'dim personRs 		
		'set personRs = cts.GetPersonSnapshotById(clng(rsOppdrag("SOPeID")))
		'if (not personRs.EOF) then
			'if (isnull(personRs("middlename"))) then
				'strKontaktperson = personRs("firstname") & " " & personRs("lastname")
			'else
				strKontaktperson = GetCRMContactPersonName(rsOppdrag("SOPeID"))
			'end if
		'end if
		'set personRs = nothing
	end if
	if(not isnull(rsOppdrag("SOCuID"))) then
		'dim rsAddress
		'set rsAddress = cts.GetAddressByContactId(clng(rsOppdrag("SOCuID")), 1)
		'if (not rsAddress.EOF) then
		Set aXmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		aXmlHTTP.Open "POST", Application("XISCRMLink")+"?PageMode=GetAccountAddress&Socuid=" + Cstr(rsOppdrag("SOCuID")) , false, Application("Xtra_Domain") +"\" + Application("Xtra_User"), Application("Xtra_Password")
		aXmlHTTP.send ""
		strAdresse = aXmlHTTP.responseText

		'	strPostadresse = rsAddress("zipcode") & " " & rsAddress("city")
		'end if		
		'set rsAddress = nothing
	end if
	'set cts = nothing

	strSQL = "SELECT " & _
	"OPPDRAG_VIKAR.VikarID, " & _
	"OPPDRAG_VIKAR.Fradato, " & _
	"OPPDRAG_VIKAR.Tildato, " & _
	"OPPDRAG_VIKAR.Frakl, " & _
	"OPPDRAG_VIKAR.Tilkl, " & _
	"OPPDRAG_VIKAR.Utvid, " & _
	"OPPDRAG_VIKAR.Timepris, " & _
	"OPPDRAG_VIKAR_CONFIRMATION_TEXT.CustomerText AS BekreftelseKundeTekst, " & _
	"Vikarnavn = VIKAR.Fornavn + ' ' + VIKAR.Etternavn, " & _
	"PostAdr = ADRESSE.Postnr + ' ' + ADRESSE.Poststed, " & _
	"VIKAR_ANSATTNUMMER.ansattnummer " & _
	"FROM OPPDRAG_VIKAR_CONFIRMATION_TEXT,OPPDRAG_VIKAR " & _
	"LEFT OUTER JOIN VIKAR ON OPPDRAG_VIKAR.Vikarid = VIKAR.Vikarid " & _
	"LEFT OUTER JOIN ADRESSE ON OPPDRAG_VIKAR.Vikarid = ADRESSE.Adresserelid " & _
	"LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON OPPDRAG_VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " & _
	"WHERE OPPDRAG_VIKAR_CONFIRMATION_TEXT.OppdragVikarID = OPPDRAG_VIKAR.OppdragVikarID and OPPDRAG_VIKAR.Oppdragvikarid = '" & lOppdragVikarID & "' " & _
	"AND OPPDRAG_VIKAR.Statusid = '4' " & _
	"AND ADRESSE.Adresserelasjon = '2' " & _
	"AND ADRESSE.AdresseType = '1' "

	'Response.write strSQL
'ss
	Set rsOppdragVikar = GetFirehoseRS( strSQL, Conn )
	
	if (NOT rsOppdrag.EOF) then
		if (len(rsOppdrag("AnsMedID")) > 0 ) then

			'Hent stedsnavn fra pålogget brukers avdelingskontor
			strSQL = "SELECT [Lokasjon].[Navn] " & _
			"FROM [Lokasjon] " & _
			"INNER JOIN [Avdelingskontor] ON [Avdelingskontor].[LokasjonID] = [Lokasjon].[LokasjonID] " & _
			"INNER JOIN [medarbeider] ON [medarbeider].[AvdelingskontorID] = [Avdelingskontor].[ID] " & _
			"WHERE [Medarbeider].[MedID] =" & rsOppdrag("AnsMedID")
			
			Set rsAvdKontor = GetFirehoseRS(strSQL, Conn)

			' Error from database ?
			If Conn.Errors.Count > 0 then
				Call SqlError()
			End if

			If not rsAvdKontor.EOF then
				Sted = rsAvdKontor("navn") & ",&nbsp;"
			Else
				Sted = ""
			End If
			rsAvdKontor.close
			set rsAvdKontor = Nothing

		end if
	end if
	
	' Utvidelse
	If rsOppdragVikar( "Utvid" ) = 1 Then
	   Heading = "Forlengelse av oppdrag"
	   strTitle =  "Oppdragsforlengelse " & rsOppdragVikar("Vikarnavn")
	Else
	   strTitle =  "Oppdragsbekreftelse " & rsOppdragVikar("Vikarnavn")
	   Heading =   "Bekreftelse av oppdrag"
	End If

	'faxnummer til pålogget konsulentleders avdelingskontor
	strSQL = "SELECT [A].[Fax] " & _
			"FROM [avdelingskontor] AS [A] " & _
			"INNER JOIN [VIKAR_ARBEIDSSTED] AS [VA] ON [VA].[AvdelingsKontorID] = [a].[ID] " & _
			"INNER JOIN [Vikar] AS [V] ON [V].[VikarID] = [Va].[VikarID] " & _
			"INNER JOIN [Medarbeider] AS [M] ON [m].[VikarID] = [V].[VikarID] " & _
			"WHERE [M].[MedID] = " & session("medarbID")
			
	set rsAvdKontor = GetFirehoseRS(strSql, Conn)
	strAvdKontorFax = rsAvdKontor("Fax")
	RSAvdKontor.close
	set RSAvdKontor = nothing
	
' Create Oppdragsbeksreftelse
%>
<html>
<head>
	<title><%=strTitle%></title>
		<style type="text/css">
			body {font:normal 11px/normal tahoma, sans-serif; margin:0; padding:0 2em; color:#000; background:transparent;} 
			h1, h2, h3, h4 {margin:.2em 0 .3em 0;}
			h1 {font-size:2.4em; font-weight:normal;}
			h2 {font-size:1.5em; font-weight:normal;} 
			h3 {font-size:1.1em;}
			h4 {font-size:1em;}
			p {font-size:1em; margin:0 0 1em 0;}
			
			table {font-size:1em; empty-cells:show;}
			tfoot {font-weight:bold;}
			caption {text-align:left; font-weight:bold; font-style:normal; font-size:1.2em;}
			th, td {text-align:left;}
			
			ul {list-style-type:square; margin-top:0; margin-bottom:0; margin-left:3em; padding:0 0 0 0; line-height:normal;}
			li {list-style-position:outside; margin:0 0 0 0; padding:0 0 0 0;}
			
			blockquote {margin: 1em 2em 1em 2em; }
			
			hr {border:inset;}
			img {border:none;}
			
			/*newWindow*/
			.newWindow {color:#000;}
			.showCV {margin:1.5em 0 0 0;}
			.newWindow .logo {position:absolute; top:0; left:84%; width:63px;}
			.showCV h1 {border-bottom:double #336666;}
			.showCV h2 {font-size:1.6em; margin-top:.5em; /*border-bottom:1px solid #336666;*/} 
			.showCV td, .showCV th {border-bottom:1px solid #cccccc;}
			.showCV table {width:100%;}
			.showCV tr {vertical-align:top;}
			.addressField {margin-left:4em; margin-bottom:1em;}

			.showAgreement h2 {font-size:1.5em; border-bottom:double #336666;}
			.showAgreement h3, .showAgreement h4 {margin-top:.5em;}
			
			/*minor details... major impacts*/ 
			.left {text-align:left;}
			.center {text-align:center;}
			.right {text-align:right;}
			.warning {color:#ff0000; background:transparent;}
			.top {vertical-align:top;}
			
			@media print {
				body {padding:0 0 0 0;}
				.showAgreement {page-break-before:always; font:normal 9px/normal arial, sans-serif;}
				.showAgreement h2, .showAgreement h3 {margin-top:0;}
				
				.showAgreement h2, .showAgreement h3, .showAgreement h4 {margin:0 0 0 0; font-weight:bolder;}
				.showAgreement h2 {font-size:1.4em; margin-bottom:.2em; border-bottom:none;}
				.showAgreement h3 {font-size:1.2em; margin-top:.5em;}
				.showAgreement h4 {font-size:1em;}
				.showAgreement p {margin:0 0 .4em 0;}
				ul {list-style-type:square; margin-top:0; margin-bottom:0; margin-left:3em; padding:0 0 0 0; line-height:normal;}
				li {list-style-position:outside; margin:0 0 0 0; padding:0 0 0 0;}
			}
		</style>	
</head>
<body class="newWindow">
	<div class="showCV">
		<div><p style="text-align:right;"><img src="http://intern.xtra.no/portals/0/site_images/xtra_logo_cv.png" width="150px" height="44px" alt="Centric logo"></p></div>
		
		<div class="addressField">
			<p>
				<%=rsOppdrag("Firma")%><br>
				v/&nbsp;<%=strKontaktperson%><br>
				<%=strAdresse%><br>
				<%=strPostadresse%>
			</p>
		</div>
		<p class="right"><%= Sted & date()%></p>
		
		<h1><%=Heading%></h1>
		<table>
			<col width="30%">
			<col width="70%">
			<tr>
				<th>Oppdragsnr:</th>
				<td><%=lOppdragID%></td>
			</tr>
			<tr>
				<th>Beskrivelse av oppdrag:</th>
				<td><%=rsOppdrag("Beskrivelse")%></td>
			</tr>
			<tr><td colspan="2">&nbsp;</td></tr>
			<tr>
				<th>Utleid medarbeider:</th>
				<td><%=rsOppdragVikar("Vikarnavn")%></td>
			</tr>
			<tr>
				<th>Ansattnummer:</th>
				<td><%=rsOppdragVikar("ansattnummer")%></td>
			</tr>
			<tr><td colspan="2">&nbsp;</td></tr>
			<tr>
				<th>Kontaktperson hos oppdragsgiver:</th>
				<td><%=strKontaktperson%></td>
			</tr>
			<tr>
				<th>Kontaktperson hos Centric:</th>
				<td><%=rsOppdrag("Ansvarlig")%>, telefon <%=rsOppdrag("dirtlf") %></td>
			</tr>
			<tr><td colspan="2">&nbsp;</td></tr>
			<tr>
				<th>Startdato:</th>
				<td><%=rsOppdragVikar("Fradato")%></td>
			</tr>
			<tr>
				<th>Sluttdato:</th>
				<td><%=rsOppdragVikar("Tildato")%></td>
			</tr>
			<tr>
				<th>Arbeidstid:</th>
				<td><%=FormatDateTime( rsOppdragVikar("FraKl"), 4)%> - <%=FormatDateTime( rsOppdragVikar("TilKl"), 4)%></td>
			</tr>
			<tr>
				<th>Timepris:</th>
				<td><%=rsOppdragVikar("Timepris")%>,-&nbsp;(eks.mva)</td>
			</tr>
			
	</table>
	<table>
		<tr>
				<td>Eventuelt overtidsarbeid skal være pålagt av oppdragsgiver og avtales direkte med vikaren, samt alltid være i henhold til bestemmelsene i Arbeidsmiljølovens § 10-6.
				</td>
			</tr>
	</table>
	<h2>Ytterligere oppdragsopplysninger:</h2>
	<p><%= replace(rsOppdragVikar.fields("BekreftelseKundeTekst").value, vbcrlf, "<br>")%>
	<%=GetSetting("OppdragBekreftelseNotat")%><br/>
	<%=rsOppdrag("Ansvarlig")%>
	</p>		
	
	</div>
</div>
</body>
</html>
<%
RSOppdrag.close
set RSOppdrag = nothing
CloseConnection(Conn)
set Conn = nothing
%>