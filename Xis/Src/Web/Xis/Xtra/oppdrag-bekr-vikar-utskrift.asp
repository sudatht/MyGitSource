<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes/xis.rights.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="includes\CRM.Integration.Contact.inc"--> 

<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim lOppdragID
	dim lOppdragVikarID
	dim Conn
	dim strSQL
	dim strTitle
	dim Sted
	dim rsAvdKontor
	dim rsOppdrag
	dim rsOppdragVikar
	dim rsAdresse
	dim kontaktperson
	dim cts
	dim adresse : adresse = vbNullString
	dim postadresse : postadresse = vbNullString

	Response.Clear
	Response.ContentType ="application/html"
	Response.AddHeader "Content-Disposition", "attachment;filename=oppdragbekreftelsekunde.htm"

	' Move parameter OppdragsID to variable
	If Request("OppdragID") <> "" Then
	   lOppdragID = CLng(Request("OppdragID"))
	Else
		AddErrorMessage("Feil: OppdragID mangler!")
		call RenderErrorMessage()	
	End If

	' Move parameter OppdragVikarID to variable
	If Request("OppdragVikarID") <> "" Then
		lOppdragVikarID = CLng( Request("OppdragVikarID") )
	Else
		AddErrorMessage("Feil: OppdragVikarID mangler!")
		call RenderErrorMessage()	
	End If

	' Open database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	' Get relevant information FROM database

	' Get oppdrags data
	strSQL = "SELECT F.Firma, F.FirmaID, F.SOCuID, Ansvarlig = M.Fornavn + ' ' + M.Etternavn, M.dirtlf, " &_ 
	"KontaktPerson = K.Fornavn + ' ' + K.Etternavn, O.SOPeID, " &_
	"O.Beskrivelse, O.ArbAdresse, O.AnsMedID " &_
	"FROM Oppdrag O, Firma F, Medarbeider M, KONTAKT K  " &_
	" WHERE O.OppdragID = " & lOppdragID & _
	" AND O.FirmaID = F.FirmaID " &_
	" AND O.AnsMedID *= M.MedID " &_
	" AND O.BestilltAv *= K.KontaktID "

	Set rsOppdrag = GetFirehoseRS(strSQL, Conn)

	strVikar = "SELECT " & _
	"OPPDRAG_VIKAR.Vikarid, " & _
	"Vikarnavn = VIKAR.Fornavn + ' ' + VIKAR.Etternavn, " & _
	"ADRESSE.Adresse, " & _
	"PostAdr = ADRESSE.Postnr + ' ' + ADRESSE.Poststed, " & _
	"OPPDRAG_VIKAR.Fradato, " & _
	"OPPDRAG_VIKAR.Tildato, " & _
	"OPPDRAG_VIKAR.Frakl, " & _
	"OPPDRAG_VIKAR.Tilkl, " & _
	"OPPDRAG_VIKAR.Utvid, " & _
	"OPPDRAG_VIKAR.Timeloenn, " & _
	"OPPDRAG_VIKAR_CONFIRMATION_TEXT.ConsultantText AS BekreftelseKonsulentTekst, " & _
	"VIKAR_ANSATTNUMMER.ansattnummer " & _
	"FROM OPPDRAG_VIKAR_CONFIRMATION_TEXT, OPPDRAG_VIKAR " & _
	"LEFT OUTER JOIN VIKAR ON OPPDRAG_VIKAR.Vikarid = VIKAR.Vikarid " & _
	"LEFT OUTER JOIN ADRESSE ON OPPDRAG_VIKAR.Vikarid = ADRESSE.Adresserelid " & _
	"LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON OPPDRAG_VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " & _
	"WHERE OPPDRAG_VIKAR_CONFIRMATION_TEXT.OppdragVikarID = OPPDRAG_VIKAR.OppdragVikarID and OPPDRAG_VIKAR.Oppdragvikarid = '" & lOppdragVikarID & "' " &  _
	"AND OPPDRAG_VIKAR.Statusid = '4' " & _
	"AND ADRESSE.Adresserelasjon = '2' " & _
	"AND ADRESSE.AdresseType = '1' "

	Set rsOppdragVikar = GetFirehoseRS(strVikar, Conn)
	
	if (HasRows(rsOppdrag) = true) then
		if (len(rsOppdrag("AnsMedID")) > 0 ) then

			'Hent stedsnavn fra pålogget brukers avdelingskontor
			strSQL = "SELECT [Lokasjon].[Navn] " & _
			"FROM [Lokasjon] " & _
			"INNER JOIN [Avdelingskontor] ON [Avdelingskontor].[LokasjonID] = [Lokasjon].[LokasjonID] " & _
			"INNER JOIN [medarbeider] ON [medarbeider].[AvdelingskontorID] = [Avdelingskontor].[ID] " & _
			"WHERE [Medarbeider].[MedID] =" & rsOppdrag("AnsMedID")
			
			Set rsAvdKontor = GetFirehoseRS(strSQL, Conn)

			If (HasRows(rsAvdKontor) = true) then
				Sted = rsAvdKontor("navn") & ",&nbsp;"
				rsAvdKontor.close
			Else
				Sted = ""
			End If
			set rsAvdKontor = Nothing
		end if
	end if
	
	If rsOppdragVikar( "Utvid" ) = 1 Then
	   Heading = "Forlengelse av oppdrag"
	   strTitle =  "Oppdragsforlengelse " & rsOppdragVikar("Vikarnavn")
	Else
	   strTitle =  "Oppdragsbekreftelse " & rsOppdragVikar("Vikarnavn")
	   Heading =   "Bekreftelse av oppdrag"
	End If

	'Opplysninger om avdelingskontor
	strSQL = "SELECT ak.telefon FROM avdelingskontor ak, oppdrag op WHERE ak.id = op.avdelingskontorID AND op.oppdragid = " & lOppdragID
	set rsAvdKontor = GetFirehoseRS(strSQL, Conn)

' Create Oppdragsbeksreftelse
%>
<html>
	<head>
		<title><%=strTitle%></title>
		<style type="text/css">
			BODY		{font-size: .7em; margin: 0; padding: 0 10px 10px 10px;}
			BODY, TABLE	{font-family: tahoma, sans-serif;}
			H1, H2, H3{margin:0 0 0 0; color:#666666; background:transparent;}
			H1			{font-size:2.4em; font-weight:normal;}
			H2			{font-size:1.5em; font-weight:normal;}
			H3          {font-size:1.1em;}
			P           {font-size:1em; margin:0 0 1em 0;}		
			TABLE		{font-size:1em; empty-cells:show;}
			TFOOT		{font-weight:bold;}
			CAPTION		{text-align:left; font-weight:bold; font-style:normal; font-size:1.2em;}
			TH, TD		{text-align:left;}	
			UL, OL		{margin-top:0; margin-bottom:0;margin-left:2em;}
			UL			{list-style-type:square;}
			OL			{list-style-type:decimal;}
			LI			{list-style-position:outside;}
			LI LI		{list-style-position:outside;}
			BLOCKQUOTE	{ margin: 1em 2em 1em 2em; }
			HR			{border:inset;}
			IMG			{border:none;}
			/*newWindow*/
			.newWindow	{color:#000000; background:#ffffff; padding:0 2em 2em 2em;}
			.showCV		{margin:1.5em 0 0 0;}
			.newWindow .logo	{position:absolute; top:0; left:84%; width:63;}
			.showCV H1	{border-bottom:double #336666;}
			.showCV H2	{font-size: 1.6em; margin-top:.5em; /*border-bottom:1px solid #336666;*/}
			.showCV TD, .showCV TH	{border-bottom:1px solid #cccccc;}
			.showCV TABLE	{width:100%;}
			.showCV TR	{vertical-align:top;}
			.addressField {margin-left: 1cm; margin-bottom:2cm;}
			/*minor details... major impacts*/
			.left		{text-align:left;}
			.center		{text-align:center;}
			.right		{text-align:right;}
			.warning	{color:#ff0000; background:transparent;}
			.top		{vertical-align:top;}
		</style>
	</head>
	<body class="newWindow">
	<table width="100%">
		<tr>
			<td align="right"><img align="right" src="http://intern.xtra.no/portals/0/site_images/xtra_logo_cv.png" width="150px" height="44px" alt="Centric logo"></td>
		</tr>
	</table>
	<div class="showCV">
		<div class="addressField">
			<p>
				<%=rsOppdragVikar("Vikarnavn")%><br>
				<%=rsOppdragVikar("Adresse")%><br>
				<%=rsOppdragVikar("PostAdr")%>
			</p>
		</div>
		<p class="right"><%= Sted & date()%></p>
		<h1><%=Heading%></h1>
		<table>
			<col width="30%">
			<col width="70%">
			<tr>
				<th>Ansattnummer:</th>
				<td><%=rsOppdragVikar("ansattnummer")%></td>
			</tr>
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
				<th>Oppdragsgiver:</th>
				<td>
					<%=rsOppdrag("Firma")%><br>
					<%											
					if ( not IsNull(rsOppdrag("SOCuid").value)) Then
						Set aXmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
						aXmlHTTP.Open "POST", Application("XISCRMLink")+"?PageMode=GetAccountAddress&Socuid=" + Cstr(rsOppdrag("SOCuiD")) , false, Application("Xtra_Domain") +"\" + Application("Xtra_User"), Application("Xtra_Password")
							aXmlHTTP.send ""
							adresse = aXmlHTTP.responseText
							'postadresse = rsAdresse("zipcode") & " " & rsAdresse("city")
					else
						strSQL = "SELECT A.Adresse, PostAdr = A.Postnr + ' ' + A.Poststed " &_
							"FROM ADRESSE A " &_
							" WHERE " &_
							" A.AdresseRelID = " & rsOppdrag("FirmaID") &_
							" AND A.AdresseRelasjon = 1 " &_
							" AND A.AdresseType = 1 "
						
						set rsAdresse = GetFirehoseRS(strSQL, Conn)
						if (HasRows(rsAdresse)) then
							adresse = rsAdresse("Adresse").value
							postadresse = rsAdresse("PostAdr").value
							rsAdresse.close
						end if
						set rsAdresse = nothing
					end if
					%>
					<%=adresse%><br>
					<%=postadresse%>
				</td>
			</tr>
			<tr><td colspan="2">&nbsp;</td></tr>
			<tr>
				<th>Arbeidsadresse:</th>
				<td><%=rsOppdrag("ArbAdresse")%></td>
			</tr>
			<tr><td colspan="2">&nbsp;</td></tr>
			<tr>
				<th>Kontaktperson hos oppdragsgiver:</th>
				<% 
				'Set cts = server.CreateObject("Integration.SuperOffice")

										if(IsNull(rsOppdrag("Kontaktperson"))) then
											
											kontaktperson = GetCRMContactPersonName(rsOppdrag("SOPeID"))

										else
											kontaktperson = rsOppdrag("Kontaktperson")
										end if
										'set cts = nothing
				%>
				<td><%=kontaktperson%></td>
			</tr>
			<tr>
				<th>Kontaktperson hos Centric:</th>
				<td><%=rsOppdrag("Ansvarlig")%>, telefon <%=rsOppdrag("dirtlf")%></td>
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
				<th>Timelønn:</th>
				<td><%=rsOppdragVikar("Timeloenn")%></td>
			</tr>
			<tr>
				<th></th>
				<td></td>
			</tr>
		</table>
		
		<h2>Ytterligere oppdragsopplysninger:</h2>
		<p><%= replace(rsOppdragVikar.fields("BekreftelseKonsulentTekst").value, vbcrlf,"<br>")%><br/>
<%=rsOppdrag("Ansvarlig")%>
</p>

		
		
	</body>
</html>
<%
rsOppdrag.close
set rsOppdrag = nothing
rsOppdragVikar.close
set rsOppdragVikar = nothing
CloseConnection(Conn)
set Conn = nothing	
%>