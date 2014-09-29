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

	dim strAction
	dim profil
	dim lOppdragID
	dim lOppdragVikarID
	dim strAffirmationText
	dim blnPrint
	dim strSQL
	dim Conn
	dim rsAdresse
	dim kontaktperson
	dim rsOppdrag
	dim cts
	dim adresse : adresse = vbNullString
	dim postadresse : postadresse = vbNullString
	dim bRegisterActivity
	
	bRegisterActivity = false
	
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
	strAffirmationText = Request("txtAffirmation")
	
	if len(strAffirmationText) > 2000 then
		AddErrorMessage("Du har skrevet en for lang tekst. Maks antall tegn er 2000!")
		call RenderErrorMessage()
	end if

	blnPrint = false

	if request("hdnPosted") = "1" then
	'This is postback..

		'What action was done?
		strAction = lcase(trim(request("hdnAction")))
		if (strAction="utskrift") or (strAction="lagre")  then
		'pop the printer window

		'Update task affirmation..
			strSQL = "UPDATE [oppdrag_vikar_confirmation_text] SET [ConsultantText]=" & Quote( PadQuotes(strAffirmationText)) & _
			"WHERE [Oppdragvikarid] = " & lOppdragVikarID

			if (ExecuteCRUDSQL(strSQL, Conn) = false) then
				CloseConnection(Conn)
				set Conn = nothing	   
				AddErrorMessage("Feil under oppdatering av bekreftelsestekst.")
				call RenderErrorMessage()
			end if
			
			bRegisterActivity = true
			

			'Prepare to pop the print window
			if (strAction="utskrift") then
				blnPrint = true
			end if

		elseif (strAction="vis") then
		'Return to view task page..
			response.redirect "WebUI/OppdragView.aspx?OppdragID=" & lOppdragID
		end if

	end if

	' Get oppdrags data
	strSQL = "SELECT F.Firma, F.FirmaID, F.SOCuID, Ansvarlig = M.Fornavn + ' ' + M.Etternavn, M.dirtlf, " &_ 
	"KontaktPerson = K.Fornavn + ' ' + K.Etternavn, O.SOPeID, " &_
	"O.Beskrivelse, O.ArbAdresse, O.AnsMedID " &_
	"FROM Oppdrag O, Firma F, Medarbeider M, KONTAKT K  " &_
	" WHERE O.OppdragID = " & lOppdragID & _
	" AND O.FirmaID = F.FirmaID " &_
	" AND O.AnsMedID *= M.MedID " &_
	" AND O.BestilltAv *= K.KontaktID "

	Set rsOppdrag  = GetFirehoseRS(strSQL, Conn)

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

	If rsOppdragVikar( "Utvid" ) = 1 Then
	   Heading = "Forlengelse av oppdrag"
	Else
	   Heading = "Bekreftelse av oppdrag"
	End If

	'Opplysninger om avdelingskontor
	strSQL = "SELECT ak.telefon FROM avdelingskontor ak, oppdrag op WHERE ak.id = op.avdelingskontorID AND op.oppdragid = " & lOppdragID
	set rsAvdKontor = GetFirehoseRS(strSQL, Conn)

	if bRegisterActivity then
	
		'commission confirmation. When the Print/Send - button is pressed
		Dim rsActivityType
		Dim nActivityTypeID
		Dim sDate
		Dim strComment
		strSQL = "SELECT AktivitetTypeID FROM H_AKTIVITET_TYPE WHERE AktivitetType = 'Oppdr. bekr.'"
		set rsActivityType = GetFirehoseRS(strSQL, Conn)
		nActivityTypeID = CInt(rsActivityType("AktivitetTypeID"))
		' Close and release recordset
		rsActivityType.Close
		Set rsActivityType = Nothing
		strComment = "Oppdragsbekreftelse skrevet ut/sendt."
		sDate = GetDateNowString()
		strSql = "INSERT INTO AKTIVITET( AktivitetTypeID, AktivitetDato, VikarID, FirmaID, Notat, OppdragID, Registrertav, RegistrertAvID, AutoRegistered ) Values(" & nActivityTypeID & ",'" & sDate & "'," & rsOppdragVikar("Vikarid") & "," & rsOppdrag("FirmaID") & ",'" & strComment & "'," & lOppdragID & ",'" & Session("Brukernavn") & "'," & Session("medarbID") & ", 1)"
		if (ExecuteCRUDSQL(strSQL, Conn) = false) then
			response.write(strSql)
			CloseConnection(Conn)
			set Conn = nothing	   
			AddErrorMessage("Aktivitetsregistrering for oppdragsbekreftelse sendt feilet.")
			call RenderErrorMessage()
		end if
			
	end if
	
	' Render Oppdragsbeksreftelse
	%>
<html>
	<head>
		<title><%=Heading%></title>
		<%
		if (blnPrint = true) then
			%>
			<script language="javaScript" type="text/javascript">
				function printMe(vikarID, oppdragID)
				{
					
					window.open('oppdrag-bekr-vikar-utskrift.asp?oppdragvikarid=' + vikarID + '&oppdragID=' + oppdragID , '', 'menubar=yes,toolbar=yes,status=yes,scrollbars=yes');
					//window.location = "oppdrag-bekr-vikar-utskrift.asp?oppdragvikarid=<%=lOppdragVikarID%>&oppdragID=<%=lOppdragID%>";
				}
			</script>
			<%
		end if
		%>
		<script type="text/javascript" src="/xtra/js/contentMenu.js"></script>
		<script type="text/javascript" src="/xtra/js/fontSizer.js"></script>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	</head>
	<body <%if blnPrint=true then%>onLoad="javascript:printMe('<%=lOppdragVikarID%>', '<%=lOppdragID%>')"<%end if%>>
		<div class="pageContainer" id="pageContainer">
			<form action="oppdrag-bekr-vikar.asp" method="post" id="frmJobConfirmation">
				<input type="hidden" name="hdnPosted" value="1">
				<input type="hidden" name="OppdragVikarID" value="<%=lOppdragVikarID%>">
				<input type="hidden" name="OppdragID" value="<%=lOppdragID%>">
				<input type="hidden" name="hdnAction" value="Lagre">
				<div class="contentHead1">
					<h1><%=Heading%></h1>
					<div class="contentMenu">
						<table cellpadding="0" cellspacing="0" width="96%">
							<tr>
								<td>
									<table cellpadding="0" cellspacing="2" ID="Table1">
										<tr>
											<td class="menu" id="menu1" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
												<a onClick="javascript:document.all.frmJobConfirmation.submit()" href="#" title="Lagre">
												<img src="/xtra/images/icon_save.gif" width="18" height="15" alt="" align="absmiddle">Lagre</a>
											</td>
											<td class="menu" id="menu2" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
												<a onClick="javascript:document.all.hdnAction.value='vis';document.all.frmJobConfirmation.submit()" href="#" title="Vis oppdrag">
												<img src="/xtra/images/icon_job.gif" width="18" height="15" alt="" align="absmiddle">Vis
											</td>							
											<td class="menu" id="menu3" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
												<a onClick="javascript:document.all.hdnAction.value='utskrift';document.all.frmJobConfirmation.submit()" href="#" title="Lagre og skriv ut oppdragsbekreftelse">
												<img src="/xtra/images/icon_print.gif" width="18" height="15" alt="" align="absmiddle">Skriv ut/Send</a>
											</td>
										</tr>
									</table>
								</td>
								<td class="right"><!--#include file="Includes/contentToolsMenu.asp"--></td>
							</tr>
						</table>
					</div>
				</div>
				<div class="content">
					<table width="96%">
						<col width="33%">
						<col width="33%">
						<col width="33%">
						<tr>
							<td>
								<table>
									<tr>
										<th>Oppdragsnr:</th>
										<td><%=lOppdragID%></td>
									</tr>
									<tr>
										<th>Dato:</th>
										<td><%=date()%></td>
									</tr>
									<tr>
										<th>Arbeidsadresse:</th>
				 						<td><%=rsOppdrag("ArbAdresse")%></td>
									</tr>
									<tr>
										<th>StartDato:</th>
										<td><%=rsOppdragVikar("Fradato")%></td>
									</tr>
									<tr>
										<th>SluttDato:</th>
										<td><%=rsOppdragVikar("Tildato")%></td>
									</tr>
									<tr>
										<th>Arbeidstid:</th>
										<td><%=FormatDateTime( rsOppdragVikar("FraKl"), 4)%> - <%=FormatDateTime( rsOppdragVikar("TilKl"), 4)%></td>
									</tr>
									<tr>
										<th>Timelønn:</th>
										<td><%=FormatNumber(rsOppdragVikar("Timeloenn"),2)%></td>
									</tr>
								</table>
							</td>
							<td>
								<table>
									<tr>
										<th>Kontaktnr:</th>
										<td><%=rsOppdrag("FirmaID")%></td>
									</tr>
									<tr>
										<th>Kontakt:</th>
										<td>
											<%=rsOppdrag("Firma")%><br>
											<%											
											if ( not IsNull(rsOppdrag("SOCuid").value)) Then
												Set aXmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
												aXmlHTTP.Open "POST", "http://xis/Xtra/CRMIntegration/Responder.aspx?PageMode=GetAccountAddress&Socuid=" + Cstr(rsOppdrag("SOCuID")) , false, Application("Xtra_Domain") +"\" + Application("Xtra_User"), Application("Xtra_Password")
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
									<tr>
										<th>Kontaktperson:</th>
										<% 
										'Set cts = server.CreateObject("Integration.SuperOffice")

										if(IsNull(rsOppdrag("Kontaktperson"))) then
											'dim personRs 		
											'set personRs = cts.GetPersonSnapshotById(clng(rsOppdrag("SOPeID")))
											'if (not personRs.EOF) then
											'	if (isnull(personRs("middlename"))) then
											'		kontaktperson = personRs("firstname") & " " & personRs("lastname")
											'	else
											kontaktperson = GetCRMContactPersonName(rsOppdrag("SOPeID"))
											'	end if
											'end if
											'personRs.close
											'set personRs = nothing
										else
											kontaktperson = rsOppdrag("Kontaktperson")
										end if
										'set cts = nothing
										%>
										<td><%=kontaktperson%></td>
									</tr>
								</table>
							</td>
							<td>
								<table>
									<tr>
										<th>Ansattnummer:</th>
										<td><%=rsOppdragVikar("ansattnummer")%></td>
									</tr>
									<tr>
										<th>Ansatt:</th>
										<td>
											<%=rsOppdragVikar("Vikarnavn")%><br>
											<%=rsOppdragVikar("Adresse")%><br>
											<%=rsOppdragVikar("PostAdr")%>
										</td>
									</tr>
									<tr>
										<th>Kontaktperson hos Xtra:</th>
										<td><%=rsOppdrag("Ansvarlig")%> på telefon <%=rsOppdrag("dirtlf")%></td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
					<h2>Oppdragsbeskrivelse:</h2>
					<p><%=rsOppdrag("Beskrivelse")%></p>
				</div>
				<div class="contentHead">
					<h2>Ytterligere oppdragsopplysninger</h2>
				</div>
				<div class="content">
					<textarea style="height:250px;" name="txtAffirmation"><%=rsOppdragVikar.fields("BekreftelseKonsulentTekst").value%></textarea>
				</div>
				</form>
			</div>
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