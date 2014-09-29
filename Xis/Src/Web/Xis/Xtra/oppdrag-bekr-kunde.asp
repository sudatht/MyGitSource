<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="includes\CRM.Integration.Contact.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim strAction
	dim lOppdragID
	dim lOppdragVikarID
	dim strAffirmationText
	dim strSQL
	dim blnPrint
	dim strAdresse : strAdresse = ""
	dim strPostadresse : strPostadresse =""
	dim strKontaktperson
	dim vikarNameText
	dim Conn
	dim strKunde
    dim ShowCondition
	dim aXmlHTTP

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
	
	strAffirmationText = Request("txtAffirmation")
	if len(strAffirmationText) > 4000 then
		AddErrorMessage("Du har skrevet en for lang tekst!")
		call RenderErrorMessage()
	end if

	blnPrint = false
        ShowCondition = request("hiddenShowCondition")
	if request("hdnPosted")="1" then
	'This is postback..

		'What action was done?
		strAction = lcase(trim(request("hdnAction")))

		if (strAction="utskrift") or (strAction="lagre")  then
		'pop the printer window

		'Update task affirmation..
			strSQL = "UPDATE [oppdrag_vikar_confirmation_text] SET [CustomerText]=" & Quote( PadQuotes(strAffirmationText)) & _
			"WHERE [Oppdragvikarid] = " & lOppdragVikarID

			call ExecuteCRUDSQL(strSQL, Conn )
			
			'Prepare to pop the print window
			if (strAction="utskrift") then
				blnPrint = true
                              
                            
			end if

		elseif (strAction="vis") then
		'Return to view task page..
			response.redirect "WebUI/OppdragView.aspx?OppdragID=" & lOppdragID
		end if

	end if

	' Get information from database
	' Get oppdrags data
	Set rsOppdrag  = GetFirehoseRS("SELECT F.Firma, isnull(F.CRMAccountGuid,'') AS CRMAccountGuid, F.FirmaId, Ansvarlig = M.Fornavn+' '+M.Etternavn, KontaktPerson=K.Fornavn + ' ' + K.Etternavn,  " &_
	"O.Beskrivelse, O.ArbAdresse, O.Notatkunde, O.SOPeID, F.SOCuID " &_
	"FROM Oppdrag O, Firma F, Medarbeider M, KONTAKT K " &_
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
		aXmlHTTP.Open "POST", "http://xis/Xtra/CRMIntegration/Responder.aspx?PageMode=GetAccountAddress&Socuid=" + Cstr(rsOppdrag("SOCuID")) , false, Application("Xtra_Domain") +"\" + Application("Xtra_User"), Application("Xtra_Password")
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
	"Vikarnavn = VIKAR.Fornavn+' '+VIKAR.Etternavn, " & _
	"PostAdr = ADRESSE.Postnr+' '+ ADRESSE.Poststed, " & _
	"VIKAR_ANSATTNUMMER.ansattnummer " & _
	"FROM OPPDRAG_VIKAR_CONFIRMATION_TEXT, OPPDRAG_VIKAR " & _
	"LEFT OUTER JOIN VIKAR ON OPPDRAG_VIKAR.Vikarid = VIKAR.Vikarid " & _
	"LEFT OUTER JOIN ADRESSE ON OPPDRAG_VIKAR.Vikarid = ADRESSE.Adresserelid " & _
	"LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON OPPDRAG_VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " & _
	"WHERE OPPDRAG_VIKAR_CONFIRMATION_TEXT.OppdragVikarID = OPPDRAG_VIKAR.OppdragVikarID and OPPDRAG_VIKAR.Oppdragvikarid = '" & lOppdragVikarID & "' " & _
	"AND OPPDRAG_VIKAR.Statusid = '4' " & _
	"AND ADRESSE.Adresserelasjon = '2' " & _
	"AND ADRESSE.AdresseType = '1' "

	'Response.write strSQL
	Set rsOppdragVikar = GetFirehoseRS( strSQL, Conn )
	
	' Utvidelse
	If rsOppdragVikar( "Utvid" ) = 1 Then
		Heading = "Forlengelse av oppdrag"
	Else
		Heading = "Bekreftelse av oppdrag"
	End If
	
	vikarNameText = rsOppdragVikar("Vikarnavn")

' Create Oppdragsbeksreftelse
%>
<html>
	<head>
		<title><%=Heading%></title> 
		<%
		if (blnPrint=true) then
		%>
			<script language="javaScript" type="text/javascript">
				function printMe(vikarID, oppdragID,vikarName)
				{

					//window.open = "oppdrag-bekr-kunde-utskrift.asp?oppdragvikarid=<%=lOppdragVikarID%>&oppdragID=<%=lOppdragID%>";
					window.open('oppdrag-bekr-kunde-utskrift.asp?oppdragvikarid=' + vikarID + '&vikarname=' + vikarName + '&oppdragID=' + oppdragID +'&ShowCondition=' + <%=ShowCondition%> , '', 'menubar=yes,toolbar=yes,status=yes,scrollbars=yes');
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
	<body <%if blnPrint=true then%>onLoad="javascript:printMe('<%=lOppdragVikarID%>', '<%=lOppdragID%>','<%=vikarNameText%>')"<%end if%>>
		<div class="pageContainer" id="pageContainer">
			<form action="oppdrag-bekr-kunde.asp" method="post" id="frmJobConfirmation">
				<input type="hidden" name="hdnPosted" value="1">
				<input type="hidden" name="OppdragVikarID" value="<%=lOppdragVikarID%>">
				<input type="hidden" name="OppdragID" value="<%=lOppdragID%>">
				<input type="hidden" name="hdnAction" value="Lagre">
                                <input type="hidden" name="hiddenShowCondition" value="false">
				<div class="contentHead1">
					<h1><%=Heading%></h1>
					<div class="contentMenu">
						<table cellpadding="0" cellspacing="0" width="96%">
							<tr>
								<td>
									<table cellpadding="0" cellspacing="2">
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
												<a onClick="javascript:document.all.hdnAction.value='utskrift'; document.all.frmJobConfirmation.submit()" href="#" title="Lagre og skriv ut oppdragsbekreftelse">
												<img src="/xtra/images/icon_print.gif" width="18" height="15" alt="" align="absmiddle">Skriv ut/Send</a>
											</td>
<!--
<td class="menu" >  <input ID="chkShowCondition" type="CheckBox" <% if ShowCondition ="true"  then Response.write("Checked=""true""") End if%>/> Utelat standard betingelser </td> 
-->
										</tr>
									</table>
								</td>
								<td class="right"><!--#include file="Includes/contentToolsMenu.asp"--></td>
							</tr>
						</table>
					</div>
				</div>
				<!--IMG align=right height=74 src="./images/logo1.gif" width=300-->
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
										<th>StartDato:</th>
										<td><%=rsOppdragVikar("Fradato")%></td>
									</tr>
									<tr>
										<th>SluttDato:</th>
										<td><%=rsOppdragVikar("Tildato")%></td>			
									<tr>
										<th>Arbeidstid:</th>
										<td><%=FormatDateTime( rsOppdragVikar("FraKl"), 4)%> - <%=FormatDateTime( rsOppdragVikar("TilKl"), 4)%></td>
									</tr>
									<tr>
										<th>Timepris:</th>
										<td><%=FormatNumber(rsOppdragVikar("Timepris"),2)%> (eks.mva)</td>
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
										<th></th>
										<td>
										<%
											linkurl = Application("CRMAccountLink") & rsOppdrag("CRMAccountGuid") & "%7d&pagetype=entityrecord"										
											strKunde = "<a href=" & linkurl & " target='_blank'>" & rsOppdrag("Firma").Value & " </a>"										
										%>
											
											<%=strKunde%><br>
											<%=strAdresse%><br>
											<%=strPostadresse%><br>
											Attn: <%=strKontaktperson%><br>
										</td>
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
										<td><%=rsOppdragVikar("Vikarnavn")%></td>
									</tr>
								</table>					
							</td>	
						</tr>
					</table>
					<h3>Oppdragsbeskrivelse:</h3>
					<%=rsOppdrag("Beskrivelse")%>
				</div>
				<div class="contentHead"><h2>Ytterligere oppdragsopplysninger:</h2></div>
				<div class="content">
					<%=rsOppdrag("Notatkunde")%>
					<TEXTAREA style="height:420px;" name="txtAffirmation"><%=rsOppdragVikar.fields("BekreftelseKundeTekst").value%></TEXTAREA>
				</div>
			</form>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>