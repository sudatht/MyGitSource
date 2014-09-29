<%@ LANGUAGE="VBSCRIPT" %>
<%
Server.ScriptTimeout = 120
%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Settings.inc"-->
<!--#INCLUDE FILE="..\includes\Validations.inc"-->
<%

      dim objResource

	set objResource = Server.CreateObject("Localizer.ResourceManager")

	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	Function sendMail(strSenderName, strSenderAdress, strReceiver, strReceiverAdress, strBody, strSubject, strRemoteHost)

		Dim objMailer 'as SMTPsvg.Mailer

		'Initiate Serverobjects' ASPMail object
		Set ObjMailer = Server.CreateObject("Persits.MailSender")

		If IsObject(ObjMailer) Then

			ObjMailer.FromName	  = strSenderName					'Name of E-mail sender
			ObjMailer.From = strSenderAdress					'Adress E-mail was sent from
			ObjMailer.Host  = strRemoteHost					'Name of mail server used to send E-mail
			ObjMailer.AddAddress strReceiverAdress	'Name & Adress of receiver of E-mail
			ObjMailer.Subject     = strSubject						'Subject of E-mail
			ObjMailer.Body    = strBody							'Body of E-mail
			'ObjMailer.CharSet 	 = 2									'Set mail encoding to ISO-8859-1

			If ObjMailer.Send Then
				sendMail = "Ok"
			Else
				sendMail = "Send Fail"
			End If
			Set ObjMailer = Nothing
		Else
			sendMail = "Couldn't create mail object."
		End If

	End Function

	'Dim objEFCommon				'As objEFCommon
	Dim Conn					'As ADODB.Connection
	Dim iTeller					'As Integer
	Dim iNummer					'As Integer

	Dim bSendEpost				'As Boolean
	Dim strEpostResultat		'As String
	Dim strSkattekortSQL		'As String
	Dim rsSkattekort			'As Recordset
	Dim strSenderAdress			'As String
	Dim strEpostBody			'As String
	Dim strEpostSubject			'As String
	Dim strRemoteHost			'As String
	Dim strBrukerSQL			'As String
	Dim rsBrukerInfo			'As Recordset
	Dim strBrukerFornavn		'As String
	Dim strBrukerEtternavn		'As String

	Dim strSQL
	dim strAdress

	strRemoteHost = Application("XtraMailServer")

	'Set objEFCommon = server.createobject("efcommon.DataValidation")

	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	'Getting user info
	strBrukerSQL = "SELECT " & _
				"BRUKER.Id, " & _
				"MEDARBEIDER.Fornavn, " & _
				"MEDARBEIDER.Etternavn, " & _
				"VIKAR.Epost " & _
				"FROM BRUKER " & _
				"LEFT OUTER JOIN MEDARBEIDER ON BRUKER.Medarbid = MEDARBEIDER.Medid " & _
				"LEFT OUTER JOIN VIKAR ON MEDARBEIDER.Vikarid = VIKAR.Vikarid " & _
				"WHERE BRUKER.Id = '" & Session("brukerID") & "' "

	set rsBrukerInfo = GetFirehoseRS(strBrukerSQL, Conn)

	If Not (rsBrukerInfo.EOF) Then
		strBrukerFornavn = rsBrukerInfo("Fornavn").Value
		strBrukerEtternavn = rsBrukerInfo("Etternavn").Value
		strSenderAdress = rsBrukerInfo("Epost").Value
	End If
	rsBrukerInfo.close
	Set rsBrukerInfo = Nothing

	If (Request.Form("eposttekst").Item <> "") Then
		strEpostBody = Request.Form("eposttekst").Item
	Else
		strEpostBody = GetSetting("TaxCardReminderEMailBody")
		if (len(strBrukerFornavn) > 0 AND len(strBrukerEtternavn) > 0) then
			strEpostBody = strEpostBody & VBCRLf & strBrukerFornavn & " " & strBrukerEtternavn
		end if
	End If

	strEpostSubject = GetSetting("TaxCardReminderEMailSubject")

	bSendEpost = Request.Form("bolSendEpost").Item

	strSkattekortSQL = "SELECT DISTINCT " & _
			"VIKAR.VikarID, " & _
			"VIKAR.Fornavn, " & _
			"VIKAR.Etternavn, " & _
			"VIKAR.Telefon, " & _
			"VIKAR.MobilTlf, " & _
			"VIKAR.Epost, " & _
			"VIKAR.MottattSkattekort, " & _
			"ADRESSE.Adresse, " & _
			"ADRESSE.Postnr, " & _
			"ADRESSE.Poststed, " & _
			"VIKAR_ANSATTNUMMER.ansattnummer " & _
			"FROM VIKAR " & _
			"LEFT OUTER JOIN ADRESSE ON VIKAR.Vikarid = ADRESSE.adresseRelID " & _
			"LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " & _
			"INNER JOIN VIKAR_UKELISTE ON VIKAR.Vikarid = VIKAR_UKELISTE.Vikarid " & _
			"WHERE VIKAR.Statusid = '3' " & _
			"AND VIKAR.typeID = 1" & _
			"AND ADRESSE.AdresseRelasjon = '2' " & _
			"AND ADRESSE.AdresseType = '1' " & _
			"AND VIKAR.AnsMedID = " & Session("medarbID") & _
			"AND ( " & _
			"( VIKAR.MottattSkattekort IS NULL ) " & _
			"OR ( VIKAR.MottattSkattekort = '-1' ) " & _
			"OR ( VIKAR.VikarId NOT IN (  " & _
				"SELECT " & _
				"VIKAR.Vikarid " & _
				"FROM VIKAR " & _
				"WHERE VIKAR.MottattSkattekort = YEAR(GETDATE()) " & _
				"OR " & _
				"( VIKAR.MottattSkattekort = (YEAR(GETDATE()) - 1) AND (MONTH(GETDATE()) = 1) ) " & _
			"))) " & _
			"ORDER BY VIKAR.Etternavn ASC "

	set rsSkattekort = GetFirehoseRS(strSkattekortSQL, Conn)
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<title>Manglende skattekort</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script language='javascript' src='../js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='../js/navigation.js' id='navigationScripts'></script>
		<script language="javaScript" type="text/javascript">
			function shortKey(e)
			{
				var keyChar = String.fromCharCode(event.keyCode);
				var modKey  = event.ctrlKey;
				var modKey2 = event.shiftKey;

				<%
				If HasUserRight(ACCESS_TASK, RIGHT_READ) Then
					%>
					if (modKey && modKey2 && keyChar=="P")
					{
						parent.frames[funcFrameIndex].location=("/xtra/hotlist/hotlistOppdrag.asp");
					}
					<%
				end if
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>
					if (modKey && modKey2 && keyChar=="V")
					{
						parent.frames[funcFrameIndex].location=("/xtra/hotlist/hotlistVikar.asp");
					}
					<%
				end if
				%>
			}

			function sendskattekortepost()
			{
				if (window.confirm("Vil du sende e-post til alle på denne listen?"))
				{
					document.frmSkattekort.bolSendEpost.value = "true";
					document.frmSkattekort.submit();
				}
			}
		</script>
	</head>
	<body onLoad="fokus()">
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1><%objResource.WriteText("Heading")%></h1>
			</div>
			<div class="content">
				<a id="Top"></a>
				<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" name="frmSkattekort" method="POST" ID="Form1">
					<input type="hidden" name="bolSendEpost" id="bolSendEpost" value="">
					<table cellpadding='0' cellspacing='0' ID="Table1">
						<tr>
							<td><textarea name="eposttekst" rows="10" cols="50" ID="Textarea1"><%=strEpostBody%></textarea></td>
						</tr>
						<tr>
							<td>
								<input type="submit" id="submitepost" name="submitepost" value="Send" style="display:none">
								<input type="button" name="sendepost" id="sendepost" value="<%objResource.WriteText("SendReminder")%>" onclick="sendskattekortepost()">
							</td>
						</tr>
					</table>
				</form>
				<div class="listing">
					<table cellspacing="1" cellpadding="1" ID="Table2">
						<tr>
							<th><%objResource.WriteText("AnsattNo")%></th>
							<th><%objResource.WriteText("Vikar")%></th>
							<th><%objResource.WriteText("Address")%></th>
							<th><%objResource.WriteText("Telephone")%></th>
							<th><%objResource.WriteText("Mobile")%></th>
							<th><%objResource.WriteText("Epost")%></th>
							<%
							If (bSendEpost = "true") Then
								%>
								<th>Resultat</th>
								<%
							End If
							%>
						</tr>
						<%
						If (rsSkattekort.EOF) Then
							Response.Write "<tr>" & chr(13)
							Response.Write "<td colspan='6'>" & chr(13)
							Response.Write "Fant ingen vikarer med manglende skattekort tilknyttet denne brukeren." & chr(13)
							Response.Write "</td>" & chr(13)
						Else
							Do While Not rsSkattekort.EOF
								strEpostResultat = ""
								' Sender Epost til vikarer
								If (bSendEpost = "true") Then
									If (isValidEmail(Trim(Cstr(rsSkattekort("Epost").Value)))) Then
										strEpostResultat = sendMail(strBrukerFornavn & " " & strBrukerEtternavn, strSenderAdress, rsSkattekort("fornavn").Value, Trim(Cstr(rsSkattekort("Epost").Value)), strEpostBody, strEpostSubject, strRemoteHost)
									Else
										strEpostResultat = "Ugyldig Epost adresse"
									End If
								End If

								strAdress = ""
								if (len(rsSkattekort("Adresse").Value) > 0) then
									strAdress = rsSkattekort("Adresse").Value
								end if
								if (len(rsSkattekort("postnr").Value) > 0) then
									if (len(strAdress)>0) then
										strAdress = strAdress & ", "
									end if
									strAdress = strAdress & rsSkattekort("postnr").Value
									if (len(rsSkattekort("PostSted").Value) > 0) then
										strAdress = strAdress  & " " & rsSkattekort("PostSted").Value
									end if
								end if
								%>
								<tr>
									<td align="center">
										<%
										If (rsSkattekort("ansattnummer").Value <> "" ) Then
											Response.Write rsSkattekort("ansattnummer").Value
										Else
											Response.Write "---"
										End If
										%>
									</td>
									<td>
										<%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "vikarvis.asp?VikarID=" & rsSkattekort( "VikarID" ), rsSkattekort("etternavn").Value & ", " & rsSkattekort("fornavn").Value, "Vis vikar " & rsSkattekort("fornavn").Value & " " & rsSkattekort("etternavn").Value )%>
									</td>
									<td><%=strAdress%>&#160;</td>
									<td class="nowrap"><%=rsSkattekort("Telefon").Value%>&#160;</td>
									<td class="nowrap"><%=rsSkattekort("MobilTlf").Value%>&#160;</td>
									<td><a href="mailto:<%=rsSkattekort("Epost").Value%>"><%=rsSkattekort("Epost").Value%></a>&#160;</td>
									<%
									If (bSendEpost = "true") Then
										Response.Write "<td>" & strEpostResultat & "</td>"
									End If
									%>
								</tr>
								<%
								rsSkattekort.MoveNext
							Loop
						End If
						rsSkattekort.Close()
						%>
					</table>
					<a href="#top"><img src="../Images/icon_GoToTop.gif" alt="<%objResource.WriteText("Top")%>"><%objResource.WriteText("Top")%><a>
				</div>
			</div>
		</div>
	</body>
</html>
<%
set objResource = Nothing
CloseConnection(Conn)
set Conn = nothing
%>
