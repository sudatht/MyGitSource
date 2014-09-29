<%@ LANGUAGE="VBSCRIPT" %>
<%
Option Explicit
%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Settings.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Settings.inc"-->
<!--#INCLUDE FILE="..\includes\Validations.inc"-->
<%

	dim objResource

	set objResource = Server.CreateObject("Localizer.ResourceManager")
	
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
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
	Dim strSQL
	Dim bSendEpost				'As Boolean
	Dim strEpostResultat		'As String
	Dim strSkattekortSQL		'As String
	Dim rsTimesheet				'As Recordset
	Dim strSenderAdress			'As String
	Dim strEpostBody			'As String
	Dim strEpostSubject			'As String
	Dim strRemoteHost			'As String
	Dim strBrukerSQL			'As String
	Dim rsBrukerInfo			'As Recordset
	Dim strBrukerFornavn		'As String
	Dim strBrukerEtternavn		'As String
	Dim strAdress
	Dim lSelectedResponsibleID

	strRemoteHost = Application("XtraMailServer")

	'Set objEFCommon = server.createobject("efcommon.DataValidation")

	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	' Get the selected person id from the query string
	if(request.form("selectedMedId") = "") then
		lSelectedResponsibleID		= Request.QueryString("MedId")
	else
		lSelectedResponsibleID = request.form("selectedMedId")
	end if
	
	'Getting user info
	strBrukerSQL = "SELECT BRUKER.Id, " & _
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
		strEpostBody = GetSetting("TimesheetReminderEMailBody")
		if (len(strBrukerFornavn) > 0 AND len(strBrukerEtternavn)>0) then
			strEpostBody = strEpostBody & strBrukerFornavn & " " & strBrukerEtternavn
		end if
	End If

	strEpostSubject = GetSetting("TimesheetReminderEMailSubject")
	bSendEpost = Request.Form("bolSendEpost").Item
	set rsTimesheet = GetFirehoseRS("EXEC [dbo].[GetNonApprovedtimesheetsForConsultantLeader] " & lSelectedResponsibleID, Conn)	
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
			//lager felles variabler
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

			// her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;

			function sendTimeSheetReminder()
			{
				if (window.confirm("Vil du sende e-post til alle på denne listen?"))
				{
					document.frmTimesheet.bolSendEpost.value = "true";
					document.frmTimesheet.submit();
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
				<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" name="frmTimesheet" method="POST" ID="Form1">
					<input type="hidden" name="bolSendEpost" id="bolSendEpost" value="">
					<input type="hidden" name="selectedMedId" id="selectedMedId" value="<%= lSelectedResponsibleID %>">
					<table cellpadding='0' cellspacing='0' ID="Table1">
						<tr>
							<td><textarea name="eposttekst" rows="10" cols="50" ID="Textarea1"><%=strEpostBody%></textarea></td>
						</tr>
						<tr>
							<td>
								<input type="submit" id="submitepost" name="submitepost" value="Send" style="display:none">
								<input type="button" name="sendepost" id="sendepost" value="<%objResource.WriteText("SendReminder")%>" onclick="sendTimeSheetReminder()">
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
						If (rsTimesheet.EOF) Then
							Response.Write "<tr>" & chr(13)
							Response.Write "<td colspan='6'>" & chr(13)
							Response.Write "Fant ingen vikarer med manglende timelister tilknyttet denne brukeren." & chr(13)
							Response.Write "</td>" & chr(13)
						Else
							Do While Not rsTimesheet.EOF
								strEpostResultat = ""

								' Sender Epost til vikarer
								If (bSendEpost = "true") Then
									If (isValidEmail(Trim(Cstr(rsTimesheet("Epost").Value)))) Then
										strEpostResultat = sendMail(strBrukerFornavn & " " & strBrukerEtternavn, strSenderAdress, rsTimesheet("vikar").Value, Trim(Cstr(rsTimesheet("Epost").Value)), strEpostBody, strEpostSubject, strRemoteHost)
									Else
										strEpostResultat = "Ikke sendt!"
									End If
								End If

								strAdress = ""
								if (len(rsTimesheet("Adresse").Value) > 0) then
									strAdress = rsTimesheet("Adresse").Value
								end if
								if (len(rsTimesheet("postnr").Value) > 0) then
									if (len(strAdress)>0) then
										strAdress = strAdress & ", "
									end if
									strAdress = strAdress & rsTimesheet("postnr").Value
									if (len(rsTimesheet("PostSted").Value) > 0) then
										strAdress = strAdress  & " " & rsTimesheet("PostSted").Value
									end if
								end if
								%>
								<tr>
									<td align="center">
										<%=rsTimesheet("ansattnummer").Value%>
									</td>
									<td>
										<%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "vikarvis.asp?VikarID=" & rsTimesheet( "VikarID" ), rsTimesheet( "vikar"), "Vis vikar " & rsTimesheet("vikar") )%>
									</td>
									<td><%=strAdress %>&#160;</td>
									<td class="nowrap"><%=rsTimesheet("Telefon").Value%>&#160;</td>
									<td class="nowrap"><%=rsTimesheet("MobilTlf").Value%>&#160;</td>
									<td><a href="mailto:<%=rsTimesheet("Epost").Value%>"><%=rsTimesheet("Epost").Value%></a>&#160;</td>
									<%
									If (bSendEpost = "true") Then
										Response.Write "<td>" & strEpostResultat & "</td>"
									End If
									%>
								</tr>
								<%
								rsTimesheet.MoveNext
							Loop
						End If
						rsTimesheet.Close()
						set rsTimesheet = nothing
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
