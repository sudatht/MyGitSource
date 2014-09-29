<%@ Language=VBScript %>

<%option explicit%>

<%
dim cvfil 'as string
dim fulltime
dim beskrivelse
dim IPadr
dim visJobgr
dim visProdgr
dim strNavn
dim strtekst
dim lVikarID
dim strVikarID
dim strCon
dim objCons
dim objAdr
dim iGodkjenn
dim objCv
dim objCv2
dim edu
dim edu2
dim allEdu
dim allEdu2
dim strKomm
dim dataval
dim forste
dim i
dim vis_kvittering
dim strKvittering

if Request.QueryString("VikarID") <> "" then
	lVikarID = Request.QueryString("VikarID")
end if

strVikarID = Cstr(lVikarID)
iGodkjenn = "1"
strKvittering = ""

'FJM.Kan du vær så snill å erstatte disse med tilsvarende obj navn?
'Vær vennlig å gi andre navn enn navn2.

set objCons	= Server.CreateObject("XtraWeb.Consultant")
set edu		= Server.CreateObject("XtraWeb.Education")
set edu2	= Server.CreateObject("XtraWeb.Education")

objCons.XtraConString = Application("Xtra_intern_ConnectionString")
objCons.XtraDataShapeConString = Application("ConXtraShape")
objCons.GetConsultant(lVikarID)

set objCv	= objCons.CV
objCv.XtraConString = Application("Xtra_intern_ConnectionString")
objCv.XtraDataShapeConString = Application("ConXtraShape")
objCv.Refresh

set objCv2	= objCons.ChangedCV
objCv2.XtraConString = Application("Xtra_intern_ConnectionString")
objCv2.XtraDataShapeConString = Application("ConXtraShape")
objCv2.Refreshchanged

objCv2.Educations.RefreshChanged

set allEdu		= objCv.Educations
set allEdu2		= objCv2.Educations

vis_kvittering = false

if Request.Form("rerun") = "1" then

	strKomm = Request.Form("kommentar")

	' Ferdig
	if Request.Form("tilbake") = " Avbryt " then
		set objCv2		= nothing
		set objCv		= nothing
		set objAdr		= nothing
		objCons.ChangedCV.cleanup
		objCons.CV.cleanup
		objCons.cleanup
		set objCons	= nothing
		Response.Redirect("vikarCVvis.asp?VikarID=" & strVikarId)
	end if

	if Request.Form("aapne") = "       Åpne CV for vikaren       " then
		objCv.UnlockCV
		set objCv2		= nothing
		set objCv		= nothing
		set objAdr		= nothing
		objCons.ChangedCV.cleanup
		objCons.CV.cleanup
		objCons.cleanup
		set objCons	= nothing
		Response.Redirect("vikarCVvis.asp?VikarID=" & strVikarId)
	end if

	if Request.Form("mail") = "Send kommentar til vikaren" then
	%>
		<!--#include file="includes/vikarCvBehandleMailInclude.inc"-->
	<%
	end if

	if Request.Form("ok") = "  Utfør  " then
	
		i = 1
		for each edu2 in allEdu2
			if Request.Form("utdCheck" & i) = "1" then
				if Request.Form("godkjenn") = "1" then
					if not objCv2.Educations.ApproveChanged(edu2.DataValues("DataId")) then
						Response.Write "Godkjennes ikke!<br>"
					end if
				elseif Request.Form("godkjenn") = "2" then
					if not objCv2.Educations.RejectChanges(edu2.DataValues("DataId")) then
						Response.Write "Avvises ikke!<br>"
					end if
				else
					Response.Write "Feil!"
				end if
			end if
		i = i + 1
		next
		set objCv2		= nothing
		set objCv		= nothing
		set objAdr		= nothing
		objCons.ChangedCV.cleanup
		objCons.CV.cleanup
		objCons.cleanup
		set objCons	= nothing
		Response.Redirect "vikarCVbehandleUtdannelse.asp?type=4&VikarID=" & strVikarID
		
	end if

end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"
    "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <meta http-equiv="Content-Style-Type" content="text/css">
    <meta http-equiv="Content-Script-Type" content="text/javascript">
    <meta name="Developer" content="Electric Farm ASA">
	<meta name="generator" content="Microsoft Visual Studio 6.0">
	<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
</head>
<body>
	<div class="pageContainer" id="pageContainer">

<h3>CV godkjenning - <%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & strVikarID, Cons.DataValues("Fornavn") & " " & Cons.DataValues("Etternavn"), "Vis vikar " & Cons.DataValues("Fornavn") & " " & Cons.DataValues("Etternavn") )%></h3>

<table cellpadding='0' cellspacing='0'>
	<tr>
		<td>&nbsp;<strong>|</strong>&nbsp;<a href="vikarCVbehandlePersonalia.asp?VikarID=<%=strVikarID%>"><strong>Personalia</strong></a></td>
		<td>&nbsp;<strong>|</strong>&nbsp;<a href="vikarCVbehandleJobbonsker.asp?VikarID=<%=strVikarID%>"><strong>fagkompetanse</strong></a></td>
		<td>&nbsp;<strong>|</strong>&nbsp;<a href="vikarCVbehandleKvalifikasjoner.asp?VikarID=<%=strVikarID%>"><strong>Produktkompetanse</strong></a></td>
		<td>&nbsp;<strong>|</strong>&nbsp;<strong>Utdannelse</strong></td>
		<td>&nbsp;<strong>|</strong>&nbsp;<a href="vikarCVbehandlePraksis.asp?VikarID=<%=strVikarID%>"><strong>Praksis</strong></a></td>
		<td>&nbsp;<strong>|</strong>&nbsp;<a href="vikarCVbehandleReferanser.asp?VikarID=<%=strVikarID%>"><strong>Referanser</strong></a>&nbsp;<strong>|</strong>&nbsp;</td>
	</tr>
</table>

<p>

<form name="behandle" action="vikarCVbehandleUtdannelse.asp?VikarID=<%=strVikarID%>" method="post">

	<table cellpadding='0' cellspacing='0'>
		<tr>
			<td>
				<table cellpadding='0' cellspacing='0'>
					<tr>
						<td colspan="4"><strong>Utdannelse - Godkjente</strong></td>
					</tr>
					<%
					for each edu in allEdu
					%>
					<tr>
						<td><%if edu.DataValues("FromMonth") < 10 then%>0<%end if%>
							<%=edu.DataValues("FromMonth")%>/<%=edu.DataValues("FromYear")%> -
							<%if edu.DataValues("ToMonth") < 10 then%>0<%end if%>
							<%=edu.DataValues("ToMonth")%>/<%=edu.DataValues("ToYear")%></td>
						<td><strong><%=edu.DataValues("Place")%></strong>, <%=edu.DataValues("Title")%></td>
					</tr>
					<tr>
						<td></td>
						<td><%=edu.DataValues("Description")%></td>
					</tr>
					<%
					next
					%>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan="2"><hr></td>
		</tr>
		<tr>
			<td>
				<table cellpadding='0' cellspacing='0'>
					<tr>
						<td colspan="4"><strong>Utdannelse - Til godkjenning</strong></td>
					</tr>
					<%
					i = 1
					for each edu2 in allEdu2
					if edu2.DataValues("Type") <> "ORIGINAL" then%>
					<tr>
						<td><input type="checkbox" name="utdCheck<%=i%>" value="1"></td>
						<td><%if edu2.DataValues("FromMonth") < 10 then%>0<%end if%>
							<%=edu2.DataValues("FromMonth")%>/<%=edu2.DataValues("FromYear")%> -
							<%if edu2.DataValues("ToMonth") < 10 then%>0<%end if%>
							<%=edu2.DataValues("ToMonth")%>/<%=edu2.DataValues("ToYear")%></td>
						<td><strong><%=edu2.DataValues("Place")%></strong>, <%=edu2.DataValues("Title")%></td>
						<td><%=edu2.DataValues("Type")%></td>
					</tr>
					<tr>
						<td></td>
						<td></td>
						<td><%=edu2.DataValues("Description")%></td>
					</tr>
					<%end if
					i = i + 1
					next
					%>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan="2"><hr></td>
		</tr>
	</table>

	<p>

	<table cellpadding='0' cellspacing='0'>
		<%if vis_kvittering then%>
		<tr>
			<td colspan="2"><%=strKvittering%></td>
		</tr>
		<%end if%>
		<tr>
			<td><strong>Kommentar: </strong></td>
		</tr>
		<tr>
			<td colspan="2">
				<textarea name="kommentar" cols="60" rows="3"></textarea>
			</td>
		</tr>
		<tr>
			
				<input type="radio" name="godkjenn" value="1" <%if iGodkjenn = 1 then%>checked<%end if%>>&nbsp;Godkjenn avkryssede endringer<br>
				<input type="radio" name="godkjenn" value="2" <%if iGodkjenn = 2 then%>checked<%end if%>>&nbsp;Avvis avkryssede endringer
			</td>
			<td class="right">
				<input type="submit" name="ok" value="  Utfør  ">
				&nbsp;
				<input type="submit" name="tilbake" value=" Avbryt ">
			</td>
		</tr>
		<tr>
			<td><br></td>
		</tr>
		<tr>
			<td  colspan="2">
				<input type="submit" name="mail" value="Send kommentar til vikaren">
			</td>
		</tr>
		<tr>
			<td  colspan="2">
				<input type="submit" name="aapne" value="       Åpne CV for vikaren       ">
			</td>
		</tr>
	</table>

	<input type="hidden" name="rerun" value="1">
	<input type="hidden" name="data">
	<input type="hidden" name="education">
	<input type="hidden" name="objId">

</form>

</body>

<%
' sletter alle CV objekter...
set objCv2		= nothing
set objCv		= nothing
set objAdr		= nothing
objCons.ChangedCV.cleanup
objCons.CV.cleanup
objCons.cleanup
set objCons	= nothing
%>
</html>
