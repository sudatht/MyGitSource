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
dim iGodkjenn
dim objCv
dim objCv2
dim exp
dim exp2
dim allExp
dim allExp2
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
set exp		= Server.CreateObject("XtraWeb.Experience")
set exp2	= Server.CreateObject("XtraWeb.Experience")

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

objCv2.Experiences.RefreshChanged

set allExp		= objCv.Experiences
set allExp2		= objCv2.Experiences

vis_kvittering = false

if Request.Form("rerun") = "1" then

	strKomm = Request.Form("kommentar")

	' Ferdig
	if Request.Form("tilbake") = " Avbryt " then
		set objCv2		= nothing
		set objCv		= nothing
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
		for each exp2 in allExp2
			if Request.Form("praCheck" & i) = "1" then
				if Request.Form("godkjenn") = "1" then
					if not objCv2.Experiences.ApproveChanged(exp2.DataValues("DataId")) then
						Response.Write "Godkjennes ikke!<br>"
					end if
				elseif Request.Form("godkjenn") = "2" then
					if not objCv2.Experiences.RejectChanges(exp2.DataValues("DataId")) then
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

		objCons.ChangedCV.cleanup
		objCons.CV.cleanup
		objCons.cleanup
		set objCons	= nothing
		Response.Redirect "vikarCVbehandlePraksis.asp?type=5&VikarID=" & strVikarID
		
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

<h3>CV godkjenning - <%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & strVikarID, objCons.DataValues("Fornavn") & " " & objCons.DataValues("Etternavn"), "Vis vikar " & objCons.DataValues("Fornavn") & " " & objCons.DataValues("Etternavn") )%></h3>

<table cellpadding='0' cellspacing='0'>
	<tr>
		<td>&nbsp;<strong>|</strong>&nbsp;<a href="vikarCVbehandlePersonalia.asp?VikarID=<%=strVikarID%>"><strong>Personalia</strong></a></td>
		<td>&nbsp;<strong>|</strong>&nbsp;<a href="vikarCVbehandleJobbonsker.asp?VikarID=<%=strVikarID%>"><strong>fagkompetanse</strong></a></td>
		<td>&nbsp;<strong>|</strong>&nbsp;<a href="vikarCVbehandleKvalifikasjoner.asp?VikarID=<%=strVikarID%>"><strong>Produktkompetanse</strong></a></td>
		<td>&nbsp;<strong>|</strong>&nbsp;<a href="vikarCVbehandleUtdannelse.asp?VikarID=<%=strVikarID%>"><strong>Utdannelse</strong></a></td>
		<td>&nbsp;<strong>|</strong>&nbsp;<strong>Praksis</strong></td>
		<td>&nbsp;<strong>|</strong>&nbsp;<a href="vikarCVbehandleReferanser.asp?VikarID=<%=strVikarID%>"><strong>Referanser</strong></a>&nbsp;<strong>|</strong>&nbsp;</td>
	</tr>
</table>

<p>

<form name="behandle" action="vikarCVbehandlePraksis.asp?VikarID=<%=strVikarID%>" method="post">

	<table cellpadding='0' cellspacing='0'>
		<tr>
			<td>
				<table cellpadding='0' cellspacing='0'>
					<tr>
						<td colspan="4"><strong>Praksis - Godkjente</strong></td>
					</tr>
					<%
					for each exp in allExp
					%>
					<tr>
						<td><%if exp.DataValues("FromMonth") < 10 then%>0<%end if%>
							<%=exp.DataValues("FromMonth")%>/<%=exp.DataValues("FromYear")%> -
							<%if exp.DataValues("ToMonth") < 10 then%>0<%end if%>
							<%=exp.DataValues("ToMonth")%>/<%=exp.DataValues("ToYear")%></td>
						<td><strong><%=exp.DataValues("Place")%></strong>, <%=exp.DataValues("Title")%></td>
					</tr>
					<tr>
						<td></td>
						<td><%=exp.DataValues("Description")%></td>
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
						<td colspan="4"><strong>Praksis - Til godkjenning</strong></td>
					</tr>
					<%
					i = 1
					for each exp2 in allExp2
					if exp2.DataValues("Type") <> "ORIGINAL" then%>
					<tr>
						<td><input type="checkbox" name="praCheck<%=i%>" value="1"></td>
						<td><%if exp2.DataValues("FromMonth") < 10 then%>0<%end if%>
							<%=exp2.DataValues("FromMonth")%>/<%=exp2.DataValues("FromYear")%> -
							<%if exp2.DataValues("ToMonth") < 10 then%>0<%end if%>
							<%=exp2.DataValues("ToMonth")%>/<%=exp2.DataValues("ToYear")%></td>
						<td><strong><%=exp2.DataValues("Place")%></strong>, <%=exp2.DataValues("Title")%></td>
						<td><%=exp2.DataValues("Type")%></td>
					</tr>
					<tr>
						<td></td>
						<td></td>
						<td><%=exp2.DataValues("Description")%></td>
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
	<table cellpadding='0' cellspacing='0'>


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
			<td><br></td>
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
	<input type="hidden" name="experience">
	<input type="hidden" name="objId">

</form>

</body>

<%
' sletter alle CV objekter...
set objCv2		= nothing
set objCv		= nothing
objCons.ChangedCV.cleanup
objCons.CV.cleanup
objCons.cleanup
set objCons	= nothing
%>
</html>
