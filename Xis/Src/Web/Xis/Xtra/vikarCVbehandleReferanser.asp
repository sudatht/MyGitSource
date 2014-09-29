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
dim ref
dim ref2
dim allRef
dim allRef2
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
set ref		= Server.CreateObject("XtraWeb.Reference")
set ref2	= Server.CreateObject("XtraWeb.Reference")

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

objCv2.References.RefreshChanged

set allRef		= objCv.References
set allRef2		= objCv2.References

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
		for each ref2 in allRef2
			if Request.Form("refCheck" & i) = "1" then
				if Request.Form("godkjenn") = "1" then
					if not objCv2.References.ApproveChanged(ref2.DataValues("ReferenceId")) then
						Response.Write "Godkjennes ikke!<br>"
					end if
				elseif Request.Form("godkjenn") = "2" then
					if not objCv2.References.RejectChanges(ref2.DataValues("ReferenceId")) then
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
		Response.Redirect "vikarCVbehandleReferanser.asp?type=6&VikarID=" & strVikarID
		
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

<h3>CV godkjenning - <%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & strVikarID, Cons.DataValues("Fornavn") & " " & objCons.DataValues("Etternavn"), "Vis vikar " & objCons.DataValues("Fornavn") & " " & Cons.DataValues("Etternavn") )%></h3>

<table cellpadding='0' cellspacing='0'>
	<tr>
		<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandlePersonalia.asp?VikarID=<%=strVikarID%>">Personalia</a></strong></td>
		<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandleJobbonsker.asp?VikarID=<%=strVikarID%>">fagkompetanse</a></strong></td>
		<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandleKvalifikasjoner.asp?VikarID=<%=strVikarID%>">Produktkompetanse</a></strong></td>
		<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandleUtdannelse.asp?VikarID=<%=strVikarID%>">Utdannelse</a></strong></td>
		<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandlePraksis.asp?VikarID=<%=strVikarID%>">Praksis</a></strong></td>
		<td>&nbsp;<strong>|&nbsp;Referanser&nbsp;|</strong>&nbsp;</td>
	</tr>
</table>

<p>

<form name="behandle" action="vikarCVbehandleReferanser.asp?VikarID=<%=strVikarID%>" method="post">

	<table cellpadding='0' cellspacing='0'>
		<tr>
			<td>
				<table cellpadding='0' cellspacing='0'>
					<tr>
						<td colspan="4"><strong>Referanser - Godkjente<strong></td>
					</tr>
					<%
					for each ref in allRef
					%>
					<tr>
						<td><%=ref.DataValues("Name")%></td>
						<td><%=ref.DataValues("Firma")%></td>
					</tr>
					<tr>
						<td></td>
						<td><%=ref.DataValues("Title")%>, Tlf. <%=ref.DataValues("Tel")%></td>
					</tr>
					<tr>
						<td></td>
						<td><%=ref.DataValues("Comment")%></td>
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
						<td colspan="4"><strong>Referanser - Til godkjenning<strong></td>
					</tr>
					<%
					i = 1
					for each ref2 in allRef2
					if ref2.DataValues("Type") <> "ORIGINAL" then%>
					<tr>
						<td><input type="checkbox" name="refCheck<%=i%>" value="1"></td>
						<td><%=ref2.DataValues("Name")%></td>
						<td><%=ref2.DataValues("Firma")%></td>
					</tr>
					<tr>
						<td></td>
						<td></td>
						<td><%=ref2.DataValues("Title")%>, Tlf. <%=ref2.DataValues("Tel")%></td>
					</tr>
					<tr>
						<td></td>
						<td></td>
						<td><%=ref2.DataValues("Comment")%></td>
						<td><%=ref2.DataValues("Type")%></td>
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
	<input type="hidden" name="reference">
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
