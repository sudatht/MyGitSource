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

set objCons	= Server.CreateObject("XtraWeb.Consultant")
set objAdr	= Server.CreateObject("XtraWeb.Address")

objCons.XtraConString = Application("Xtra_intern_ConnectionString")
objCons.XtraDataShapeConString = Application("ConXtraShape")
objCons.GetConsultant(lVikarID)

set objAdr = objCons.Addresses(1)
set objCv	= objCons.CV
objCv.XtraConString = Application("Xtra_intern_ConnectionString")
objCv.XtraDataShapeConString = Application("ConXtraShape")
objCv.Refresh

set objCv2	= objCons.ChangedCV
objCv2.XtraConString = Application("Xtra_intern_ConnectionString")
objCv2.XtraDataShapeConString = Application("ConXtraShape")
objCv2.Refreshchanged

vis_kvittering = false

if objCv.Datavalues.Count > 1 then
	if not isNull(objCv.DataValues("Filename")) then
		cvfil = objCv.DataValues("Filename")
	else
		cvfil = ""
	end if
else
	cvfil = ""
end if

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

		set objCv2		= nothing
		set objCv		= nothing
		set objAdr		= nothing
		objCons.ChangedCV.cleanup
		objCons.CV.cleanup
		objCons.cleanup
		set objCons	= nothing
		Response.Redirect "vikarCVbehandlePersonalia.asp?type=1&VikarID=" & strVikarID
		
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

<h3>CV godkjenning - <a href="vikarvis.asp?VikarID=<%=strVikarID%>"><%=objCons.DataValues("Fornavn")%>&nbsp;<%=objCons.DataValues("Etternavn")%></a></h3>
<table cellpadding='0' cellspacing='0'>
	<tr>
		<td>&nbsp;<strong>|&nbsp;Personalia</strong></td>
		<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandleJobbonsker.asp?VikarID=<%=strVikarID%>"><strong>fagkompetanse</a></strong></td>
		<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandleKvalifikasjoner.asp?VikarID=<%=strVikarID%>"><strong>Produktkompetanse</a></strong></td>
		<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandleUtdannelse.asp?VikarID=<%=strVikarID%>">Utdannelse</a></strong></td>
		<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandlePraksis.asp?VikarID=<%=strVikarID%>">Praksis</a></strong></td>
		<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandleReferanser.asp?VikarID=<%=strVikarID%>">Referanser</a>&nbsp;|</strong></td>
	</tr>
</table>

<p>

<form name="behandle" action="vikarCVbehandlePersonalia.asp?VikarID=<%=strVikarID%>" method="post">

	<table cellpadding='0' cellspacing='0'>
		<tr>
			<td colspan="2"><strong>Personalia</strong></td>
		</tr>
		<tr>
			<td>Navn: </td>
			<td><%=objCons.DataValues("Fornavn")%>&nbsp;<%=objCons.DataValues("Etternavn")%></td>
		</tr>
		<tr>
			<td>Adresse:</td>
			<td><%=objAdr.DataValues("Adresse")%></td>
		</tr>
		<tr>
			<td></td>
			<td><%=objAdr.DataValues("Postnr")%>&nbsp;<%=objAdr.DataValues("Poststed")%></td>
		</tr>
		<tr>
			<td>Telefon:</td>
			<td><%=objCons.DataValues("Telefon")%></td>
		</tr>
		<tr>
			<td>Mobil:</td>
			<td><%=objCons.DataValues("MobilTlf")%></td>
		</tr>
		<tr>
			<td>E-post:</td>
			<td><%=objCons.DataValues("EPost")%></td>
		</tr>
		<tr>
			<td>Fødselsdato:</td>
			<td><%=objCons.DataValues("Foedselsdato")%></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td colspan="2"><input type="checkbox" name="forerkort" <%if objCons.DataValues("foererkort") = 1 then%>checked<%end if%>>Har førerkort</td>
		</tr>
		<tr>
			<td colspan="2"><input type="checkbox" name="bildisp" <%if objCons.DataValues("bil") = 1 then%>checked<%end if%>>Disponerer bil</td>
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
