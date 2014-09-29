<%@ Language=VBScript %>

<%option explicit%>

<%
'File purpose:		Lets the consultant contact approve or reject changes made
'					to a consultants CV on the external web.
'Created by:		Monica Johansen@electricfarm.no
'Changed by:		Fred.myklebust@electricfarm.no
'Changed Date:		22.02.2001
'Changes:			removed non used variables, added use jobgroup.GetAllApproved,
'					jobgroup.GetAllNonApproved methods & added comments.
'Tested date:

dim strNavn					'as string
dim strtekst				'as string
dim StrJobGroup				'as string
dim strJobWishes			'as string
dim lVikarID				'as long
dim strVikarID 				'as string
dim objCons					'xtraweb.consultant
dim iGodkjenn				'as integer
dim objCv					'xtraweb.cv
dim objCv2					'xtraweb.cv
dim ObjJobgroup				'xtraweb.jobgroup
dim strKomm					'as string
dim i
dim vis_kvittering
dim strKvittering 			'as string
dim RsJobWishes				'as adodb.recordset

if Request.QueryString("VikarID") <> "" then
	lVikarID = Request.QueryString("VikarID")
end if

strVikarID = Cstr(lVikarID)
iGodkjenn = "1"
strKvittering = ""

'FJM.Kan du vær så snill å erstatte disse med tilsvarende obj navn?
'Vær vennlig å gi andre navn enn navn2.

set objCons	= Server.CreateObject("XtraWeb.Consultant")
set ObjJobgroup	= Server.CreateObject("XtraWeb.JobGroup")


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

objCv2.JobGroups.RefreshChanged

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

		set ObjJobgroup = server.CreateObject("xtraweb.jobgroup")
		set RsJobWishes = ObjJobgroup.GetAllNonApproved(objCv.XtraConString, objCv.datavalues("cvid").value)

		with RsJobWishes
			while not .EOF
				if Request.Form("JobBox" & .fields("k_tittelid").value) = "1" then
					if Request.Form("godkjenn") = "1" then
						if not objCv.Jobgroups.ApproveChanged(.fields("K_TittelID")) then
							Response.Write "Godkjennes ikke!<br>"
						end if
					elseif Request.Form("godkjenn") = "2" then
						if not objCv.Jobgroups.RejectChanges(.fields("K_TittelID")) then
							Response.Write "ikke Avviset!<br>"
						end if
					else
						Response.Write "Feil!"
					end if
				end if
				.movenext
			wend
		end with
		set RsJobWishes 	= nothing
		set ObjJobgroup		= nothing
		set objCv		= nothing
		objCons.ChangedCV.cleanup
		objCons.CV.cleanup
		objCons.cleanup
		set objCons	= nothing
		Response.Redirect "vikarCVbehandleJobbonsker.asp?VikarID=" & strVikarID
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

<h1>CV godkjenning - <%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & strVikarID, objCons.DataValues("Fornavn") & " " & objCons.DataValues("Etternavn"), "Vis vikar " & objCons.DataValues("Fornavn") & " " & objCons.DataValues("Etternavn") )%></h1>
<table cellpadding='0' cellspacing='0'>
	<tr>
		<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandlePersonalia.asp?VikarID=<%=strVikarID%>">Personalia</a></td>
		<td>&nbsp;<strong>|&nbsp;fagkompetanse</strong></td>
		<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandleKvalifikasjoner.asp?VikarID=<%=strVikarID%>">Produktkompetanse</a></strong></td>
		<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandleUtdannelse.asp?VikarID=<%=strVikarID%>">Utdannelse</a></strong></td>
		<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandlePraksis.asp?VikarID=<%=strVikarID%>">Praksis</a></strong></td>
		<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandleReferanser.asp?VikarID=<%=strVikarID%>">Referanser</a>&nbsp;|</strong></td>
	</tr>
</table>

<p>

<form name="behandle" action="vikarCVbehandleJobbonsker.asp?VikarID=<%=strVikarID%>" method="post">

	<table cellpadding='0' cellspacing='0'>
		<tr>
			<td colspan="2"><strong>fagkompetanse - Godkjente</strong></td>
		</tr>
		<tr>
			<td>
				<table cellpadding='0' cellspacing='0'>
					<%
					set ObjJobgroup = server.CreateObject("xtraweb.jobgroup")
					set RsJobWishes = ObjJobgroup.GetAllApproved(objCv.XtraConString, objCv.datavalues("cvid").value)
					StrJobGroup = ""
					with RsJobWishes
					while not .EOF
						if .fields("navn").value <> StrJobGroup then
							if StrJobGroup <> "" then
								Response.Write left(strJobWishes, len(strJobWishes)-2)
								Response.Write "</td></tr></table></td></tr>"
								strJobWishes = ""
							end if
							Response.Write "<tr><td height=""10""></td></tr>"
							StrJobGroup = .fields("navn").value
							Response.Write "<tr><td> <i>" & StrJobGroup & "</i></td></tr>"
							Response.Write "<tr><td> <table cellpadding=""0"" cellspacing=""0""><tr>"
							Response.Write "<td> "
						end if
						strJobWishes = strJobWishes & .fields("Ktittel").value & ", "
						RsJobWishes.movenext
					wend
					end with
					if len(strJobWishes)> 0 then
						Response.Write left(strJobWishes, len(strJobWishes)-2)
					end if
					Response.Write "</td></tr></table></td></tr>"
					set ObjJobgroup	= nothing
					set RsJobWishes = nothing
					%>
				</table>
			</td>
		<tr>
		<tr>
			<td colspan="2"><hr></td>
		</tr>
		<tr>
			<td colspan="2"><strong>fagkompetanse - Til godkjenning</strong></td>
		</tr>
		<tr>
			<td>
				<table cellpadding='0' cellspacing='0'>
					<%
					set ObjJobgroup = server.CreateObject("xtraweb.jobgroup")
					set RsJobWishes = ObjJobgroup.GetAllNonApproved(objCv.XtraConString, objCv.datavalues("cvid").value)
					StrJobGroup = ""
					with RsJobWishes
					while not .EOF
						if .fields("navn").value <> StrJobGroup then
							if StrJobGroup <> "" then
								Response.Write "</td></tr></table></td></tr>"
							end if
							Response.Write "<tr><td height=""10""></td></tr>"
							StrJobGroup = .fields("navn").value
							Response.Write "<tr><td> <i>" & StrJobGroup & "</i></td></tr>"
							Response.Write "<tr><td> <table cellpadding=""0"" cellspacing=""0""><tr>"
						end if
						Response.Write "<td> <input type=""checkbox"" name=""JobBox" & .fields("k_tittelid").value &""" value=""1""></td>"
						Response.Write "<td>" & .fields("Ktittel").value & "</td>"
						Response.Write "<td> "& .fields("type").value & "</td></tr>"
						RsJobWishes.movenext
					wend
					end with
					Response.Write "</table></td></tr>"
					set ObjJobgroup	= nothing
					set RsJobWishes = nothing
					%>
				</table>
			</td>
		<tr>
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
objCons.ChangedCV.cleanup
objCons.CV.cleanup
objCons.cleanup
set objCons	= nothing
%>
</html>
