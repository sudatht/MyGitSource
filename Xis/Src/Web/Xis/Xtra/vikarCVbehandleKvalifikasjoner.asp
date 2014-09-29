<%@ Language=VBScript %>
<%option explicit%>
<%
'File purpose:		Lets the consultant contact approve or reject changes made
'					to a consultants CV on the external web.
'Created by:		Monica Johansen@electricfarm.no
'Changed by:		Fred.myklebust@electricfarm.no
'Changed Date:		13.02.2001
'Changes:			removed non used variables, added use productgroup.GetAllApproved,
'					productgroup.GetAllNonApproved methods & added comments.
'Tested date:

	dim strNavn				'as string
	dim strTekst			'as string
	dim strProducts			'as string
	dim lVikarID			'as long
	dim strVikarID			'as string
	dim objCons				'as xtraweb.consultant
	dim iGodkjenn			'as integer
	dim objCv				'as xtraweb.cv
	dim ObjProdgroup		'as xtraweb.productgroup
	dim strKomm				'as string
	dim vis_kvittering		'as boolean
	dim strKvittering		'as string
	dim StrProductGroup		'as string
	dim RsQualifications	'as adodb.recordset

	if Request.QueryString("VikarID") <> "" then
		lVikarID = Request.QueryString("VikarID")
	end if

	strVikarID = Cstr(lVikarID)
	iGodkjenn = "1"
	strKvittering = ""

	set objCons	= Server.CreateObject("XtraWeb.Consultant")
	set ObjProdgroup	= Server.CreateObject("XtraWeb.ProductGroup")

	objCons.XtraConString = Application("Xtra_intern_ConnectionString")
	objCons.XtraDataShapeConString = Application("ConXtraShape")
	objCons.GetConsultant(lVikarID)

	set objCv	= objCons.ChangedCV
	objCv.XtraConString = Application("Xtra_intern_ConnectionString")
	objCv.XtraDataShapeConString = Application("ConXtraShape")
	objCv.Refreshchanged

	objCv.ProductGroups.RefreshChanged

	if Request.Form("rerun") = "1" then

		strKomm = Request.Form("kommentar")

		' Ferdig
		if Request.Form("tilbake") = " Avbryt " then
			set objCv		= nothing
			objCons.ChangedCV.cleanup
			objCons.CV.cleanup
			objCons.cleanup
			set objCons	= nothing
			Response.Redirect("vikarCVvis.asp?VikarID=" & strVikarId)
		end if

		if Request.Form("aapne") = "       Åpne CV for vikaren       " then
			objCv.UnlockCV
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
			set ObjProdgroup = server.CreateObject("xtraweb.productgroup")
			set RsQualifications = ObjProdgroup.GetAllNonApproved(objCv.XtraConString, objCv.datavalues("cvid").value)
			with RsQualifications
				while not .EOF
					if Request.Form("prodBox" & .fields("k_tittelid").value) = "1" then
						if Request.Form("godkjenn") = "1" then
							if not objCv.ProductGroups.ApproveChanged(.fields("K_TittelID")) then
								Response.Write "Godkjennes ikke!<br>"
							end if
						elseif Request.Form("godkjenn") = "2" then
							if not objCv.ProductGroups.RejectChanges(.fields("K_TittelID")) then
								Response.Write "ikke Avviset!<br>"
							end if
						else
							Response.Write "Feil!"
						end if
					end if
					RsQualifications.movenext
				wend
			end with
			set RsQualifications = nothing
			set ObjProdgroup		= nothing
			set objCv		= nothing
			objCons.ChangedCV.cleanup
			objCons.CV.cleanup
			objCons.cleanup
			set objCons	= nothing
			Response.Redirect "vikarCVbehandleKvalifikasjoner.asp?type=3&VikarID=" & strVikarID

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
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	</head>
	<body>
	<div class="pageContainer" id="pageContainer">

	<h3>CV godkjenning - <%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & strVikarID, objCons.DataValues("Fornavn") & " " & objCons.DataValues("Etternavn"), "Vis vikar " & objCons.DataValues("Fornavn") & " " & objCons.DataValues("Etternavn") )%></h3>

	<table cellpadding='0' cellspacing='0'>
		<tr>
			<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandlePersonalia.asp?VikarID=<%=strVikarID%>">Personalia</a></strong></td>
			<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandleJobbonsker.asp?VikarID=<%=strVikarID%>">fagkompetanse</a></strong></td>
			<td>&nbsp;<strong>|&nbsp;Produktkompetanse</strong></td>
			<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandleUtdannelse.asp?VikarID=<%=strVikarID%>">Utdannelse</a></strong></td>
			<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandlePraksis.asp?VikarID=<%=strVikarID%>">Praksis</a></strong></td>
			<td>&nbsp;<strong>|&nbsp;<a href="vikarCVbehandleReferanser.asp?VikarID=<%=strVikarID%>"><strong>Referanser</a>&nbsp;|</strong></td>
		</tr>
	</table>

	<p>

	<form name="behandle" action="vikarCVbehandleKvalifikasjoner.asp?VikarID=<%=strVikarID%>" method="post">
		<table cellpadding='0' cellspacing='0'>
			<tr>
				<td colspan="2"><strong>Produktkompetanse - Godkjente</strong></td>
			</tr>
			<tr>
				<td colspan="2">
					<table cellpadding='0' cellspacing='0'>
						<%
						set ObjProdgroup = server.CreateObject("xtraweb.productgroup")
						set RsQualifications = ObjProdgroup.GetAllApproved(objCv.XtraConString, objCv.datavalues("cvid").value)
						StrProductGroup = ""
						with RsQualifications
						while not .EOF
							if .fields("Produktomrade").value <> StrProductGroup then
								if StrProductGroup <> "" then
									Response.Write left(strProducts, len(strProducts)-2)
									Response.Write "</td></tr></table></td></tr>"
									strProducts = ""
								end if
								Response.Write "<tr><td height=""10""></td></tr>"
								StrProductGroup = .fields("Produktomrade").value
								Response.Write "<tr><td> <i>" & StrProductGroup & "</i></td></tr>"
								Response.Write "<tr><td> <table cellpadding=""0"" cellspacing=""0""><tr>"
								Response.Write "<td> "
							end if
							strProducts = strProducts & .fields("Ktittel").value & ", "
							RsQualifications.movenext
						wend
						end with
						if len(strProducts)> 0 then
							Response.Write left(strProducts, len(strProducts)-2)
						end if
						Response.Write "</td></tr></table></td></tr>"
						set ObjProdgroup		= nothing
						%>
					</table>
				</td>
			</tr>
			<tr>
				<td colspan="2"><hr></td>
			</tr>
			<tr>
				<td colspan="2"><strong>Produktkompetanse - Til godkjenning</strong></td>
			</tr>
			<tr>
				<td>
					<table cellpadding='0' cellspacing='0'>
						<%
						set ObjProdgroup = server.CreateObject("xtraweb.productgroup")
						set RsQualifications = ObjProdgroup.GetAllNonApproved(objCv.XtraConString, objCv.datavalues("cvid").value)
						StrProductGroup = ""
						while not RsQualifications.EOF
							if RsQualifications.fields("Produktomrade").value <> StrProductGroup then
								if StrProductGroup <> "" then
									Response.Write "</table></td></tr>"
								end if
								Response.Write "<tr><td height=""10""></td></tr>"
								StrProductGroup = RsQualifications.fields("Produktomrade").value
								Response.Write "<tr><td valign=""top"" colspan=""2""><i>" & StrProductGroup & "</i></td></tr>"
								Response.Write "<tr><td> <table cellpadding=""0"" cellspacing=""0""><tr>"
							end if
							Response.Write "<td> <input type=""checkbox"" name=""prodBox" & RsQualifications.fields("k_tittelid").value &""" value=""1""></td>"
							Response.Write "<td>" & RsQualifications.fields("Ktittel").value & "</td>"
							Response.Write "<td> "& RsQualifications.fields("type").value & "</td></tr>"
							RsQualifications.movenext
						wend
						Response.Write "</table></td></tr>"
						set ObjProdgroup = nothing
						set RsQualifications = nothing
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
		<input type="hidden" name="objId">

	</form>

	</body>

	<%
	' sletter alle CV objekter...
	set objCv		= nothing
	objCons.ChangedCV.cleanup
	objCons.CV.cleanup
	objCons.cleanup
	set objCons	= nothing
	%>
	</html>
