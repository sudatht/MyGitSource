<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<!-- #include file = "cuteeditor_files/include_CuteEditor.asp" --> 
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if	
	
	dim ref
	dim blnLocked
	dim fulltime
	dim beskrivelse
	dim ref_id
	dim feilmelding
	dim dette_aar
	dim strAction

	'Global variables - used in tool menu
	dim blnShowHotList
	dim lngVikarID

	// Text Editor Code - CuteEditor	
	Dim editor	
	Set editor = New CuteEditor
	editor.ID = "Editor1"
	editor.AutoConfigure = "Simple"	

	'Initialize variables
	blnShowHotList = false

	set cons	= Server.CreateObject("XtraWeb.Consultant")

	lngVikarID	= CStr(Request("VikarID"))
	cons.XtraConString = Application("XtraWebConnection")
	cons.GetConsultant(lngVikarID)

	dette_aar = Year(date)

	set cv	= cons.CV
	cv.XtraConString = Application("Xtra_intern_ConnectionString")
	cv.XtraDataShapeConString = Application("ConXtraShape")
	cv.Refresh

	if cons.CV.DataValues.Count = 0 then
		cons.CV.Save
	end if

	blnLocked = cv.islocked

	set allRef = cv.References

	set ref	= Server.CreateObject("XtraWeb.Reference")
	set ref.Owner = cv

	ref_id = 0

	if Request.QueryString("refid").Count > 0 then
		ref_id = Request.QueryString("refid")
		set ref = cv.References("ID" & ref_id)
		refKomm = ref.DataValues("Comment")
	end if

	if Request.Form("rerun") = "1" then

		strAction  = lcase(trim(request("pbnDataAction")))

		' Utdannelse..
		if (strAction = "lagre") then
			'if cv.islocked then
			'	cv.UnlockCV
			'	cv.deleteChanged
			'	blnLocked = false
			'end if

			ref_id = Request.Form("objId")
			if ref_id = "" then
				ref_id = 0
			end if
			if ref_id > 0 then
				set ref = cv.References("ID" & ref_id)
			end if
			refNavn		= CStr(Request.Form("refNavn"))
			refFirma	= CStr(Request.Form("refFirma"))
			refTittel	= CStr(Request.Form("refTittel"))
			refTelefon	= CStr(Request.Form("refTelefon"))
			refKomm		= CStr(Request.Form("Editor1"))

			if refTelefon <> "" then
				if not IsNumeric(refTelefon) then
					feilmelding = feilmelding & "Ugyldig telefonnummer! Bruk kun heltall (og evt. mellomrom)!<br>"
				end if
			end if

			if refNavn <> "" and refFirma <> "" and refTittel <> "" and refTelefon <> "" and feilmelding = "" then
				ref.DataValues("Name")		= refNavn
				ref.DataValues("Firma")		= refFirma
				ref.DataValues("Title")		= refTittel
				ref.DataValues("Tel")		= refTelefon
				ref.DataValues("Comment")	= refKomm
				ref.Save
				set cv		= nothing
				cons.CV.cleanup
				cons.cleanup
				set cons	= nothing
				'Response.End
				Response.Redirect("vikarCVnyReferanser.asp?type=6&VikarID=" & lngVikarID)
			else
				feilmelding = feilmelding & "Alle opplysninger bortsett fra Kommentarer er obligatoriske!<br>"
			end if
	end if
	if (strAction = "slett") then
		'if cv.islocked then
		'	cv.UnlockCV
		'	cv.deleteChanged
		'	blnLocked = false
		'end if

		ref_id = clng(Request.Form("objId"))
		if ref_id <> 0 then
			set ref = cv.References("ID" & ref_id)
			if not ref.Delete then
				Response.Write "Ikke slettet!<br>"
			end if
			set cv		= nothing
			cons.CV.cleanup
			cons.cleanup
			set cons	= nothing
			Response.Redirect("vikarCVnyReferanser.asp?type=6&VikarID=" & lngVikarID)
		end if
	end if
end if
refKomm = "<link type='text/css' rel='stylesheet' href='http://" & Application("HTTPadress") & "/xtra/css/CV.css' title='default style'>" & refKomm
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
	<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print"><strong></strong>

	<script language="javaScript" type="text/javascript">
		function slett_data(id) 
		{
			document.all.pbnDataAction.value='slett';
			document.CVForm.objId.value = id;
			document.CVForm.submit();
		}
		function lagre_data(id) {
			document.CVForm.objId.value = id;
			document.CVForm.submit();
		}
	</script>
	<script type="text/javascript" src="/xtra/js/CVFunctions.js"></script>
	<script type="text/javascript" src="/xtra/js/contentMenu.js"></script>
	<script type="text/javascript" src="/xtra/js/fontSizer.js"></script>
</head>
<body style="PADDING-RIGHT: 0px; PADDING-LEFT: 0px;">
	<div class="pageContainer" id="pageContainer">
		<div class="contentHead1">
			<H1 style="font-size:160%; font-weight:bold";>CV redigering - <%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & lngVikarID, Cons.DataValues("Fornavn") & " " & Cons.DataValues("Etternavn"), "Vis vikar " & Cons.DataValues("Fornavn") & " " & Cons.DataValues("Etternavn") )%></h1>
			<div class="contentMenu">
				<table cellpadding="0" cellspacing="0" width="96%">
					<tr>
						<td><!--#include file="vikarCVnyToolbar.asp"--></td>
						<td class="right">
						<!--#include file="Includes/contentToolsMenu.asp"-->
						</td>
					</tr>
				</table>
			</div>
		</div>
		<div class="contentMenu2">
			<span class="menu2" id="1" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyPersonalia.asp?VikarID=<%=lngVikarID%>">Personalia</a></span>
			<span class="menu2" id="2" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyJobbonsker.asp?VikarID=<%=lngVikarID%>">Fagkompetanse</a></span>
			<span class="menu2" id="3" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyKvalifikasjoner.asp?VikarID=<%=lngVikarID%>">Produktkompetanse</a></span>
			<!--
			<span class="menu2" id="4" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyNokkelKvalifikasjoner.asp?VikarID=<%=lngVikarID%>">Kandidatpresentasjon</a></span>
			-->
			<span class="menu2" id="5" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyUtdannelse.asp?VikarID=<%=lngVikarID%>">Utdannelse</a></span>
			<span class="menu2" id="6" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyCourses.asp?VikarID=<%=lngVikarID%>">Kurs</a></span>
			<span class="menu2" id="7" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyPraksis.asp?VikarID=<%=lngVikarID%>">Yrkeserfaring</a></span>
			<span class="menu2" id="8" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyAndreOppl.asp?VikarID=<%=lngVikarID%>">Kjernekompetanse</a></span>
			<span class="menu2 active" id="9"><strong>Referanser</strong></span>
			<span class="menu2" id="10" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVvis.asp?VikarID=<%=lngVikarID%>">&nbsp;<img src="/xtra/images/icon_new2.gif" width="14" height="14" alt="" border="0" align="absmiddle">&nbsp;Generere CV</a></span>
		</div>
		<div class="content">
			<form name="CVForm" action="vikarCVnyReferanser.asp?VikarID=<%=lngVikarID%>" method="post">
				<input type="hidden" id="rerun" name="rerun" value="1">
				<input Type="hidden" Name="pbnDataAction" Value="lagre">
				<input type="hidden" name="objId" value="<%=ref_id%>">
				<table width="96%" class="layout" cellpadding="0" cellspacing="0">
					<col width="50%">
					<col width="50%">
				<tr>
					<td>
						<table width="100%" cellpadding='0' cellspacing='0'>
							<tr>
								<td>Navn:</td>
								<td><input type="text" name="refNavn" size="40" maxlength="50"
									<%if ref_id <> 0 then%> value="<%=ref.DataValues("Name")%>"
									<%else%> value="<%=refNavn%>" <%end if%> class="mandatory"></td>
							</tr>
							<tr>
								<td>Kontakt:</td>
								<td><input type="text" name="refFirma" size="40" maxlength="50"
									<%if ref_id <> 0 then%> value="<%=ref.DataValues("Firma")%>"
									<%else%> value="<%=refFirma%>" <%end if%> class="mandatory"></td>
							</tr>
							<tr>
								<td>Tittel:</td>
								<td><input type="text" name="refTittel" size="40" maxlength="50"
									<%if ref_id <> 0 then%> value="<%=ref.DataValues("title")%>"
									<%else%> value="<%=refTittel%>" <%end if%> class="mandatory"></td>
							</tr>
							<tr>
								<td>Telefon:</td>
								<td><input type="text" name="refTelefon" size="40" maxlength="50"
									<%if ref_id <> 0 then%> value="<%=ref.DataValues("Tel")%>"
									<%else%> value="<%=refTelefon%>" <%end if%> class="mandatory"></td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
							<tr>							
								<td>Kommentarer:</td>
								<td>
									<%
								    editor.Text = refKomm
									editor.Draw()
								%>
								</td>
							</tr>								
						</table>
					</td>
					<td>
						<%if feilmelding <> "" then%>
							<p class="warning"><%=feilmelding%></p>
						<%end if%>&nbsp;
					</td>
				</tr>
				</table>
				
				<br/><br/>
				<span class="menuInside" style="margin-left:70px;" title="Lagre informasjonen"><a href="#" onClick="javascript:lagre_data(<%=ref_id%>);"><img src="images/icon_save.gif" width="18" height="15" alt="" border="0" align="absmiddle">Lagre</a></span>
				<div class="listing">
					<table cellpadding='0' cellspacing='1'>
						<tr>
							<th colspan="7">Registrerte referanser</th>
						</tr>
						<tr>
							<th>Navn</th>
							<th>Kontakt</th>
							<th>Tittel</th>
							<th>Telefon</th>
							<th>Kommentarer</th>
							<th class="center">Endre</th>
							<th class="center">Slette</th>
						</tr>
						<%
						for i=1 to allRef.Count
							set ref = allRef.Item(i)
							id = ref.DataValues("ReferenceID")%>
						<tr>
							<td><%=ref.DataValues("Name")%></td>
							<td><%=ref.DataValues("Firma")%></td>
							<td><%=ref.DataValues("Title")%></td>
							<td><%=ref.DataValues("Tel")%></td>
							<td><%=ref.DataValues("Comment")%>&nbsp;</td>
							<td class="center"><a href="vikarCVnyReferanser.asp?type=6&VikarID=<%=lngVikarID%>&refid=<%=id%>"><img src="/xtra/images/icon_changeInfo.gif" width="18" height="15" alt="Endre denne oppføringen" border="0"></a></td>
							<td class="center"><a href="javascript:slett_data(<%=id%>)"><img src="/xtra/images/icon_delete.gif" width="14" height="14" alt="Slette denne oppføringen" border="0"></a></td>
						</tr>
						<%next%>
					</table>
				</div>
			</form>
		</div>
<%
' sletter alle CV objekter...
set cv		= nothing
cons.CV.cleanup
cons.cleanup
set cons	= nothing
%>
</div>
</body>
</html>

