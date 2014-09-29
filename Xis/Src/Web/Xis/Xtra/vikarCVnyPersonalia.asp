<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	
	dim blnLocked
	dim feilmelding
	dim harForerkort
	dim disponererBil
	dim cons
	dim adr
	dim	strOppsigelsestid
	dim	strWorkType
	dim	strCountry

	'Global variables - used in tool menu
	dim blnShowHotList
	dim lngVikarID

	'Initialize variables
	blnShowHotList = false

	set cons	= Server.CreateObject("XtraWeb.Consultant")
	set adr		= Server.CreateObject("XtraWeb.Address")

	lngVikarID	= CStr(Request("VikarID"))
	cons.XtraConString = Application("XtraWebConnection")
	cons.GetConsultant(lngVikarID)
	
	lCountryID = 1  'default value	
	if cons.DataValues("Country") > 0 then
		lCountryID		= cons.DataValues("Country")
	end if	
	
	' Open database connection
   	Set objCon = GetConnection(GetConnectionstring(XIS, ""))
	set rsCountry = GetFirehoseRS("SELECT CountryID,PrintableName FROM COUNTRY WHERE CountryID = " & lCountryID, objCon)	
	strCountryName	= rsCountry("PrintableName").Value	
	
	strOppsigelsestid		= cons.DataValues("Oppsigelsestid")
        strWorkType		= cons.DataValues("WorkType")
	
	select case strOppsigelsestid
		case ""
			strOppsigelsestid = "selected"
		case "0"
			strOppsigelsestid = "Ingen"
		case "1"
			strOppsigelsestid = "14 dager"
		case "2"
			strOppsigelsestid = "1 m&aring;ned"
		case "3"
			strOppsigelsestid = "2 m&aring;neder"
		case "4"
			strOppsigelsestid = "3 m&aring;neder"
		case "5"
			strOppsigelsestid = "Over 3 m&aring;neder"
		end select

		select case strWorkType
		case "0"
			strWorkType = "Ikke angitt"
		case "1"
			strWorkType = "Kun fulltid"
		case "2"
			strWorkType = "Kun deltid"
		case "3"
			strWorkType = "Fulltid og deltid"
		end select


	if cons.DataValues("foererkort") = 1 then
		harForerkort = 1
	else 
		harForerkort = 0
	end if
	
	if cons.DataValues("hasCar") = 1 then
		disponererBil = 1
	else 
		disponererBil = 0
	end if

	dette_aar = Year(date)

	set adr = cons.Addresses(1)
	set cv	= cons.CV
	cv.XtraConString = Application("Xtra_intern_ConnectionString")
	cv.XtraDataShapeConString = Application("ConXtraShape")
	cv.Refresh

	if cons.CV.DataValues.Count = 0 then
		cons.CV.Save
	end if

	blnLocked = cv.islocked

	if Request.Form("rerun") = "1" then
		' Personalia
		if Request.Form("pbnDataAction") = "lagre" then

			if Request.Form("forerkort") = "1" then
				harForerkort = 1
			else
				harForerkort = 0
			end if
			if Request.Form("bildisponering") = "1" then
				disponererBil = 1
			else
				disponererBil = 0
			end if
			cons.DataValues("foererkort") = harForerkort
			cons.DataValues("hasCar") = disponererBil
			cons.Save
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
	<script language="javaScript" type="text/javascript">
		function lagre_data() 
		{
			document.all.pbnDataAction.value='lagre';
			document.CVForm.submit();
		}
	</script>
	<script type="text/javascript" src="/xtra/js/CVFunctions.js"></script>
	<script type="text/javascript" src="/xtra/js/contentMenu.js"></script>
</head>
<body>
	<div class="pageContainer" id="pageContainer">
		<div class="contentHead1">
			<h1>CV redigering - <%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & lngVikarID, cons.DataValues("Fornavn") & " " & cons.DataValues("Etternavn"), "Vis vikar " & cons.DataValues("Fornavn") & " " & cons.DataValues("Etternavn") )%></h1>
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
			<span class="menu2 active" id="1"><strong>Personalia</strong></span>
			<span class="menu2" id="2" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyJobbonsker.asp?VikarID=<%=lngVikarID%>">Fagkompetanse</a></span>
			<span class="menu2" id="3" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyKvalifikasjoner.asp?VikarID=<%=lngVikarID%>">Produktkompetanse</a></span>
			<!--
			<span class="menu2" id="4" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyNokkelKvalifikasjoner.asp?VikarID=<%=lngVikarID%>">Kandidatpresentasjon</a></span>
			-->
			<span class="menu2" id="5" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyUtdannelse.asp?VikarID=<%=lngVikarID%>">Utdannelse</a></span>
			<span class="menu2" id="6" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyCourses.asp?VikarID=<%=lngVikarID%>">Kurs</a></span>
			<span class="menu2" id="7" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyPraksis.asp?VikarID=<%=lngVikarID%>">Yrkeserfaring</a></span>
			<span class="menu2" id="8" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyAndreOppl.asp?VikarID=<%=lngVikarID%>">Kjernekompetanse</a></span>
			<span class="menu2" id="9" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyReferanser.asp?VikarID=<%=lngVikarID%>">Referanser</a></span>
			<span class="menu2" id="10" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVvis.asp?VikarID=<%=lngVikarID%>">&nbsp;<img src="/xtra/images/icon_new2.gif" width="14" height="14" alt="" border="0" align="absmiddle">&nbsp;Generere CV</a></span>
		</div>
		<div class="content">
		<div class="listing">
			<form name="CVForm" action="vikarCVnyPersonalia.asp?VikarID=<%=lngVikarID%>" method="post">
				<input type="hidden" id="rerun" name="rerun" value="1">
				<input Type="hidden" Name="pbnDataAction" Value="lagre">
				<input type="hidden" id="vikarID" name="vikarid" value="<%=lngVikarID%>">
				<table cellpadding='0' cellspacing='1'>
					<tr>
						<td>Navn: </td>
						<td><%=cons.DataValues("Fornavn")%>&nbsp;<%=cons.DataValues("Etternavn")%></td>
					</tr>
					<tr>
						<td>Adresse:</td>
						<td><%=adr.DataValues("Adresse")%>&nbsp;</td>
					</tr>
					<tr>
						<td></td>
						<td><%=adr.DataValues("Postnr")%>&nbsp;<%=adr.DataValues("Poststed")%></td>
					</tr>
					<tr>
						<td>Telefon:</td>
						<td><%=cons.DataValues("Telefon")%>&nbsp;</td>
					</tr>
					<tr>
						<td>Mobil:</td>
						<td><%=cons.DataValues("MobilTlf")%>&nbsp;</td>
					</tr>
					<tr>
						<td>E-post:</td>
						<td><%=cons.DataValues("EPost")%>&nbsp;</td>
					</tr>
					<tr>
						<td>F&oslash;dselsdato:</td>
						<td><%=cons.DataValues("Foedselsdato")%>&nbsp;</td>
					</tr>
					<tr>
						<td>Oppsigelsestid:</td>
						<td><%=strOppsigelsestid%>&nbsp;</td>
					</tr>
					<tr>
						<td>Stillingsbrøk:</td>
						<td><%=strWorkType%>&nbsp;</td>
					</tr>
					<tr>
						<td>Nasjonalitet:</td>
						<td><%=strCountryName%>&nbsp;</td>
					</tr>
					<tr>
						<td colspan="2"></td>
					</tr>					
				</table>
				<table cellpadding='0' cellspacing='1'>
					<tr>
						<td><input type="checkbox" class="checkbox" name="forerkort" value="1" <%if harForerkort = 1 then%> checked="true" <%end if%>>Jeg har førerkort</td>
						<td><input type="checkbox" class="checkbox" name="bildisponering" value="1" <%if disponererBil = 1 then%> checked="true" <%end if%>>Jeg disponerer bil</td>
					</tr>
				</table>
				<br>
				<span class="menuInside" title="Lagre informasjonen"><a href="#" onClick="javascript:lagre_data();"><img src="images/icon_save.gif" width="18" height="15" alt="" border="0" align="absmiddle">Lagre</a></span>
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
</div>
</body>
</html>

