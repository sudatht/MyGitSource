<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<!-- #include file = "cuteeditor_files/include_CuteEditor.asp" --> 
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	'File purpose:		Show/edit key qualification in the HTML editor.
	'Created Date:		05.03.2001
	'Created by:		Fred.myklebust@electricfarm.no
	'Changed by:		-
	'Changed Date:		11.11.02
	'Changes:			-
	'Tested date:		-

	dim StrVikarnavn		'holds Name of consultant
	dim strKeyQualifications			'String containing "key qualifications"
	dim strAction
	dim lngCVID
	dim oCons
	dim oCV
	dim blnLocked
	dim ObjCon				'Holds connection to Xtra's DB
		
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
	'Get Consultant id
	If Request("VikarID") <> "" Then
	   LngVikarid = CLng( Request("VikarID"))
	Else
		AddErrorMessage("Error in Parameter. VikarID has no value!")
		call RenderErrorMessage()
	End If

	set oCons	= Server.CreateObject("XtraWeb.Consultant")

	oCons.XtraConString = Application("XtraWebConnection")
	oCons.GetConsultant(LngVikarid)
	StrVikarnavn = oCons.Datavalues("fornavn") & " " & oCons.datavalues("etternavn")

	set oCV	= oCons.CV
	oCV.XtraConString = Application("Xtra_intern_ConnectionString")
	oCV.XtraDataShapeConString = Application("ConXtraShape")
	oCV.Refresh

	if oCons.CV.DataValues.Count = 0 then
		oCons.CV.Save
	end if

	blnLocked = oCV.islocked

	lngCVID = oCV.datavalues("CVid")

	strAction  = lcase(trim(request("pbnDataAction")))

	if (len(strAction) > 0) then
		strKeyQualifications = request("Editor1")
		oCV.datavalues("Key_Qualifications") = strKeyQualifications
		oCV.save
		if strAction = "ferdig" then
			set oCV		= nothing
			oCons.CV.cleanup
			oCons.cleanup
			set oCons	= nothing
			Response.Clear
			Response.Redirect("vikarCVvis.asp?VikarID=" & LngVikarid)
		end if
	else
		if (not isnull(oCV.datavalues("Key_Qualifications"))) then
			strKeyQualifications = oCV.datavalues("Key_Qualifications").value
		else
			strKeyQualifications = ""
		end if
	end if
	'strKeyQualifications =  "<link type='text/css' rel='stylesheet' href='http://" & Application("HTTPadress") & "/xtra/css/CV.css' title='default style'>"	 & strKeyQualifications
	

	set oCV		= nothing
	oCons.CV.cleanup
	oCons.cleanup
	set oCons	= nothing

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<link REL="stylesheet" TYPE="text/css" HREF="dhtml/toolbars/toolbars.css">
		<title>Nøkkel kvalifikasjoner</title>
		<script type="text/javascript" src="/xtra/js/CVFunctions.js"></script>
		<script type="text/javascript" src="/xtra/js/contentMenu.js"></script		
	</head>
	<body>
		<form Name="KompForm" id="KompForm" Action="VikarCVnyNokkelKvalifikasjoner.asp" METHOD="POST">
			<input NAME="VikarID" TYPE="HIDDEN" VALUE="<%=LngVikarid%>">
			<input Type="hidden" Name="pbnDataAction" Value="lagre">
			<div class="pageContainer" id="pageContainer">
				<div class="contentHead1">
					<h1>CV redigering - <%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & lngVikarID, StrVikarnavn, "Vis vikar " & StrVikarnavn )%></h1>
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
					<span class="menu2" id="1" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyPersonalia.asp?VikarID=<%=LngVikarID%>">Personalia</a></span>
					<span class="menu2" id="2" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyJobbonsker.asp?VikarID=<%=LngVikarID%>">Fagkompetanse</a></span>
					<span class="menu2" id="3" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyKvalifikasjoner.asp?VikarID=<%=LngVikarID%>">Produktkompetanse</a></span>
					<!--
					<span class="menu2 active" id="4"><strong>Kandidatpresentasjon</strong></span>
					-->
					<span class="menu2" id="5" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyUtdannelse.asp?VikarID=<%=LngVikarid%>">Utdannelse</a></span>
					<span class="menu2" id="6" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyCourses.asp?VikarID=<%=lngVikarID%>">Kurs</a></span>
					<span class="menu2" id="7" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyPraksis.asp?VikarID=<%=LngVikarID%>">Yrkeserfaring</a></span>
					<span class="menu2" id="8" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyAndreOppl.asp?VikarID=<%=LngVikarID%>">Kjernekompetanse</a></span>
					<span class="menu2" id="9" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyReferanser.asp?VikarID=<%=LngVikarID%>">Referanser</a></span>
					<span class="menu2" id="10" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVvis.asp?VikarID=<%=lngVikarID%>">&nbsp;<img src="/xtra/images/icon_new2.gif" width="14" height="14" alt="" border="0" align="absmiddle">&nbsp;Generere CV</a></span>
				</div>
				<div class="content">					
					<br/>				
					<br/>				
					<%
						editor.Text = strKeyQualifications
						editor.Draw()	
				    %>
					<br/>
					<br/>
					<span class="menuInside" style="margin-left:35px;" title="Lagre informasjonen"><a href="#" onClick="document.all.KompForm.submit();"><img src="images/icon_save.gif" width="18" height="15" alt="" border="0" align="absmiddle">Lagre</a></span>					
					&nbsp;
				</div>
			</div>
		</form>	
	</body>
</html>