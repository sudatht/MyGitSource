<%@ LANGUAGE="VBSCRIPT" %>
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

	dim edu
	dim cvfil 'as string
	dim blnLocked 			'as boolean
	dim fulltime
	dim beskrivelse
	dim edu_id
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

	lVikarID    = Request.Querystring("VikarID")
	lngVikarID	= CStr(lVikarID)
	cons.XtraConString = Application("XtraWebConnection")
	cons.GetConsultant(lVikarID)

	dette_aar = Year(date)

	set cv	= cons.CV
	cv.XtraConString = Application("Xtra_intern_ConnectionString")
	cv.XtraDataShapeConString = Application("ConXtraShape")
	cv.Refresh

	'If Consultant doesn't have CV, create one
	if cons.CV.DataValues.Count = 0 then
		cons.CV.Save
	end if

	blnLocked = cv.islocked

	set allEdu = cv.Educations
	set edu	= Server.CreateObject("XtraWeb.Education")
	set edu.Owner = cv

	edu_id = 0

	if Request.QueryString("eduid").Count > 0 then
		edu_id = Request.QueryString("eduid")
		set edu = cv.Educations("ID" & edu_id)
		utdKomm  = edu.DataValues("Description")
	end if

	if Request.Form("rerun") = "1" then
		strAction  = lcase(trim(request("pbnDataAction")))

		' Utdannelse...
		if (strAction = "lagre") then

			edu_id = clng(Request.Form("objId"))
			if edu_id = null then
				edu_id = 0
			end if
			if edu_id > 0 then
				set edu = cv.Educations("ID" & edu_id)
			end if
			utdFraMnd	= CStr(Request.Form("eduFraMnd"))
			utdFraAar	= CStr(Request.Form("eduFraAar"))
			utdTilMnd	= CStr(Request.Form("eduTilMnd"))
			utdTilAar	= CStr(Request.Form("eduTilAar"))
			utdSted		= CStr(Request.Form("eduSted"))
			utdLinje	= CStr(Request.Form("eduLinje"))
			utdKomm		= CStr(Request.Form("Editor1"))

			if utdFraMnd <> "" and utdFraAar <> "" then
				if CInt(utdFraMnd) < 1 or CInt(utdFraMnd) > 12 or CInt(utdFraAar) < 0 or CInt(utdFraAar) > dette_aar then
					feilmelding = feilmelding & "Ugyldig måned eller årstall i feltet Fra!<br>"
				end if
			end if
			if (utdTilMnd="") and (utdTilAar="")  then
			elseif CInt(utdTilMnd) < 1 or CInt(utdTilMnd) > 12 or CInt(utdTilAar) < 0 or CInt(utdTilAar) > dette_aar then
					feilmelding = feilmelding & "Ugyldig måned eller årstall i feltet Til!<br>"
			end if

			if utdFraMnd <> "" and utdFraAar <> "" and utdSted <> "" and utdLinje <> "" and feilmelding = "" then
				edu.DataValues("FromMonth")	= utdFraMnd
				edu.DataValues("FromYear")	= utdFraAar
				if (utdTilMnd<>"") and (utdTilAar<>"")  then
					edu.DataValues("ToMonth") = utdTilMnd
					edu.DataValues("ToYear") = utdTilAar

				else
					edu.DataValues("ToMonth") = empty
					edu.DataValues("ToYear") = empty
				end if
				edu.DataValues("Place")			= utdSted
				edu.DataValues("Title")			= utdLinje
				edu.DataValues("Description") 	= utdKomm
				if not edu.Save	then
					Response.Write "Ikke Lagret!<br>"
					response.end
				end if
				set cv	= nothing
				cons.CV.cleanup
				cons.cleanup
				set cons	= nothing
				Response.Redirect("vikarCVnyUtdannelse.asp?type=4&VikarID=" & lngVikarID)
			else
				feilmelding = feilmelding & "Alle opplysninger bortsett fra Kommentarer er obligatoriske!<br>"
			end if
		end if

		if strAction = "slett" then
			edu_id = clng(Request.Form("objId"))
			if edu_id <> 0 then
				set edu = cv.Educations("ID" & edu_id)
				if not edu.Delete then
					Response.Write "Ikke slettet!<br>"
				end if
				set cv		= nothing
				cons.CV.cleanup
				cons.cleanup
				set cons	= nothing
				Response.Redirect("vikarCVnyUtdannelse.asp?type=4&VikarID=" & lngVikarID)
			end if
		end if
	end if
	utdKomm = "<link type='text/css' rel='stylesheet' href='http://" & Application("HTTPadress") & "/xtra/css/CV.css' title='default style'>" & utdKomm
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
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
			function slett_data(id) 
			{
				document.all.pbnDataAction.value='slett';
				document.CVForm.objId.value = id;
				document.CVForm.submit();
			}
			function lagre_data(id) 
			{
				document.all.pbnDataAction.value='lagre';
				document.CVForm.objId.value = id;
				document.CVForm.submit();
			}
		</script>
		<script type="text/javascript" src="/xtra/js/CVFunctions.js"></script>
		<script type="text/javascript" src="/xtra/js/contentMenu.js"></script>		
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<H1 style="font-size:160%; font-weight:bold";> CV redigering - <%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & lngVikarID, Cons.DataValues("Fornavn") & " " & Cons.DataValues("Etternavn"), "Vis vikar " & Cons.DataValues("Fornavn") & " " & Cons.DataValues("Etternavn") )%></H1>
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
				<span class="menu2 active" id="5"><strong>Utdannelse</strong></span>
				<span class="menu2" id="6" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyCourses.asp?VikarID=<%=lngVikarID%>">Kurs</a></span>
				<span class="menu2" id="7" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyPraksis.asp?VikarID=<%=lngVikarID%>">Yrkeserfaring</a></span>
				<span class="menu2" id="8" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyAndreOppl.asp?VikarID=<%=lngVikarID%>">Kjernekompetanse</a></span>
				<span class="menu2" id="9" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyReferanser.asp?VikarID=<%=lngVikarID%>">Referanser</a></span>
				<span class="menu2" id="10" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVvis.asp?VikarID=<%=lngVikarID%>">&nbsp;<img src="/xtra/images/icon_new2.gif" width="14" height="14" alt="" border="0" align="absmiddle">&nbsp;Generere CV</a></span>
			</div>
			<div class="content">
				<form name="CVForm" action="vikarCVnyUtdannelse.asp?VikarID=<%=lngVikarID%>" method="post">
					<input type="hidden" id="rerun" name="rerun" value="1">
					<input Type="hidden" Name="pbnDataAction" Value="lagre">
					<input type="hidden" id="education" name="education">
					<input type="hidden" id="objId" name="objId">
					<table width="96%" class="layout" cellpadding="0" cellspacing="0">
						<col width="80%">
						<col width="20%">
					<tr>
						<td>
						<table width="100%" border="0" cellpadding='0' cellspacing='0'>
							<tr>
								<td width="100">Fra (mm/åååå):</td>
								<td width="30"><input type="text" name="eduFraMnd" size="2" maxlength="2"
									<%
									if edu_id <> 0 then
									%> value="<%if edu.DataValues("FromMonth") < 10 then%>0<%end if%>
									<%=edu.DataValues("FromMonth")%>"
									<%else%> value="<%
										if utdFraMnd <> "" then
											if ((CInt(utdFraMnd) < 10) and (CInt(utdFraMnd) > 0))  then
												%>0<%
											end if
										end if%>
									<%=utdFraMnd%>"
									<%
									end if
									%> class="mandatory"></td>
								<td width="10">&nbsp;&nbsp;/&nbsp;&nbsp;</td>
								<td width="60"><input type="text" name="eduFraAar" size="4" maxlength="4"
									<%if edu_id <> 0 then%> value="<%=edu.DataValues("FromYear")%>"
									<%else%> value="<%=utdFraAar%>" <%end if%> class="mandatory">
								</td>
								<td width="100">Til (mm/åååå):</td>
								<td width="30"><input type="text" name="eduTilMnd" size="2" maxlength="2"
									<%if edu_id <> 0 then%> value="<%if edu.DataValues("ToMonth") < 10 then%>0<%end if%> <%=edu.DataValues("ToMonth")%>"
									<%else%> value="<%if utdTilMnd <> "" then
										if ((CInt(utdTilMnd) < 10) and (CInt(utdTilMnd) > 0)) then
											%>0<%
										end if
									end if%> <%=utdTilMnd%>" <%end if%> class="mandatory">
								</td>
								<td width="10">&nbsp;&nbsp;/&nbsp;&nbsp;</td>
								<td width="60"><input type="text" name="eduTilAar" size="4" maxlength="4"
									<%if edu_id <> 0 then%> value="<%=edu.DataValues("ToYear")%>"
									<%else%> value="<%=utdTilAar%>" <%end if%> class="mandatory">
								</td>
								<td width="200">&nbsp;</td>
							</tr>
							<tr>
								<td width="100">Sted:</td>
								<td colspan="8"><input type="text" name="eduSted" size="40" maxlength="50"
									<%if edu_id <> 0 then%> value="<%=edu.DataValues("Place")%>"
									<%else%> value="<%=utdSted%>" <%end if%> class="mandatory">
								</td>
							</tr>
							<tr>
								<td width="100">Linje:</td>
								<td colspan="8"><input type="text" name="eduLinje" size="40" maxlength="50"
									<%if edu_id <> 0 then%> value="<%=edu.DataValues("Title")%>"
									<%else%> value="<%=utdLinje%>" <%end if%> class="mandatory">
								</td>
							</tr>
							<tr>
								<td width="100">&nbsp;</td>
								<td colspan="8">&nbsp;</td>
							</tr>
							<tr>							
								<td width="100">Kommentarer:</td>
								<td colspan="8">
									<%
								    editor.Text = utdKomm
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
				<span class="menuInside" style="margin-left:100px;" title="Lagre informasjonen"><a href="#" onClick="javascript:lagre_data(<%=edu_id%>);"><img src="images/icon_save.gif" width="18" height="15" alt="" border="0" align="absmiddle">Lagre</a></span>
				<div class="listing">
				<table cellpadding='0' cellspacing='1'>
					<tr>
						<th colspan="7">Registrert utdannelse</th>
					</tr>
					<tr>
						<th>Fra</th>
						<th>Til</th>
						<th>Sted</th>
						<th>Linje</th>
						<th>Kommentarer</th>
						<th class="center">Endre</th>
						<th class="center">Slette</th>
					</tr>
					<%
					for i = 1 to allEdu.Count
						set edu = allEdu.Item(i)
						id = edu.DataValues("DataId")

						StrFromMonth = cint(edu.DataValues("FromMonth").value)
						strFromYear = cint(edu.DataValues("FromYear").value)

						if (isnull(edu.DataValues("ToMonth").value)) then
							StrToMonth 	= cint(0)
						else
							StrToMonth 	= cint(edu.DataValues("ToMonth").value)
						end if


						if (isnull(edu.DataValues("ToYear").value)) then
							strToYear 	=  cint(0)
						else
							strToYear 	= cint(edu.DataValues("ToYear").value)
						end if


						if (StrFromMonth < 10) then
							StrFromMonth = "0" & StrFromMonth
						end if
						StrFromPeriod = StrFromMonth & "/" & strFromYear

						if (StrToMonth < 10) and (StrToMonth >= 1) then
							StrToMonth = "0" & StrToMonth
						end if

						if (StrToMonth = 0) and (strToYear = 0) then
							StrToPeriod = "d.d."
						else
							if(strToYear = 0) then
								StrToPeriod = StrToMonth & "/????"
							else
								StrToPeriod = StrToMonth & "/" & strToYear
							end if
						end if
						%>
						<tr>
							<td><%=StrFromPeriod%></td>
							<td><%=StrToPeriod%></td>
							<td><%=edu.DataValues("Place")%></td>
							<td><%=edu.DataValues("Title")%></td>
							<td><%=edu.DataValues("Description")%>&nbsp;</td>
							<td class="center"><a href="vikarCVnyUtdannelse.asp?type=4&VikarID=<%=lngVikarID%>&eduid=<%=id%>"><img src="/xtra/images/icon_changeInfo.gif" width="18" height="15" alt="Endre denne oppføringen" border="0"></a></td>
							<td class="center"><a href="javascript:slett_data(<%=id%>)"><img src="/xtra/images/icon_delete.gif" width="14" height="14" alt="Slette denne oppføringen" border="0"></a></td>
						</tr>
					<%next%>
				</table>
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