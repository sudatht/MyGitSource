<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->

<%
'File purpose:		Shows all available jobwishes in a table, and lets the
'					the user select/deselect and write a commentary for it.
'					The data is processed in vikarjobbonskelistedb.asp

	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim StrSQL				'Holds temporary SQL
	dim StrVikarnavn		'holds Name of consultant
	dim rsVikar				'as adodb.recordset
	dim rsProfession			'as adodb.recordset
	dim rsAreas				'as adodb.recordset
	dim LngVikarid			'Holds id of consultant
	dim intTittelID			'Holds id of current Jobwish in recordset loop
	dim IntSelectedAreaID	'Holds id of current Jobwish Area
	dim StrChecked			'Is Current value checked "" or "CHECKED"
	dim strSelected			'Is Current value selected "" or "selected"
	dim lngAreaID			'Holds id of current Jobwish area in recordset loop
	dim Conn				'Connection to database
	dim intErfaringNiva
	dim strSelected1
	dim strSelected2
	dim strSelected3
	dim intUtdannelseNiva
	dim strSelected4
	dim strSelected5
	dim strSelected6
	dim rsTest
	dim lngDisplayType

	'Hotlist menu items variables
	dim strAddToHotlistLink
	dim strHotlistType
	dim blnShowHotList

	'Consultant menu variables
	dim strClass
	dim strJSEvents

	' Initialize values..
	profil = Session("Profil")
	brukerID = Session("BrukerID")
	
	'Get Consultant id
	If Request("VikarID") <> "" Then
	   LngVikarid = Request("VikarID")
	Else
	   LngVikarid = 0
	   Response.write "Error in Parameter. VikarID has no value!"
	   Response.end
	End If

	'Get product area
	If Request("dbxArea") <> "" Then
	   IntSelectedAreaID = Request("dbxArea")
	Else
	   IntSelectedAreaID = 0
	End If

	If Request("dbxShowAll") <> "" Then
	   lngDisplayType = clng(Request("dbxShowAll"))
	Else
	   lngDisplayType = 0
	End If

	' initialize and connect to database
	Set objCon = GetConnection(GetConnectionstring(XIS, ""))

    ' Get consultant name and store for later use
	set rsVikar = GetFirehoseRS("select Navn = fornavn + ' ' + etternavn, Kompetansedato from vikar where vikarid=" & LngVikarid, objCon)    
	StrVikarnavn = rsVikar("navn")
	strDato = rsVikar("Kompetansedato")
	rsVikar.close
    Set rsVikar = Nothing

	'sjekk for å forhindre dobbelreg i Hotlist
	set rsTest = GetFirehoseRS("Select * from HOTLIST Where status=3 And BrukerID=" & brukerID & " And navnID=" & lngVikarID, objCon)    
	if (rsTest.EOF) then
		blnShowHotList = true
		strAddToHotlistLink = "addHotlist.asp?kode=3&vikarNavn=" & server.URLEncode(strEtternavn & " " & strFornavn) & "&vikarNr=" & lngVikarID
		strHotlistType = "vikar"
	else
		blnShowHotList = true
		strAddToHotlistLink = ""
		strHotlistType = "vikar"
	end if
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
		<title>Velg kompetanse</title>
		<script language='javascript' src='/xtra/js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='/xtra/js/navigation.js' id='navigationScripts'></script>					
		<script language="javaScript" type="text/javascript">
		<!--
			function endre_gruppe() {
				document.all.KompForm.submit();
			}
			function endre_data() {
				document.all.KompForm.dbxEndret.value = "1";
			}

			function Toggle(sType, strKompID) {
				if (sType=="check")
				{
					if (eval("document.KompForm.TittelID"+strKompID+".checked==false"))
					{
						eval("document.KompForm.rdoErfaringniva"+strKompID+"[0].checked=false");
						eval("document.KompForm.rdoErfaringniva"+strKompID+"[1].checked=false");
						eval("document.KompForm.rdoErfaringniva"+strKompID+"[2].checked=false");
						eval("document.KompForm.rdoUtdannelseNiva"+strKompID+"[0].checked=false");
						eval("document.KompForm.rdoUtdannelseNiva"+strKompID+"[1].checked=false");
						eval("document.KompForm.rdoUtdannelseNiva"+strKompID+"[2].checked=false");
					};
				}else if (sType=="rdo"){
					eval("document.KompForm.TittelID"+strKompID+".checked=true");
				}
			}
		//-->
		</script>
		<script type="text/javascript" src="/xtra/js/CVFunctions.js"></script>
		<script type="text/javascript" src="/xtra/js/contentMenu.js"></script>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Velg fagkompetanse for <%=StrVikarnavn%></h1>
				<div class="contentMenu">
					<table cellpadding="0" cellspacing="0" width="96%" ID="Table1">
						<tr>
							<td>
								<%
								strClass = "menu disabled"
								strJSEvents = ""
								%>
								<table cellpadding="0" cellspacing="2" ID="Table2">
										<td class="<%=strClass%>" id="menu1" <%=strJSEvents%>>
											<img src="/xtra/images/icon_save.gif" width="18" height="15" alt="" align="absmiddle">Lagre
										</td>
										<td class="menu" id="menu2" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
											<a href="/xtra/vikarvis.asp?vikarid=<%=lngVikarID%>" title="Vis vikar">
											<img src="/xtra/images/icon_consultant.gif" width="18" height="15" alt="" align="absmiddle">Vis</a>
										</td>
									<%
									If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE)) Then
									%>
										<td class="menu" id="menu3" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
										<form ACTION="vikarny.asp?VikarID=<%=lngVikarID%>" METHOD="POST" id="frmConsultantChange">
											<input NAME="pbnDataAction" TYPE="hidden" VALUE="Endre kons.opplysninger" ID="Hidden5">
											<a href="javascript:document.all.frmConsultantChange.submit();" title="Endre vikar">
											<img src="/xtra/images/icon_changeInfo.gif" width="18" height="15" alt="" align="absmiddle">Endre</a>
										</form>
										</td>
										<td class="menu" id="menu4" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);"><strong>CV</strong>&nbsp;<select id="cboCVChoice" onChange="javascript:Vis_CV(<%=lngVikarID%>);" NAME="cboCVChoice"><option value="0"></option><option value="1">Se</option><option value="2">Endre</option><option value="3">Presentere</option></select></td>
										 
									<%
									End If
									If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ)) Then
									%>
										<td class="menu" id="menu6" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
										<form ACTION="vikar-kunder.asp?VikarID=<%=lngVikarID%>" METHOD="POST" id="frmConsultantFormerClients">
											<a href="javascript:document.all.frmConsultantFormerClients.submit();" title="Vis tidligere oppdragsgivere"><img src="/xtra/images/icon_tidl-kunder.gif" alt="" width="18" height="15" border="0" align="absmiddle">Tidligere oppdragsgivere</a>
										</form>
										</td>
										<td class="menu" id="menu7" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
										<form ACTION="Aktivitet-vikar.asp?VikarID=<%=lngVikarID%>" METHOD="POST" id="frmConsultantActivities">
											<a href="javascript:document.all.frmConsultantActivities.submit();" title="Vis aktiviteter for vikaren"><img src="/xtra/images/icon_activities.gif" alt="" width="18" height="15" border="0" align="absmiddle">Aktiviteter</a>
										</form>
										</td>
									</tr>
									<%
									End If
									%>
								</table>
							</td>
							<td class="right">
							<!--#include file="Includes/contentToolsMenu.asp"-->
							</td>
						</tr>
					</table>
				</div>
			</div>
			<form Name="KompForm" Action="Vikarjobbonskelistedb.asp" METHOD="POST" ID="KompForm">
				<input NAME="VikarID" TYPE="HIDDEN" VALUE="<%=LngVikarid%>" ID="VikarID">
				<input NAME="dbxEndret" TYPE="HIDDEN" VALUE="0" ID="dbxEndret">
				<input NAME="dbxOldArea" TYPE="HIDDEN" VALUE="<%=IntSelectedAreaID%>" ID="dbxOldArea">
				<input NAME="dbxJobwishSource" TYPE="HIDDEN" VALUE="vikarJobbonskeliste.asp" ID="dbxJobwishSource">
	 			<div class="content">
					Velg fagomr&aring;de:
					<select NAME="dbxArea" onchange="endre_gruppe()" ID="dbxArea">
					<%
					'Get all availble profession areas
					set rsAreas = GetFirehoseRS("SELECT * FROM [H_KOMP_FAGOMRADE]", objCon)
					if clng(IntSelectedAreaID) = 0 then
						Response.Write "<option value=""0"">Alle</option>"
					else
						Response.Write "<option value=""0"" selected>Alle</option>"
					end if
					Do Until rsAreas.EOF
						lngAreaID = rsAreas("FagID")
						If clng(lngAreaID) = clng(IntSelectedAreaID) Then
							strSelected = "selected"
						Else
							strSelected = ""
						End If
						Response.Write "<option value='" & lngAreaID & "' " & strSelected & ">" & rsAreas("Fagomrade") & "</Option>" & vbcrlf
						rsAreas.MoveNext
					Loop
					rsAreas.close
					Set rsAreas = Nothing
					%>
					</select>
					vis:
					<select NAME="dbxShowAll" onchange="endre_gruppe()" ID="dbxShowAll">
						<%if (lngDisplayType=0) then
							strSelected = "selected"
						else
							strSelected = ""
						end if%>
						<option value="0" <%=strSelected%>>Valgte</option>
						<%if (lngDisplayType=1) then
							strSelected = "selected"
						else
							strSelected = ""
						end if%>
						<option value="1" <%=strSelected%>>Alle</option>
					</select>
					(Kompetanse sist oppdatert: <strong><% =strDato %></strong>)
					<br><br>
					<span class="menuInside"" title="Lagre informasjonen"><a href="#" onClick="document.all.KompForm.submit();"><img src="images/icon_save.gif" width="18" height="15" alt="" border="0" align="absmiddle">Lagre</a></span>
					<div class="listing">
						<table cellpadding="0" cellspacing="1" ID="Table3">
							<tr>
								<th>&nbsp;</th>
								<th>fagkompetanse</th>
								<%
								if IntSelectedAreaID = 0 then
								%>
								<th>Omr&aring;de</th>
								<%
								end if
								%>
								<th class="center">Erfaring<br><span class="normal">lite|noe|mye</span></th>
								<th class="center">Utdannelse<br><span class="normal">lite|noe|mye</span></th>
								<th>kommentar</th>
								<th>web?</th>
							</tr>
							<%
							if lngDisplayType=1 then
								StrSQL = "[GetAllProfessionsListForConsultant] " & LngVikarid & "," & IntSelectedAreaID
							else
								StrSQL = "[GetProfessionListForConsultant] " & LngVikarid & "," & IntSelectedAreaID
							end if
							Set rsProfession = GetFirehoseRS(StrSQL, objCon)							
							Do Until (rsProfession.EOF)
								intTittelID = clng(rsProfession.fields("K_TittelID").value)
								StrChecked = cstr(rsProfession.fields("besitter").value)
								%>
								<tr>
									<td><input class="checkbox" onchange="endre_data()" onClick="Toggle('check','<%=intTittelID%>')" id="TittelID<%=intTittelID%>" name="<%=intTittelID%>TittelID" type="checkbox" <%=strchecked%>></td><td><%=rsProfession.fields("ktittel").value%>&nbsp;</td>
								<%
								if IntSelectedAreaID = 0 then
									%> <td><%=rsProfession.fields("fagomrade").value%></td><%
								end if

								if isnull(rsProfession.fields("Relevant_WorkExperience").value) then
									intErfaringNiva =	0
								else
									intErfaringNiva = cint(rsProfession.fields("Relevant_WorkExperience").value)
								end if

								if intErfaringNiva = 1 then
									strSelected1 = "checked"
									strSelected2 = ""
									strSelected3 = ""
								elseif intErfaringNiva = 2 then
									strSelected1 = ""
									strSelected2 = "checked"
									strSelected3 = ""
								elseif intErfaringNiva = 3 then
									strSelected1 = ""
									strSelected2 = ""
									strSelected3 = "checked"
								else
									strSelected1 = ""
									strSelected2 = ""
									strSelected3 = ""
								end if
								%>
								<td>
									<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="rdoErfaringniva<%=intTittelID%>" name="<%=intTittelID%>rdoErfaringniva" value="1" <%=strSelected1%>>
									<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="Radio1" name="<%=intTittelID%>rdoErfaringniva" value="2" <%=strSelected2%>>
									<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="Radio2" name="<%=intTittelID%>rdoErfaringniva" value="3" <%=strSelected3%>>
								</td>
								<%
								if isnull(rsProfession.fields("Relevant_Education").value) then
									intUtdannelseNiva =	0
								else
									intUtdannelseNiva = cint(rsProfession.fields("Relevant_Education").value)
								end if

								if intUtdannelseNiva = 1 then
									strSelected4 = "checked"
									strSelected5 = ""
									strSelected6 = ""
								elseif intUtdannelseNiva = 2 then
									strSelected4 = ""
									strSelected5 = "checked"
									strSelected6 = ""
								elseif intUtdannelseNiva = 3 then
									strSelected4 = ""
									strSelected5 = ""
									strSelected6 = "checked"
								else
									strSelected4 = ""
									strSelected5 = ""
									strSelected6 = ""
								end if
								%>
								<td>
									<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="rdoUtdannelseNiva<%=intTittelID%>" name="<%=intTittelID%>rdoUtdannelseNiva" value="1" <%=strSelected4%>>
									<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="Radio3" name="<%=intTittelID%>rdoUtdannelseNiva" value="2" <%=strSelected5%>>
									<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="Radio4" name="<%=intTittelID%>rdoUtdannelseNiva" value="3" <%=strSelected6%>>
								</td>
								<td><input size="30" onchange="endre_data()" maxlength="256" name="<%=intTittelID%>tbxKommentar" value="<%=rsProfession.fields("kommentar").value%>" ID="Text1"></td>
								<%
								if (rsProfession.fields("web_jobwish_visible").value = true) then
									%>
									<td><img src="\xtra\images\published_true.gif"></td></tr>
									<%
								else
									%>
									<td>&nbsp;</td>
								</tr>
									<%
								end if
								rsProfession.MoveNext
							Loop
							Set rsProfession = Nothing
							%>
							</tr>
							</table>
							<br>
							<span class="menuInside"" title="Lagre informasjonen"><a href="#" onClick="document.all.KompForm.submit();"><img src="images/icon_save.gif" width="18" height="15" alt="" border="0" align="absmiddle">Lagre</a></span>
						</div>
				</form>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(objCon)
set objCon = nothing
%>