<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<%

	'File purpose:		Shows all available qualifications in a table, and lets the
	'					the user select/deselect, rank, write a commentary and level it.
	'					The data is processed in vikarkomplistedb.asp
	'
	dim StrSQL				'Holds temporary SQL
	dim StrVikarnavn		'holds Name of consultant
	dim rsVikar				'as adodb.recordset
	dim rsKompetanse		'as adodb.recordset
	dim rsLevel				'as adodb.recordset
	dim rsAreas				'as adodb.recordset
	dim intRangering		'Holds rank of current Qualification in recordset loop
	dim intTittelID			'Holds id of current Qualification in recordset loop
	dim IntSelectedAreaID	'Holds id of current Qualification Area
	dim StrChecked			'Is Current value checked "" or "CHECKED"
	dim strSelected			'Is Current value selected "" or "selected"
	dim strLevel			'Holds level of current Qualification in recordset loop
	dim IntAreaID			'Holds id of current Qualification area in recordset loop
	dim ObjCon				'Holds connection to Xtra's DB
	dim strDato				'Dato containg last update
	dim strSelected3
	dim strSelected4
	dim strSelected5
	dim strSelected6
	dim rsTest
	dim profil
	dim brukerID
	dim rsVik
	dim lngDisplayType

	'Hotlist menu items variables
	dim strAddToHotlistLink
	dim strHotlistType
	dim LngVikarid
	dim blnShowHotList

	dim strEtternavn
	dim strFornavn

	'Consultant menu variables
	dim strClass
	dim strJSEvents

    dim ccAction
    
    ccAction=false
	' Initialize values..
	profil = Session("Profil")
	brukerID = Session("BrukerID")

	'Get Consultant id
	If Request.QueryString("VikarID") <> "" Then
	   LngVikarid = CLng( Request.QueryString("VikarID") )
	elseif Request.Form("VikarID") <> "" then
	   LngVikarid = CLng( Request.form("VikarID") )
	Else
	   LngVikarid = 0
	   Response.write "Error in Parameter. VikarID has no value!"
	   Response.end
	End If

	'Get product area
	If Request.QueryString("dbxArea") > 0 Then
	   IntSelectedAreaID = CLng( Request.QueryString("dbxArea") )
	Else
	   IntSelectedAreaID = 0
	End If

	If Request("dbxShowAll") <> "" Then
	   lngDisplayType = clng(Request("dbxShowAll"))
	Else
	   lngDisplayType = 0
	End If

	' initialize and connect to database
	Set ObjCon = GetConnection(GetConnectionstring(XIS, ""))	
	
    ' Get consultant name and store for later use
    set rsVikar = GetFirehoseRS("select Navn = fornavn + ' ' + etternavn from vikar where vikarid=" & LngVikarid, objCon)
	StrVikarnavn = rsVikar("navn")
	rsVikar.close
	Set rsVikar = Nothing

	strSQL = "select [Kompetansedato] from [VIKAR] where [VikarID] = " & LngVikarid

	set rsVik = GetFirehoseRS(strSQL, objCon)
	If Not rsVik.EOF Then
		strDato = rsVik("Kompetansedato")
	End If
	rsVik.close
	Set rsVik = Nothing

	'sjekk for å forhindre dobbelreg i Hotlist
	strSQL = "Select * from HOTLIST Where status=3 And BrukerID=" & brukerID & " And navnID=" & lngVikarID
	set rsTest = GetFirehoseRS(strSQL, objCon)
	if (rsTest.EOF) then
		blnShowHotList = true
		strAddToHotlistLink = "addHotlist.asp?kode=3&vikarNavn=" & server.URLEncode(strEtternavn & " " & strFornavn) & "&vikarNr=" & lngVikarID
		strHotlistType = "vikar"
	else
		blnShowHotList = true
		strAddToHotlistLink = ""
		strHotlistType = "vikar"
	end if
	rsTest.close
	set rsTest = nothing

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<meta NAME="GENERATOR" Content="Microsoft FrontPage 5.0">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<title>Velg kompetanse</title>
		<script language="javaScript" type="text/javascript">
		<!--
			function endre_gruppe() {
				document.KompForm.submit();
			}

			function endre_data() {
				document.KompForm.dbxEndret.value = "1";
			}

			function Toggle(sType, strKompID) {
				if (sType=="check")
				{
					if (eval("document.KompForm.TittelID"+strKompID+".checked==false"))
					{
						eval("document.KompForm.rdorangering"+strKompID+"[0].checked=false");
						eval("document.KompForm.rdorangering"+strKompID+"[1].checked=false");
						eval("document.KompForm.rdorangering"+strKompID+"[2].checked=false");
						eval("document.KompForm.rdorangering"+strKompID+"[3].checked=false");
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
				<h1>Velg produktkompetanse for <%=StrVikarnavn%></h1>
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
									If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
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
								</table>
							</td>
							<td class="right">
							<!--#include file="Includes/contentToolsMenu.asp"-->
							</td>
						</tr>
					</table>
				</div>
			</div>
			<form Name="KompForm" Action="VikarKomplisteDB.asp" METHOD="POST" ID="Form1">
			
				<input NAME="VikarID" TYPE="HIDDEN" VALUE="<%=LngVikarid%>" ID="Hidden1">
				<input NAME="dbxEndret" TYPE="HIDDEN" VALUE="0" ID="Hidden2">
				<input NAME="dbxOldArea" TYPE="HIDDEN" VALUE="<%=IntSelectedAreaID%>" ID="Hidden3">
				<input NAME="dbxQualificationSource" TYPE="HIDDEN" VALUE="Vikarkompliste.asp" ID="Hidden4">
	 			<div class="content">
					&nbsp;Velg produktomr&aring;de:
					<select NAME="dbxArea" onchange="endre_gruppe()" ID="Select1">
					<%
					'Get all available product areas
					set rsAreas = GetFirehoseRS("SELECT * FROM [H_KOMP_AREA]", ObjCon)
					if IntSelectedAreaID = 0 then
						Response.Write "<option VALUE=""0"">Alle</option>"
					else
						Response.Write "<option VALUE=""0"" SELECTED>Alle</option>"
					end if
					Do Until rsAreas.EOF
						If rsAreas("ProdOmradeID") = IntSelectedAreaID Then
							IntAreaID = rsAreas("ProdOmradeID")
							strSelected = "SELECTED"
						Else
							IntAreaID = rsAreas("ProdOmradeID")
							strSelected = ""
						End If
						Response.Write "<option VALUE=""" & IntAreaID & """ " & strSelected & ">" & rsAreas("Produktomrade") & "</option>"
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
								<th>Produktkompetanse</th>
								<%
								if IntSelectedAreaID = 0 then
									%>
									<th>Omr&aring;de</th>
									<%
								end if
								%>
								<th>Nivå</th>
								<th>Bruker niv&aring;<br>
									<span class="normal"><img src="images/levels.gif"></span>
								</th>
								<th>kommentar</th>
								<th>web?</th>
								</tr>
									<%
									if lngDisplayType = 1 then
										StrSQL = "[GetAllQualificationsListForConsultant] " & LngVikarid & "," & IntSelectedAreaID
									else
										StrSQL = "[GetQualificationListForConsultant] " & LngVikarid & "," & IntSelectedAreaID
									end if
									set rsKompetanse = GetFirehoseRS(StrSQL, ObjCon)

									' Get Qualification levels..
									set rsLevel = GetDynamicRS("select K_LevelID, Klevel from H_KOMP_LEVEL" , ObjCon)							

									while (not rsKompetanse.EOF)
										intTittelID = clng(rsKompetanse.fields("K_TittelID").value)
										StrChecked = cstr(rsKompetanse.fields("besitter").value)
										%>
										<tr>
											<td>
                                            <input onChange="endre_data('check','<%=intTittelID%>')" onClick="Toggle('check','<%=intTittelID%>')" id="TittelID<%=intTittelID%>" name="<%=intTittelID%>TittelID" class="checkbox" type="checkbox" <%=strchecked%> value="ON"></td>
											<td><%=rsKompetanse.fields("ktittel").value%>&nbsp;</td>
											<%
											if IntSelectedAreaID = 0 then
												Response.Write "<td>" & rsKompetanse.fields("ProduktOmrade").value & "&nbsp;" & "</td>" & chr(13)
											end if
											Response.Write "<td><select name=""" & intTittelID & "dbxLevel"" onchange=""endre_data()""><option VALUE=""0"">"

											rsLevel.movefirst
											While (NOT rsLevel.EOF)
												If rsLevel("K_LevelID") = rsKompetanse.fields("K_LevelID").value Then
													strLevel = rsLevel("K_LevelID")
													strSelected = "SELECTED"
												Else
													strLevel = rsLevel("K_LevelID")
													strSelected = ""
												End If
												Response.Write "<option VALUE=""" & strLevel & """ " & strSelected & ">" & rsLevel("KLevel") & "</option>" &  vbCrLf
												rsLevel.MoveNext
											Wend
											%>
											</select></td>
											<%

											if isnull(rsKompetanse.fields("rangering").value) then
												intRangering =	0
											else
												intRangering =	cint(rsKompetanse.fields("rangering").value)
											end if

											if intRangering = 3 then
												strSelected3 = "checked"
												strSelected4 = ""
												strSelected5 = ""
												strSelected6 = ""
											elseif intRangering = 4 then
												strSelected3 = ""
												strSelected4 = "checked"
												strSelected5 = ""
												strSelected6 = ""
											elseif intRangering = 5 then
												strSelected3 = ""
												strSelected4 = ""
												strSelected5 = "checked"
												strSelected6 = ""
											elseif intRangering = 6 then
												strSelected3 = ""
												strSelected4 = ""
												strSelected5 = ""
												strSelected6 = "checked"
											else
												strSelected3 = ""
												strSelected4 = ""
												strSelected5 = ""
												strSelected6 = ""
											end if
											%>
											<td>
												<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="rdorangering<%=intTittelID%>" name="<%=intTittelID%>rdorangering" value="3" <%=strSelected3%>>
												<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="Radio1" name="<%=intTittelID%>rdorangering" value="4" <%=strSelected4%>>
												<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="Radio2" name="<%=intTittelID%>rdorangering" value="5" <%=strSelected5%>>
												<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="Radio3" name="<%=intTittelID%>rdorangering" value="6" <%=strSelected6%>>
											</td>
											<td>
												<input onchange="endre_data()" size="30" maxlength="256" name="<%=intTittelID%>tbxKommentar" value="<%=rsKompetanse.fields("kommentar").value%>" ID="Text1">
											</td>
											<%
											if rsKompetanse.fields("web_komp_visible").value then
												Response.Write "<td><img src=""\xtra\images\published_" & rsKompetanse.fields("web_komp_visible").value & ".gif""></td></tr>" & chr(13)
											else
												%>
												<td>&nbsp;</td>
											</tr>
											<%
										end if
								rsKompetanse.MoveNext
							wend
							rsKompetanse.close
							Set rsKompetanse = Nothing
							' Close and release nivå recordset
							rsLevel.close
							Set rsLevel = Nothing
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