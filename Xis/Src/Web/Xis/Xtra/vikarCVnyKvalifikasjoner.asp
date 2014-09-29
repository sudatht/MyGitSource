<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	'File purpose:		Shows all available qualifications in a table, and lets the
	'					the user select/deselect, rank, write a commentary and level it.
	'					The data is processed in vikarkomplistedb.asp
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
	dim lngAreaID			'Holds id of current Qualification area in recordset loop
	dim ObjCon				'Holds connection to Xtra's DB
	dim strDato				'Dato containg last update
	dim strSelected3
	dim strSelected4
	dim strSelected5
	dim strSelected6
	dim lngDisplayType

	'Global variables - used in tool menu
	dim blnShowHotList
	dim lngVikarID

	'Initialize variables
	blnShowHotList = false
	'Get Consultant id
	If Request.QueryString("VikarID") <> "" Then
	   LngVikarid = CLng( Request.QueryString("VikarID") )
	elseif Request.Form("VikarID") <> "" then
	   LngVikarid = CLng( Request.form("VikarID") )
	Else
	   LngVikarid = 0
	   AddErrorMessage("Feil: Vikarid Parameteren mangler.")
	   call RenderErrorMessage()
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
	Set objCon = GetConnection(GetConnectionstring(XIS, ""))	

   	' Get consultant name and store for later use
   	StrSQL = "select Navn = fornavn + ' ' + etternavn from vikar where vikarid=" & LngVikarid
	set rsVikar = GetFirehoseRS(StrSQL, objCon)
	StrVikarnavn = rsVikar("navn")
	rsVikar.close
	Set rsVikar = Nothing

	strSQL = "select [Kompetansedato] from [VIKAR] where [VikarID] = " & LngVikarid
	set rsVikar = GetFirehoseRS(StrSQL, objCon)
	If HasRows(rsVikar) = true Then
		strDato = rsVikar("Kompetansedato")
	End If
	Set rsVikar = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<title>Velg produktkompetanse</title>
		<script language="javaScript" type="text/javascript">

			function endre_gruppe() 
			{
				document.KompForm.submit();
			}

			function endre_data() 
			{
				document.KompForm.dbxEndret.value = "1";
			}

			function Toggle(sType, strKompID) 
			{
				if (sType=="check")
				{
					if (eval("document.KompForm.TittelID"+strKompID+".checked==false"))
					{
						eval("document.KompForm.rdorangering"+strKompID+"[0].checked=false");
						eval("document.KompForm.rdorangering"+strKompID+"[1].checked=false");
						eval("document.KompForm.rdorangering"+strKompID+"[2].checked=false");
						eval("document.KompForm.rdorangering"+strKompID+"[3].checked=false");
					};
				}
				else if (sType=="rdo")
				{
					eval("document.KompForm.TittelID"+strKompID+".checked=true");
				}
			}
		</script>
		<script type="text/javascript" src="/xtra/js/CVFunctions.js"></script>
		<script type="text/javascript" src="/xtra/js/contentMenu.js"></script>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<form Name="KompForm" Action="Vikarkomplistedb.asp" METHOD="POST" ID="Form1">
				<div class="contentHead1">
					<h1>CV redigering - <%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & lngVikarID, StrVikarnavn, "Vis vikar " & StrVikarnavn )%></h1>
					<input NAME="VikarID" TYPE="HIDDEN" VALUE="<%=LngVikarid%>">
					<input NAME="dbxEndret" TYPE="HIDDEN" VALUE="0">
					<input NAME="dbxOldArea" TYPE="HIDDEN" VALUE="<%=IntSelectedAreaID%>">
					<input NAME="dbxQualificationSource" TYPE="HIDDEN" VALUE="VikarCVnyKvalifikasjoner.asp">
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
					<span class="menu2 active" id="3"><strong>Produktkompetanse</strong></span>
					<!--
					<span class="menu2" id="4" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyNokkelKvalifikasjoner.asp?VikarID=<%=LngVikarID%>">Kandidatpresentasjon</a></span>
					-->
					<span class="menu2" id="5" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyUtdannelse.asp?VikarID=<%=LngVikarid%>">Utdannelse</a></span>
					<span class="menu2" id="6" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyCourses.asp?VikarID=<%=lngVikarID%>">Kurs</a></span>
					<span class="menu2" id="7" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyPraksis.asp?VikarID=<%=LngVikarID%>">Yrkeserfaring</a></span>
					<span class="menu2" id="8" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyAndreOppl.asp?VikarID=<%=LngVikarID%>">Kjernekompetanse</a></span>
					<span class="menu2" id="9" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyReferanser.asp?VikarID=<%=LngVikarID%>">Referanser</a></span>
					<span class="menu2" id="10" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVvis.asp?VikarID=<%=lngVikarID%>">&nbsp;<img src="/xtra/images/icon_new2.gif" width="14" height="14" alt="" border="0" align="absmiddle">&nbsp;Generere CV</a></span>
				</div>
				<div class="content">
					Velg produktomr&aring;de:
					<select NAME="dbxArea" onchange="endre_gruppe()">
						<%
						'Get all available product areas
						StrSQL = "select * from h_komp_area"
						set rsAreas = GetFirehoseRS(StrSQL, objCon)
						if clng(IntSelectedAreaID) = 0 then
							Response.Write "<option VALUE=""0"">Alle"
						else
							Response.Write "<option VALUE=""0"" SELECTED>Alle"
						end if
						Do Until rsAreas.EOF
							lngAreaID = rsAreas("ProdOmradeID")
							If clng(lngAreaID) = clng(IntSelectedAreaID) Then
								strSelected = "selected"
							Else
								strSelected = ""
							End If
							Response.Write "<option VALUE=""" & lngAreaID & """ " & strSelected & ">" & rsAreas("Produktomrade")
							rsAreas.MoveNext
						Loop
						rsAreas.close
						Set rsAreas = Nothing
						%>
					</select>
					vis:
					<select NAME="dbxShowAll" onchange="endre_gruppe()">
						<%
						if (lngDisplayType=0) then
							strSelected = "selected"
						else
							strSelected = ""
						end if
						%>
						<option value="0" <%=strSelected%>>Valgte</option>
						<%
						if (lngDisplayType=1) then
							strSelected = "selected"
						else
							strSelected = ""
						end if
						%>
						<option value="1" <%=strSelected%>>Alle</option>
					</select>
					(Kompetanse sist oppdatert: <strong><% =strDato %></strong>)
					<br><br>
					<span class="menuInside"" title="Lagre informasjonen"><a href="#" onClick="document.all.KompForm.submit();"><img src="images/icon_save.gif" width="18" height="15" alt="" border="0" align="absmiddle">Lagre</a></span>
					<div class="listing">
							<table cellpadding="0" cellspacing="1">
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
		   							<th class="center">Nivå</th>
		   							<th class="center">Bruker niv&aring;<br><span class="normal">Enkelt|Erfaren|Avansert|Ekspert</span></th>
									<th class="center">kommentar</th>
									<th class="center">web?</th>
								</tr>
								<%
								if lngDisplayType = 1 then
									StrSQL = "[GetAllQualificationsListForConsultant] " & LngVikarid & "," & IntSelectedAreaID
								else
									StrSQL = "[GetQualificationListForConsultant] " & LngVikarid & "," & IntSelectedAreaID
								end if

								set rsKompetanse = GetFirehoseRS(strSQL, objCon)	

								' Get Qualification levels..
								set rsLevel = GetDynamicRS("select K_LevelID, Klevel from H_KOMP_LEVEL", objCon)	
							
								while (not rsKompetanse.EOF)
									intTittelID = clng(rsKompetanse.fields("K_TittelID").value)
									StrChecked = cstr(rsKompetanse.fields("besitter").value)
									%>
									<tr>
										<td><input onChange="endre_data('check','<%=intTittelID%>')" onClick="Toggle('check','<%=intTittelID%>')" id="TittelID<%=intTittelID%>" name="<%=intTittelID%>TittelID" type="checkbox" class="checkbox" <%=strchecked%>></td>
										<td><%=rsKompetanse.fields("ktittel").value%>&nbsp;</td>
									<%
									if IntSelectedAreaID = 0 then
										Response.Write "<td>" & rsKompetanse.fields("ProduktOmrade").value & "&nbsp;" & "</td>" & chr(13)
									end if
									Response.Write "<td><select name=""" & intTittelID & "dbxLevel"" onchange=""endre_data()""><option VALUE=""0"">"

									rsLevel.movefirst
									Do Until rsLevel.EOF
										If rsLevel("K_LevelID") = rsKompetanse.fields("K_LevelID").value Then
											strLevel = rsLevel("K_LevelID")
											strSelected = "SELECTED"
										Else
											strLevel = rsLevel("K_LevelID")
											strSelected = ""
										End If
										Response.Write "<option VALUE=""" & strLevel & """ " & strSelected & ">" & rsLevel("KLevel") & vbCrLf
										rsLevel.MoveNext
									Loop
									%>
									</select></td>
									<%
									if isnull(rsKompetanse.fields("rangering").value) then
										intRangering = 0
									else
										intRangering = cint(rsKompetanse.fields("rangering").value)
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
									<td class="center">
										<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="rdorangering<%=intTittelID%>" name="<%=intTittelID%>rdorangering" value="3" <%=strSelected3%>>
										<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="rdorangering<%=intTittelID%>" name="<%=intTittelID%>rdorangering" value="4" <%=strSelected4%>>
										<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="rdorangering<%=intTittelID%>" name="<%=intTittelID%>rdorangering" value="5" <%=strSelected5%>>
										<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="rdorangering<%=intTittelID%>" name="<%=intTittelID%>rdorangering" value="6" <%=strSelected6%>>
									</td>
									<td>
										<input onchange="endre_data()" size="30" maxlength="256" name="<%=intTittelID%>tbxKommentar" value="<%=rsKompetanse.fields("kommentar").value%>">
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
					</div>
					<span class="menuInside" title="Lagre informasjonen"><a href="#" onClick="document.all.KompForm.submit();"><img src="images/icon_save.gif" width="18" height="15" alt="" border="0" align="absmiddle">Lagre</a></span>
				</div>
			</form>
		</div>
	</body>
</html>
<%
CloseConnection(objCon)
set objCon = nothing
%>

