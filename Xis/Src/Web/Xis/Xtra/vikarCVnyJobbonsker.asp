<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	'File purpose:		Shows all available jobwishes in a table, and lets the
	'					the user select/deselect and write a commentary for it.
	'					The data is processed in vikarjobbonskelistedb.asp

	dim StrSQL				'Holds temporary SQL
	dim StrVikarnavn		'holds Name of consultant
	dim rsVikar				'as adodb.recordset
	dim rsProfession		'as adodb.recordset
	dim rsAreas				'as adodb.recordset
	dim intTittelID			'Holds id of current Jobwish in recordset loop
	dim IntSelectedAreaID	'Holds id of current Jobwish Area
	dim StrChecked			'Is Current value checked "" or "CHECKED"
	dim strSelected			'Is Current value selected "" or "selected"
	dim IntAreaID			'Holds id of current Jobwish area in recordset loop
	dim objConn				'objConnection to database
	dim intErfaringNiva
	dim strSelected1
	dim strSelected2
	dim strSelected3
	dim intUtdannelseNiva
	dim strSelected4
	dim strSelected5
	dim strSelected6
	dim lngDisplayType
	dim strDato
	dim lngAreaID

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
		AddErrorMessage("Error in Parameter. VikarID has no value!")
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

	' initialize and objConnect to database
	Set objConn = GetConnection(GetConnectionstring(XIS, ""))		

    ' Get consultant name and store for later use
    strSQL = "SELECT Navn = fornavn + ' ' + etternavn from vikar where vikarid=" & LngVikarid
	set rsVikar = GetFirehoseRS(strSQL, objConn)	
	StrVikarnavn = rsVikar("navn")
	rsVikar.close
    Set rsVikar = Nothing

	strSQL = "select [Kompetansedato] from [VIKAR] where [VikarID] = " & LngVikarid
	set rsVikar = GetFirehoseRS(strSQL, objConn)	
	If hasRows(rsVikar) = true Then
		strDato = rsVikar("Kompetansedato")
	End If
	rsVikar.close
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
		<title>Velg fagkompetanse</title>
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
						eval("document.KompForm.rdoErfaringniva"+strKompID+"[0].checked=false");
						eval("document.KompForm.rdoErfaringniva"+strKompID+"[1].checked=false");
						eval("document.KompForm.rdoErfaringniva"+strKompID+"[2].checked=false");
						eval("document.KompForm.rdoUtdannelseNiva"+strKompID+"[0].checked=false");
						eval("document.KompForm.rdoUtdannelseNiva"+strKompID+"[1].checked=false");
						eval("document.KompForm.rdoUtdannelseNiva"+strKompID+"[2].checked=false");
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
			<form Name="KompForm" Action="VikarJobbOnskelisteDB.asp" METHOD="POST">
				<div class="contentHead1">
					<h1>CV redigering - <%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & lngVikarID, StrVikarnavn, "Vis vikar " & StrVikarnavn )%></h1>
					<input NAME="VikarID" TYPE="HIDDEN" VALUE="<%=LngVikarid%>">
					<input NAME="dbxEndret" TYPE="HIDDEN" VALUE="0">
					<input NAME="dbxOldArea" TYPE="HIDDEN" VALUE="<%=IntSelectedAreaID%>">
					<input NAME="dbxJobwishSource" TYPE="HIDDEN" VALUE="vikarCVnyJobbOnsker.asp">
					<div class="contentMenu">
						<table cellpadding="0" cellspacing="0" width="96%">
							<tr>
								<td><!--#include file="vikarCVnyToolbar.asp"--></td>
								<td class="right"><!--#include file="Includes/contentToolsMenu.asp"--></td>
							</tr>
						</table>
					</div>
				</div>
				<div class="contentMenu2">
					<span class="menu2" id="1" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyPersonalia.asp?VikarID=<%=LngVikarID%>">Personalia</a></span>
					<span class="menu2 active" id="2"><strong>Fagkompetanse</strong></span>
					<span class="menu2" id="3" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyKvalifikasjoner.asp?VikarID=<%=LngVikarID%>">Produktkompetanse</a></span>
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
					Velg fagomr&aring;de:
					<select NAME="dbxArea" onchange="endre_gruppe()">
						<%
						'Get all availble profession areas
						strSQL = "SELECT * FROM [H_KOMP_FAGOMRADE]"
						set rsAreas = GetFirehoseRS(strSQL, objConn)	

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
							Response.Write "<option value='" & lngAreaID & "' " & strSelected & ">" & rsAreas("Fagomrade") & "</Option>"
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
						<%if (lngDisplayType = 1) then
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
						<table cellpadding="0" cellspacing="1">
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
								StrSQL = "[GetAllProfessionsListForConsultant] " & lngVikarID & "," & IntSelectedAreaID
							else
								StrSQL = "[GetProfessionListForConsultant] " & LngVikarid & "," & IntSelectedAreaID
							end if
							set rsProfession = GetFirehoseRS(strSQL, objConn)
					
							Do Until rsProfession.EOF
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
									intErfaringNiva = clng(rsProfession.fields("Relevant_WorkExperience").value)
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
								<td class="center">
									<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="rdoErfaringniva<%=intTittelID%>" name="<%=intTittelID%>rdoErfaringniva" value="1" <%=strSelected1%>>
									<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="rdoErfaringniva<%=intTittelID%>" name="<%=intTittelID%>rdoErfaringniva" value="2" <%=strSelected2%>>
									<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="rdoErfaringniva<%=intTittelID%>" name="<%=intTittelID%>rdoErfaringniva" value="3" <%=strSelected3%>>
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
								<td class="center">
									<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="rdoUtdannelseNiva<%=intTittelID%>" name="<%=intTittelID%>rdoUtdannelseNiva" value="1" <%=strSelected4%>>
									<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="rdoUtdannelseNiva<%=intTittelID%>" name="<%=intTittelID%>rdoUtdannelseNiva" value="2" <%=strSelected5%>>
									<input type="radio" class="radio" onclick="endre_data();Toggle('rdo','<%=intTittelID%>')" id="rdoUtdannelseNiva<%=intTittelID%>" name="<%=intTittelID%>rdoUtdannelseNiva" value="3" <%=strSelected6%>>
								</td>
								<td><input size="30" onchange="endre_data()" maxlength="256" name="<%=intTittelID%>tbxKommentar" value="<%=rsProfession.fields("kommentar").value%>"></td>
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
							rsProfession.close
							Set rsProfession = Nothing
							%>
							</tr>
						</table>
					</div>
					<span class="menuInside"" title="Lagre informasjonen"><a href="#" onClick="document.all.KompForm.submit();"><img src="images/icon_save.gif" width="18" height="15" alt="" border="0" align="absmiddle">Lagre</a></span>				
				</div>
			</form>
		</div>
	</body>
</html>
<%
CloseConnection(objConn)
set objConn = nothing
%>
