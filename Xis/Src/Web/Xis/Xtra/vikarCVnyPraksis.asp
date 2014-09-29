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

	dim exp
	dim allExp
	dim expKomm
	dim blnLocked
	dim beskrivelse
	dim exp_id
	dim feilmelding
	dim dette_aar
	dim strAction
	dim praFraMnd
	dim praFraAar
	dim praTilMnd
	dim praTilAar
	dim praArbgiver
	dim praTittel
	dim praKomm

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
	if Request.Form("data") <> "" then
		cvtype = Request.Form("data")
	end if

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

	if cons.CV.DataValues.Count = 0 then
		cons.CV.Save
	end if

	blnLocked = cv.islocked

	set allExp = cv.Experiences
	set exp	= Server.CreateObject("XtraWeb.Experience")
	set exp.Owner = cv

	exp_id = 0

	if Request.QueryString("expid").Count > 0 then
		exp_id = Request.QueryString("expid")
		set exp = cv.Experiences("ID" & exp_id)
		praKomm  = exp.DataValues("Description")
	end if

	if Request.Form("rerun") = "1" then
		strAction  = lcase(trim(request("pbnDataAction")))

		' Praksis...
		if (strAction = "lagre") then

			exp_id = clng(Request.Form("objId"))
			if exp_id = null then
				exp_id = 0
			end if
			if exp_id > 0 then
				set exp = cv.Experiences("ID" & exp_id)
			end if
			praFraMnd	= CStr(Request.Form("praFraMnd"))
			praFraAar	= CStr(Request.Form("praFraAar"))
			praTilMnd	= CStr(Request.Form("praTilMnd"))
			praTilAar	= CStr(Request.Form("praTilAar"))
			praArbgiver	= CStr(Request.Form("praArbeidsgiver"))
			praTittel	= CStr(Request.Form("praTittel"))
			praKomm		= CStr(Request.Form("Editor1"))

			if praFraMnd <> "" and praFraAar <> "" then
				if CInt(praFraMnd) < 1 or CInt(praFraMnd) > 12 or CInt(praFraAar) < 0 or CInt(praFraAar) > dette_aar then
					feilmelding = feilmelding & "Ugyldig måned eller årstall i feltet Fra!<br>"
				end if
			end if

			praFraAar = replace(praFraAar,",","")
			praFraAar = replace(praFraAar,".","")
			praFraAar = replace(praFraAar," ","")
			
			if ((not isnumeric(trim(praFraAar))) or (len(praFraAar)<4)) then
				feilmelding = feilmelding & "Ugyldig årstall i feltet Fra, må være på formatet ""ÅÅÅÅ""!<br>"
			end if

			if praTilMnd = "" and praTilAar = "" then
			'Do nothing
			elseif CInt(praTilMnd) < 1 or CInt(praTilMnd) > 12 or CInt(praTilAar) < 0 or CInt(praTilAar) > dette_aar then
				feilmelding = feilmelding & "Ugyldig måned eller årstall i feltet Til!<br>"
			end if

			if (praTilAar <> "") then
				praTilAar = replace(praTilAar,",","")
				praTilAar = replace(praTilAar,".","")
				praTilAar = replace(praTilAar," ","")
				if ((not isnumeric(trim(praTilAar))) or (len(praTilAar)<4)) then
					feilmelding = feilmelding & "Ugyldig årstall i feltet Til, må være på formatet ""ÅÅÅÅ""!<br>"
				end if
			end if


			if (len(trim(praFraMnd)) > 0 and len(trim(praFraAar)) > 0 and len(trim(praArbgiver)) > 0 and len(trim(praTittel)) > 0 and len(trim(feilmelding)) = 0) then
				exp.DataValues("FromMonth") = praFraMnd
				exp.DataValues("FromYear")	= praFraAar
				if praTilMnd <> "" and praTilAar <> "" then
					exp.DataValues("ToMonth") = praTilMnd
					exp.DataValues("ToYear")  = praTilAar
				else
					exp.DataValues("ToMonth") = empty
					exp.DataValues("ToYear")  = empty
				end if
				exp.DataValues("Place")		= praArbgiver
				exp.DataValues("Title")		= praTittel
				exp.DataValues("Description") = praKomm
				if not exp.Save then
					Response.Write "Ikke lagret!<br>"
				end if
				set cv		= nothing
				cons.CV.cleanup
				cons.cleanup
				set cons	= nothing

				Response.Redirect("vikarCVnyPraksis.asp?type=5&VikarID=" & lngVikarID)
			else
				response.write "Feilmelding:" & feilmelding & "!!!<br>"
				feilmelding = feilmelding & "Alle opplysninger bortsett fra Kommentarer er obligatoriske!<br>"
				if left(praTilMnd,1) = "0" then
					praTilMnd = mid(praTilMnd,2,1)
				end if
				if left(praFraMnd,1) = "0" then
					praFraMnd = mid(praFraMnd,2,1)
				end if
			end if
		end if

		if (strAction = "slett") then
			exp_id = clng(Request.Form("objId"))
			if exp_id <> 0 then
				set exp = cv.Experiences("ID" & exp_id)
				if not exp.Delete then
					Response.Write "Ikke slettet!<br>"
				end if
				set cv		= nothing
				cons.CV.cleanup
				cons.cleanup
				set cons	= nothing
				Response.Redirect("vikarCVnyPraksis.asp?type=5&VikarID=" & lngVikarID)
			end if
		end if

	end if
	praKomm = "<link type='text/css' rel='stylesheet' href='http://" & Application("HTTPadress") & "/xtra/css/CV.css' title='default style'>" & praKomm
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
<body style="PADDING-RIGHT: 0px; PADDING-LEFT: 0px;">
	<div class="pageContainer" id="pageContainer">
		<div class="contentHead1">
			<H1 style="font-size:160%; font-weight:bold; font-color:#000000;"> CV redigering - <%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & lngVikarID, Cons.DataValues("Fornavn") & " " & Cons.DataValues("Etternavn"), "Vis vikar " & Cons.DataValues("Fornavn") & " " & Cons.DataValues("Etternavn") )%></h1>
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
			<span class="menu2 active" id="7"><strong>Yrkeserfaring</strong></span>
			<span class="menu2" id="8" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyAndreOppl.asp?VikarID=<%=lngVikarID%>">Kjernekompetanse</a></span>
			<span class="menu2" id="9" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyReferanser.asp?VikarID=<%=lngVikarID%>">Referanser</a></span>
			<span class="menu2" id="10" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVvis.asp?VikarID=<%=lngVikarID%>">&nbsp;<img src="/xtra/images/icon_new2.gif" width="14" height="14" alt="" border="0" align="absmiddle">&nbsp;Generere CV</a></span>
		</div>
		<div class="content">
		<form name="CVForm" action="vikarCVnyPraksis.asp?VikarID=<%=lngVikarID%>" method="post">
			<input type="hidden" id="rerun" name="rerun" value="1">
			<input Type="hidden" Name="pbnDataAction" Value="lagre">
			<input type="hidden" id="data" name="data">
			<input type="hidden" id="objId" name="objId">
			<table width="96%" class="layout" cellpadding="0" cellspacing="0">
				<col width="80%">
				<col width="20%">
			<tr>
				<td>
				<table width="100%" cellpadding='0' cellspacing='0'>
					<tr>
						<td width="100">Fra (mm/åååå):</td>
						<td width="30"><input type="text" name="praFraMnd" size="2" maxlength="2"
							<%if exp_id <> 0 then%> value="<%if exp.DataValues("FromMonth") < 10 then%>0<%end if%> <%=exp.DataValues("FromMonth")%>"
							<%else%> value="<%if praFraMnd <> "" then
								if ((CInt(praFraMnd) < 10) and (CInt(praFraMnd) > 0))  then
									%>0<%
								end if
							end if%> <%=praFraMnd%>" <%end if%> class="mandatory"></td>
						<td width="10">&nbsp;&nbsp;/&nbsp;&nbsp;</td>
						<td width="60"><input type="text" name="praFraAar" size="4" maxlength="4"
							<%if exp_id <> 0 then%> value="<%=exp.DataValues("FromYear")%>"
							<%else%> value="<%=praFraAar%>" <%end if%> class="mandatory">
						</td>
						<td width="100">Til (mm/åååå):</td>
						<td width="30"><input type="text" name="praTilMnd" size="2" maxlength="2"
							<%if exp_id <> 0 then%> value="<%if exp.DataValues("ToMonth") < 10 then%>0<%end if%> <%=exp.DataValues("ToMonth")%>"
							<%else%> value="<%if praTilMnd <> "" then
								 if ((CInt(praTilMnd) < 10) and (CInt(praTilMnd) > 0)) then
									%>0<%
								end if
							end if%> <%=praTilMnd%>" <%end if%> class="mandatory">
						</td>
						<td width="10">&nbsp;&nbsp;/&nbsp;&nbsp;</td>
						<td width="60"><input type="text" name="praTilAar" size="4" maxlength="4"
							<%if exp_id <> 0 then%> value="<%=exp.DataValues("ToYear")%>"
							<%else%> value="<%=praTilAar%>" <%end if%> class="mandatory">
						</td>
						<td width="200">&nbsp;</td>
					</tr>
		<tr>
			<td width="100">Arbeidsgiver:</td>
			<td colspan="8"><input type="text" name="praArbeidsgiver" size="40" maxlength="50"
				<%if exp_id <> 0 then%> value="<%=exp.DataValues("Place")%>"
				<%else%> value="<%=praArbeidsgiver%>" <%end if%> class="mandatory">
			</td>
		</tr>
		<tr>
			<td width="100">Tittel:</td>
			<td colspan="8"><input type="text" name="praTittel" size="40" maxlength="50"
				<%if exp_id <> 0 then%> value="<%=exp.DataValues("Title")%>"
				<%else%> value="<%=praTittel%>" <%end if%> class="mandatory">
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
			    editor.Text = praKomm			    
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
	<span class="menuInside" style="margin-left:104px;" title="Lagre informasjonen"><a href="#" onClick="javascript:lagre_data(<%=exp_id%>);"><img src="images/icon_save.gif" width="18" height="15" alt="" border="0" align="absmiddle">Lagre</a></span>
	<div class="listing">
	<table cellpadding='0' cellspacing='1'>
		<tr>
			<th colspan="7">Registrert Yrkeserfaring</th>
		</tr>
		<tr>
			<th>Fra</th>
			<th>Til</th>
			<th>Arbeidsgiver</th>
			<th>Tittel</th>
			<th>Kommentarer</th>
			<th class="center">Endre</th>
			<th class="center">Slette</th>
		</tr>
		<%
		for i = 1 to allexp.Count
			set exp = allexp.Item(i)
			id = exp.DataValues("DataId")

			StrFromMonth = exp.DataValues("FromMonth").value
			strFromYear = exp.DataValues("FromYear").value
			StrToMonth 	= exp.DataValues("ToMonth").value
			strToYear 	= exp.DataValues("ToYear").value

			if (StrFromMonth < 10) then
				StrFromMonth = "0" & StrFromMonth
			end if
			StrFromPeriod = StrFromMonth & "/" & strFromYear

			if (StrToMonth < 10) and (StrToMonth >= 1) then
				StrToMonth = "0" & StrToMonth
			end if

			if (StrToMonth=0) or (strToYear=0) or isNull(StrToMonth) or isNull(strToYear) then
				StrToPeriod = "d.d."
			else
				StrToPeriod = StrToMonth & "/" & strToYear
			end if
			%>
			<tr>
				<td><%=StrFromPeriod%></td>
				<td><%=StrToPeriod%></td>
				<td><%=exp.DataValues("Place")%></td>
				<td><%=exp.DataValues("Title")%></td>
				<td><%=exp.DataValues("Description")%>&nbsp;</td>
				<td class="center"><a href="vikarCVnyPraksis.asp?type=4&VikarID=<%=lngVikarID%>&expid=<%=id%>"><img src="/xtra/images/icon_changeInfo.gif" width="18" height="15" alt="Endre denne oppføringen" border="0"></a></td>
				<td class="center"><a href="javascript:slett_data(<%=id%>)"><img src="/xtra/images/icon_delete.gif" width="14" height="14" alt="Slette denne oppføringen" border="0"></a></td>
			</tr>
			<%
		next
		%>
	</table>
	</div>

</form>
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
