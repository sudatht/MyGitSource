<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!-- #include file = "cuteeditor_files/include_CuteEditor.asp" --> 
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim lngAktivitetID
	dim lngFirmaID
	dim lngVikarID
	dim rsFirma
	dim lngbrukerID
	dim strProfil
	dim lngOppdragID
	dim strNavn
	dim strNotat
	dim RecordsFound : RecordsFound = false
	dim source

	'Hotlist menu items variables
	dim strAddToHotlistLink
	dim strHotlistType
	dim blnShowHotList
	dim rsTest

	'Consultant menu variables
	dim strClass
	dim strJSEvents

	// Text Editor Code - CuteEditor	
	Dim editor	
	Set editor = New CuteEditor
	editor.ID = "txtNotat"
	editor.AutoConfigure = "Simple"	

	strProfil = Session("Profil")
	lngbrukerID = Session("BrukerID")

	' Move input values to local variables
	lngAktivitetID = Request("AktivitetID")
	If Request("FirmaID") <> "" Then
		lngFirmaID = CLng( Request("FirmaID") )
	End If

	If Request("VikarID") <> "" Then
		lngVikarID = CLng( Request("VikarID") )
	End If

	If Request("OppdragID") <> "" Then
		lngOppdragID = CLng( Request("OppdragID") )
	End If

	If  len(trim(Request("source"))) = 0 Then
		source = ""
	Else
		source = Request("source")
	End If

	' Open database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

	If (lngVikarID > 0) Then
		' Get vikarname
		set rsVikar = GetFirehoseRS("SELECT Navn=Fornavn+' '+Etternavn FROM Vikar WHERE VikarID = " & lngVikarID, Conn)
		strNavn = rsVikar( "Navn" )
		rsVikar.Close
		set rsVikar = Nothing

		'Sjekk for å forhindre dobbelreg i Hotlist
		
		set rsTest = GetFirehoseRS("SELECT * from HOTLIST Where status=3 And BrukerID=" & lngbrukerID & " And navnID=" & lngVikarID, Conn)
		if (HasRows(rsTest) = false) then
			blnShowHotList = true
			strAddToHotlistLink = "addHotlist.asp?kode=3&vikarNavn=" & server.URLEncode(Navn) & "&vikarNr=" & lngVikarID
			strHotlistType = "vikar"
		else
			blnShowHotList = true
			strAddToHotlistLink = ""
			strHotlistType = "vikar"
		end if
		rsTest.close
		set rsTest = nothing
	End If

	If lngOppdragID <> 0 Then
		' Get vikarname
		set rsOppdrag = GetFirehoseRS("SELECT Beskrivelse FROM Oppdrag WHERE OppdragID = " & lngOppdragID, Conn)
		Heading = "Oppdrag '" & rsOppdrag( "beskrivelse" ) & "'"
		rsOppdrag.Close
		set rsOppdrag = Nothing
	End If

	' Aktivitet exist ?	
	If lngAktivitetID <> "" Then

		' Build sql-statement
		strSql = "SELECT [A].[Aktivitetdato], [B].[Navn], [AT].[AktivitetTypeID], [AT].[AktivitetType], [A].[FirmaID], [A].[VikarID], [A].[OppdragID], [A].[RegistrertAv], [A].[Notat] "&_
		" FROM [AKTIVITET] AS [A] " &_
		" INNER JOIN [H_AKTIVITET_TYPE] AS [AT] ON [A].[AktivitetTypeID] = [AT].[AktivitetTypeID] " & _
		" INNER JOIN [Bruker] AS [B] ON [A].[RegistrertAvID] = [B].[MedarbID] " & _
		" WHERE [A].[AktivitetID] = " & lngAktivitetID

		' Get all records
		set rsAktivitet = GetFirehoseRS(strSql, Conn)

		'Records found ?
		If (HasRows(rsAktivitet)) Then
			RecordsFound = true
		End If

		AktivitetDato	= rsAktivitet( "Aktivitetdato" )
		TypeID			= rsAktivitet( "AktivitetTypeID" )
		strNotat		= Trim( rsAktivitet( "Notat" ) )	
		

		' Close and release recordset
		rsAktivitet.Close
		Set rsAktivitet = Nothing
		' Set heading
		Heading = "Endre aktivitet"
	Else
		' Set heading
		Heading = "Ny aktivitet"
		AktivitetDato = Date()
		strNotat = ""
	End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"  "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">		
		<title><%=strHeading %></title>
		<script type="text/javascript" src="/xtra/Js/javascript.js"></script>
		<script type="text/javascript" src="/xtra/js/CVFunctions.js"></script>
		<script type="text/javascript" src="/xtra/js/contentMenu.js"></script>		
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1><%=Heading%></h1>
				<div class="contentMenu">
					<%
					If (lngVikarID > 0) Then
						%>
						<table cellpadding="0" cellspacing="0" width="96%">
							<tr>
								<td>
									<table cellpadding="0" cellspacing="2">
										<tr>
											<td class="menu disabled" id="menu1">
												<img src="/xtra/images/icon_save.gif" width="18" height="15" alt="" align="absmiddle">Lagre
											</td>
											<td class="menu" id="menu2" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
												<a href="/xtra/VikarVis.asp?vikarid=<%=lngVikarID%>" title="Vis vikar">
												<img src="/xtra/images/icon_consultant.gif" width="18" height="15" alt="" align="absmiddle">Vis</a>
											</td>
											<%
											If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = true) Then
												%>
												<td class="menu" id="menu3" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
													<form ACTION="VikarNy.asp?VikarID=<%=lngVikarID%>" METHOD="POST" id="frmConsultantChange">
														<input NAME="pbnDataAction" TYPE="hidden" VALUE="Endre kons.opplysninger" ID="Hidden1">
														<a href="javascript:document.all.frmConsultantChange.submit();" title="Endre vikar">
														<img src="/xtra/images/icon_changeInfo.gif" width="18" height="15" alt="" align="absmiddle">Endre</a>
													</form>
												</td>
												<td class="menu" id="menu4" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);"><strong>CV</strong>&nbsp;<select id="cboCVChoice" onChange="javascript:Vis_CV(<%=lngVikarID%>);" NAME="cboCVChoice"><option value="0"></option><option value="1">Se</option><option value="2">Endre</option><option value="3">Presentere</option></select></td>
												
												<%
											End If
											%>
											<td class="menu" id="menu6" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
												<form ACTION="vikar-Kunder.asp?VikarID=<%=lngVikarID%>" METHOD="POST" id="frmConsultantFormerClients">
													<a href="javascript:document.all.frmConsultantFormerClients.submit();" title="Vis tidligere oppdragsgivere"><img src="/xtra/images/icon_tidl-kunder.gif" alt="" width="18" height="15" border="0" align="absmiddle">Tidligere oppdragsgivere</a>
												</form>
											</td>
											<td class="menu" id="menu7" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
												<form ACTION="AktivitetVikar.asp?VikarID=<%=lngVikarID%>" METHOD="POST" id="frmConsultantActivities">
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
						<%
					end if
					%>
				</div>
			</div>
			<div class="content">
				<form Name="VIKAR" ACTION="aktivitetdb.asp" METHOD="POST" name="frmAktivitet" id="frmAktivitet">
					<input NAME="source" TYPE="HIDDEN" VALUE="<%=source%>" ID="Hidden6">
					<input NAME="tbxAktivitetID" TYPE="HIDDEN" VALUE="<%=lngAktivitetID%>" ID="Hidden2">
					<input NAME="tbxVikarID" TYPE="HIDDEN" VALUE="<%=lngVikarID%>" ID="Hidden3">
					<input NAME="tbxFirmaID" TYPE="HIDDEN" VALUE="<%=lngFirmaID%>" ID="Hidden4">
					<input NAME="tbxOppdragID" TYPE="HIDDEN" VALUE="<%=lngOppdragID%>" ID="Hidden5">
					<input NAME="pbnDataAction" TYPE="HIDDEN" VALUE="" ID="pbnDataAction">
					<div class="listing">
						<table>
							<tr>
								<th>Dato:</th>
								<th>Type:</th>
							</tr>
							<tr>
								<td><input NAME="tbxAktivitetDato" TYPE="TEXT" SIZE="8" MAXLENGTH="10" VALUE="<%=AktivitetDato%>" ONBLUR="dateCheck(this.form, this.name)" ID="Text1"></td>
								<td>
									<select NAME="dbxType">
										<%
										'aktiviteter
										set rsType = GetFirehoseRS("select AktivitetTypeID, AktivitetType from H_AKTIVITET_TYPE where AutoRegister=0 order by 2", Conn)
										Do Until rsType.EOF
											If rsType("AktivitetTypeID") = TypeID Then sel = " SELECTED" Else sel = ""%>
												<option VALUE="<% =rsType("AktivitetTypeID") %>" <%=sel %>><%=rsType("AktivitetType") %>
												<% 
											rsType.MoveNext
										Loop
										rsType.Close
										Set rsType = Nothing
										%>
									</select>
								</td>
							</tr>
						</table>
						
						<%						
						
						editor.Text = strNotat
						editor.Draw()
						
						
						%>
						
						<br/>
						<br/>
						<span class="menuInside" style="margin-left:10px;" title="Lagre aktivitet"><a href="#" onClick="document.all.frmAktivitet.submit();"><img src="images/icon_save.gif" width="18" height="15" alt="" border="0" align="absmiddle">Lagre</a></span>
						<%
						' Show appropiate buttons depending on AktivitetID
						If lngAktivitetID > 0 Then
							%>
							<span class="menuInside" style="margin-left:10px;" title="Slette aktivitet"><a href="#" onClick="document.frmAktivitet.pbnDataAction.value='Slette aktivitet';document.all.frmAktivitet.submit();"><img src="images/icon_delete.gif" alt="" border="0" align="absmiddle">&nbsp;Slette</a></span>
							<%
						End If
						%>
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