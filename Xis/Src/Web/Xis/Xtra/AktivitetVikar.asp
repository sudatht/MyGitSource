<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim strSQL
	dim StrSQLWhere
	dim rsAktivitet
	dim lngVikarID
	dim strSQLType
	dim lngActivityType
	dim strRegistereredBy
	dim strSQLRegisteredBy
	dim strFraDato
	dim strTilDato
	dim lngBrukerID
	dim strNavn
	dim strKunde

	'Hotlist menu items variables
	dim strAddToHotlistLink
	dim strHotlistType
	dim blnShowHotList

	'Consultant menu variables
	dim strClass
	dim strJSEvents

	dim strShowAuto
	
	lngbrukerID = Session("BrukerID")

	strShowAuto = Request.Querystring( "chkShowAutoRegister" )
	
	' Check input values

	' Do we have VikarID ?
	If Request( "VikarID" ) = "" Then
		AddErrorMessage("Feil: Mangler parameter vikarID. Noter navn på vikar og kontakt systemansvarlig.")
		call RenderErrorMessage()
	Else
		lngVikarID = Request( "VikarID" )
	End If

	' Get the people that has registered activities on this oppdrag
	If (Request.Querystring( "dbxRegisteredBy" ) = "" OR Request.Querystring( "dbxRegisteredBy" ) = "Alle") Then
		strRegistereredBy = ""
		strSQLRegisteredBy = ""
	Else
		strRegistereredBy = cstr(Request.Querystring( "dbxRegisteredBy" ))
		strSQLRegisteredBy = " AND [A].[registrertavID]  = " & strRegistereredBy & " "
	End If

	' is there an activity selected
	If (Request.Querystring( "dbxActivityType" ) = "") OR (Request.Querystring( "dbxActivityType" ) = 0) Then
		lngActivityType = 0
		strSQLType = ""
	Else
		lngActivityType = clng(Request.Querystring( "dbxActivityType" ))
		strSQLType = " AND [A].[AktivitetTypeID] =" & lngActivityType & " "
	End If

	' Add selection on Fradato ?
	If Request.querystring("tbxFradato") <> "" Then
		StrSQLWhere = strSql & " AND [A].[AktivitetDato] >= " & DbDate( Request.querystring("tbxFradato") )
		strFraDato = Request.querystring("tbxFradato")
	End If

	'on/off auto registered activity list
	If strShowAuto <> "showAuto" Then
		StrSQLWhere = StrSQLWhere & " AND [A].[AutoRegistered] = 0"
	End If
	
	' Add selection on Tildato ?
	If Request.querystring("tbxTildato") <> "" Then
		' Add selection on tildato
		StrSQLWhere = StrSQLWhere & " AND [A].[AktivitetDato]  <= " & DbDate( Request.querystring("tbxTildato") )
		strTilDato = Request.querystring("tbxTildato")
	End If
	
	' Open connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

	' Get vikarname
	set rsVikar = GetFirehoseRS("SELECT [Fornavn] +' '+ [Etternavn] AS [Navn] FROM [VIKAR] WHERE [VikarID] = " & lngVikarID, Conn)
	strNavn = rsVikar( "Navn" )
	rsVikar.Close
	set rsVikar = Nothing

	'Sjekk for å forhindre dobbelreg i Hotlist
	set rsTest = GetFirehoseRS("SELECT * FROM [HOTLIST] WHERE [status] = 3 AND [BrukerID] = " & lngbrukerID & " AND [navnID] =" & lngVikarID, Conn)
	if (HasRows(rsTest) = true) then
		blnShowHotList = true
		strAddToHotlistLink = "addHotlist.asp?kode=3&vikarNavn=" & server.URLEncode(strNavn) & "&vikarNr=" & lngVikarID
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
		<meta http-equiv="Content-
		-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<title>Aktiviteter</title>
		<script language="javascript" src="Js/javascript.js" type="text/javascript"></script>
		<script type="text/javascript" src="/xtra/js/CVFunctions.js"></script>
		<script type="text/javascript" src="/xtra/js/contentMenu.js"></script>
		<script language='javascript' src='js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='js/navigation.js' id='navigationScripts'></script>			
		<script language="javaScript">
			function shortKey(e) 
			{
				var keyChar = String.fromCharCode(event.keyCode);
				var modKey  = event.ctrlKey;
				var modKey2 = event.shiftKey;

				//linker i submeny
				<% 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>		
					if (modKey && modKey2 && keyChar=="S")
					{
						parent.frames[funcFrameIndex].location=("/xtra/vikarSoek.asp");
					}
					<% 
				End If 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
					%>
					if (modKey && modKey2 && keyChar=="Y")
					{
						parent.frames[funcFrameIndex].location=("/xtra/VikarDuplikatSoek.asp");
					}
					<% 
				End If 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>
					if (modKey && modKey2 && keyChar=="W")
					{
						parent.frames[funcFrameIndex].location=("/xtra/jobb/SuspectList.asp");
					}
					<% 
				End If 
				%>
			}
			// her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;
		</script>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Aktiviteter</h1>
				<div class="contentMenu">
					<table cellpadding="0" cellspacing="0" width="96%" ID="Table1">
						<tr>
							<td>
								<table cellpadding="0" cellspacing="2" ID="Table2">
									<tr>
										<td class="menu disabled" id="menu1">
											<img src="/xtra/images/icon_save.gif" width="18" height="15" alt="" align="absmiddle">Lagre
										</td>
										<td class="menu" id="menu2" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
											<a href="/xtra/vikarvis.asp?vikarid=<%=lngVikarID%>" title="Vis vikar">
											<img src="/xtra/images/icon_consultant.gif" width="18" height="15" alt="" align="absmiddle">Vis</a>
										</td>
										<%
										If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = true) Then
											%>
											<td class="menu" id="menu3" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
												<form ACTION="vikarny.asp?VikarID=<%=lngVikarID %>" METHOD="POST" id="frmConsultantChange">
													<input NAME="pbnDataAction" TYPE="hidden" VALUE="Endre kons.opplysninger" ID="Hidden1">
													<a href="javascript:document.all.frmConsultantChange.submit();" title="Endre vikar">
													<img src="/xtra/images/icon_changeInfo.gif" width="18" height="15" alt="" align="absmiddle">Endre</a>
												</form>
											</td>
											<%
										End If
										%>
										<td class="menu" id="menu4" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);"><strong>CV</strong>&nbsp;<select id="cboCVChoice" onChange="javascript:Vis_CV(<%=lngVikarID%>);" NAME="cboCVChoice"><option value="0"></option><option value="1">Se</option><%If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then%><option value="2">Endre</option><%end if%><option value="3">Presentere</option></select></td>
										 
										<td class="menu" id="menu6" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
											<form ACTION="vikar-kunder.asp?VikarID=<%=lngVikarID%>" METHOD="POST" id="frmConsultantFormerClients">
												<a href="javascript:document.all.frmConsultantFormerClients.submit();" title="Vis tidligere oppdragsgivere"><img src="/xtra/images/icon_tidl-kunder.gif" alt="" width="18" height="15" border="0" align="absmiddle">Tidligere oppdragsgivere</a>
											</form>
										</td>
										<td class="menu disabled">
											<img src="/xtra/images/icon_activities.gif" alt="" width="18" height="15" border="0" align="absmiddle">Aktiviteter
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
			<div class="content">
			
					<table>
					<tr>
					<td style="width:150px;">
				
					<%
				If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = true) Then
					%>
					<form name='formTo' ACTION='AktivitetNy.asp?OppdragID=<%=OppdragID%>&FirmaID=<%=FirmaID%>&VikarID=<%=lngVikarID%>' method="post" ID="Form2">
						<table cellpadding='0' cellspacing='1' ID="Table3">
							<tr>
								<td>
									<INPUT NAME="source" Type="HIDDEN" Value="vikar" ID="Hidden3">
									<INPUT NAME="tbxOppdragID" Type="HIDDEN" Value="<%=OppdragID%>" ID="Hidden4">
									<INPUT NAME="pbnDataAction" TYPE="SUBMIT" VALUE='Ny aktivitet' ID="Submit2">					
								</td>
							</tr>
						</table>
						
					</form>
					<%
				end if%>
				
					<form name="formEn" ACTION="aktivitetvikar.asp" ID="Form1">
					
					</td>
					<td style="width:150px;"></td>
					<td style="width:100px;">Vis auto aktiviteter:</td>
					<td>			
					<%if (strShowAuto="showAuto") then
						%>
						<input class="checkbox" type="checkbox" id="chkShowAutoRegister" name="chkShowAutoRegister" value="showAuto" onclick="submit()" checked/>	
						<%
					else
						%>
						<input class="checkbox" type="checkbox" id="chkShowAutoRegister" name="chkShowAutoRegister" value="showAuto" onclick="submit()" />
						<%
					end if
					%><br/>
					<INPUT TYPE="HIDDEN" NAME="VikarID" VALUE="<%=lngVikarID%>" ID="Hidden2">
					<%
					strSql = "select medid from medarbeider where fornavn = 'Vikaren'"
					set rsDummy = GetFirehoseRS(strSql, Conn)
					if rsDummy.EOF then
						Dim abcd
						abcd = "test 1"
						
					Else 
						strDummyId = rsDummy( "medid" )
						
					End if
					rsDummy.Close
	                		set rsDummy = Nothing
	                		
					%>
					</td>
					</tr>
					<tr>
					<td style="width:150px;">
					Start dato: <input ID="lnk1" onFocus="focused(this.id)" NAME="tbxFradato" TYPE="TEXT" SIZE="10" MAXLENGTH="10" ONBLUR="dateCheck(this.form, this.name)" value="<%=strFraDato%>">
					</td>
					<td style="width:150px;">
					Slutt dato: <input ID="lnk2" onFocus="focused(this.id)" NAME="tbxTilDato" TYPE="TEXT" SIZE="10" MAXLENGTH="10" ONBLUR="dateCheck(this.form, this.name)" value="<%=strTilDato%>">
					</td>
					<td style="width:100px;">Aktivitetstype: </td>
					<td>
					
					<select NAME="dbxActivityType" ID="Select1">
						<option VALUE="0">Alle typer</option>
						<%
						strSql = "SELECT [AktivitetTypeID], [AktivitetType] FROM [H_AKTIVITET_TYPE] ORDER BY AutoRegister,2"
						set rsAktivitet = GetFirehoseRS(strSql, Conn)
						While not rsAktivitet.EOF
							If clng(rsAktivitet("AktivitetTypeID")) = clng(lngActivityType) Then
								strSelected = " SELECTED"
							Else
								strSelected = ""
							End If

							response.write "<OPTION VALUE='" & rsAktivitet("AktivitetTypeID") & "' " & strSelected & ">" & rsAktivitet("AktivitetType") & "</OPTION>"
							rsAktivitet.movenext
						wend
						rsAktivitet.close
						set rsAktivitet = nothing
						
						%>
					</select>
					
					</td>
					<td>
					
						Registrert av: 
					
						<select NAME="dbxRegisteredBy" ID="Select2">
							<option VALUE="Alle">Alle</option>
							<%
							
							'strSql ="SELECT DISTINCT [Bruker].[Navn], [Bruker].[MedarbID] FROM [Bruker] INNER JOIN [Aktivitet] ON [Aktivitet].[RegistrertAvID] = [Bruker].[MedarbID] WHERE [Aktivitet].[vikarid] = " & lngVikarID  & " ORDER BY [Bruker].[Navn]"
	                        strSql ="select DISTINCT [medarbeider].[fornavn] + ' ' + [medarbeider].[etternavn] as medName, [medarbeider].[medid] from [medarbeider] INNER JOIN [Aktivitet] ON [Aktivitet].[RegistrertAvID]= [medarbeider].[medid] WHERE [Aktivitet].[vikarid] = " & lngVikarID  & " ORDER BY medName" 
							set rsRegistrertBy = GetFirehoseRS(strSql, Conn)
							While not rsRegistrertBy.EOF
								If cstr(rsRegistrertBy("medid")) = cstr(strRegistereredBy) Then
									strSelected = " SELECTED"
								Else
									strSelected = ""
								End If
								If cstr(rsRegistrertBy("medid")) = cstr(strDummyId) Then
									strName = strNavn
								Else
									strName = rsRegistrertBy("medName")
								End If
											
								Response.Write "<OPTION VALUE='" & rsRegistrertBy("medid") & "' " & strSelected & ">" & strName & "</OPTION>"
								rsRegistrertBy.movenext
							wend
							rsRegistrertBy.close
							set rsRegistrertBy = nothing
							%>
						</select>
						<INPUT TYPE="SUBMIT" VALUE="Søk" ID="Submit1" NAME="Submit1">
						
					</td>
					</tr>
					</table>
				
					
					
				</form>
				<%
				' Build sql-statement
				
				strSql = "SELECT [A].[AktivitetID],[A].[registrertavid], [A].[Aktivitetdato], [A].[AutoRegistered], [AT].[AktivitetType], [F].[SOCuID],  isnull([F].CRMAccountGuid,'') AS CRMAccountGuid, [F].[Firma], [A].[FirmaID], [A].[VikarID], " & _
						" [A].[OppdragID], [V].[Fornavn] + ' ' + [V].[Etternavn] AS [Vikarnavn], CASE WHEN ([M].[fornavn]) IS NULL THEN [A].RegistrertAv ELSE ([M].[fornavn] + ' ' + [M].[etternavn]) END AS Navn, [A].[Notat] " & _
						" FROM [AKTIVITET] AS [A] " & _
						" INNER JOIN [H_AKTIVITET_TYPE] AS [AT] ON [AT].[AktivitetTypeID] = [A].[AktivitetTypeID] " & _
						" LEFT OUTER JOIN [Firma] AS [F] ON [A].[FirmaID] = [F].[FirmaID] " & _
						" LEFT OUTER JOIN [medarbeider] AS [M] ON [A].[RegistrertAvID] = [M].[medid] " & _
						" INNER JOIN [Vikar] AS [V] ON [A].[VikarID] = [V].[VikarID] " & _
 						" WHERE (([A].[VikarID] = " & lngVikarID & ") OR ([A].[VikarID] = " & lngVikarID & " AND [A].[OppdragID] > 0 )) " & _
						strSQLType & _
						strSQLRegisteredBy & _
						StrSQLWhere & _
						" ORDER BY [A].[AktivitetDato] DESC "

				' Get all records
				set rsAktivitet = GetFirehoseRS(strSql, Conn)

				' No records found ?
				if (HasRows(rsAktivitet) = true) then
					' Create table only when records found
					%>
					<div class="listing">
						<table cellpadding='0' cellspacing='1' ID="Table4">
							<tr>
								<th>Dato</th>
								<th class="nowrap" nowrap>Reg. av</th>
								<th>Type</th>
								<th>Vikar</th>
								<th>Kontakt</th>
								<th>Oppdrag</th>
								<th>Kommentar</th>
								<%
								If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = true) Then
									%>
									<th>Slette</th>
									<%
								end if
								%>
							</tr>
							<%
							Do Until (rsAktivitet.EOF)
								' Create row
								Response.Write "<tr>"
								Response.Write "<td>" & rsAktivitet( "AktivitetDato") & "</td>"
																
								If cstr(rsAktivitet("registrertavid")) = cstr(strDummyId) Then
									strName = strNavn
								Else
									strName = rsAktivitet( "navn")
								End If
								
								Response.Write "<td>" & strName & "</td>"
								'Response.Write "<td>" & rsAktivitet( "navn") & "</td>"
								If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = true) Then
									If (rsAktivitet("AutoRegistered") = 0) Then
										Response.Write "<td><A href='AktivitetNy.asp?AktivitetID=" & rsAktivitet( "AktivitetID" ) & "&source=vikar&FirmaID=" &rsAktivitet( "FirmaID") & "&VikarID=" & rsAktivitet( "VikarID") & "&OppdragID=" &rsAktivitet( "OppdragID") & "'>" & rsAktivitet( "AktivitetType") & "</a></td>"
									Else
										Response.Write "<td>" & rsAktivitet( "AktivitetType") & "</td>"
									End If
								Else
									Response.Write "<td>" & rsAktivitet( "AktivitetType") & "</td>"
								end if						
								
								' Show link only if value
								If  rsAktivitet( "VikarID" ) <> 0 Then
									Response.Write "<td class='nowrap'><A href='vikarVis.asp?VikarID=" & rsAktivitet( "VikarID" ) & "'>" & rsAktivitet( "Vikarnavn") & "</a></td>"
								Else
									Response.Write "<td>&nbsp;</td>"
								End If
								
								' Show link only if value
								linkurl = Application("CRMAccountLink") & rsAktivitet("CRMAccountGuid") & "%7d&pagetype=entityrecord"
								
								If  rsAktivitet( "SOCuID" ) <> 0 Then
									strKunde = "<a href=" & linkurl & " target='_blank'>" & rsAktivitet("Firma").Value & " </a>"									
								Else
									strKunde = rsAktivitet("Firma").Value
								End If
								
								Response.Write "<td>" & strKunde & "</td>"					
															

								' Show link only if value
								
								If  rsAktivitet( "OppdragID" ) <> 0 Then
									Response.Write "<td>" & CreateSOLink(SUPEROFFICE_XIS_TASK_URL, "", "WebUI/Oppdragview.aspx?OppdragID=" & rsAktivitet("oppdragID").Value, rsAktivitet( "OppdragID"), "Vis Oppdrag" ) & "</td>"
								Else
									Response.Write "<td>&nbsp;</td>"
								End If

								Response.Write "<td>" & rsAktivitet( "Notat" ) & "</td>"
								If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = true) Then
									If (rsAktivitet("AutoRegistered") = 0) Then
										Response.Write "<td class='center'><a href='aktivitetDB.asp?source=vikar&pbnDataAction=slette&tbxFirmaID=" & rsAktivitet( "FirmaID" ) & "&tbxVikarID=" & rsAktivitet( "VikarID" ) & "&tbxOppdragID=" & rsAktivitet( "OppdragID") & "&tbxAktivitetID=" & rsAktivitet( "AktivitetID" ) & "&chkShowAutoRegister=" & strShowAuto & "'><img src='/xtra/images/icon_delete.gif' alt='Slett' width='12' height='12' border='0'></a></td>"
									Else
										Response.Write "<td class='center'>&nbsp;</td>"
									End If
								End If
								Response.Write "</tr>"
								rsAktivitet.MoveNext
							Loop
							' Close recordset
							rsAktivitet.Close
							set rsAktivitet = Nothing
							%>
						</table>
					</div>
					<%
				end if
				%>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(oConn)
set oConn = nothing
%>