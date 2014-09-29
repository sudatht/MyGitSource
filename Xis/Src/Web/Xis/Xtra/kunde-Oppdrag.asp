<%@ Language="VbScript" %>
<%option explicit
Response.Expires = 0
%>
<!--#INCLUDE FILE="includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="includes/xis.rights.inc"-->
<%
	If HasUserRight(ACCESS_CUSTOMER, RIGHT_READ) = false Then
		Response.Redirect("\xtra\IngenTilgang.asp")
	end if
	
	const DEFAULT_OPPDRAGVIKAR_STATUS = 1 'foresprsel
	dim foundRecords : foundRecords = false
	dim strSQL	
	dim Conn
	dim selection
	dim selected
	dim FirmaID
	dim rsRapport
	dim rsContact
	dim rsOppdrag
	dim rsStatus
	dim AkseptSelected
	dim gamleMed : gamleMed = false
	dim AlleSelected
	dim OPPD
	dim vid
	dim antDager
	dim nyttOppdrag
	dim tildato
	dim fradato
	dim prevOppdragID
	dim prevVikarID
	dim loenn
	dim pris
	dim prevOppdragStatus
	
		
   ' Open database connection
   Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	' Check for mandatory values
	'Check is this is first time we display this page..
	If lenb(Request.Form( "IsPostback")) = 0 Then
		'First time check for mandatory values
		' Do we have FirmaID ?
		If lenb(Request( "cuid" )) = 0 Then
			AddErrorMessage("Systemfeil: FirmaID mangler!")
			call RenderErrorMessage()
		Else
			strSQL = "SELECT [FirmaId] FROM [Firma] WHERE [SOCUID] = " & Request("cuid")
			set rsContact = GetFirehoseRS(strSQL, Conn)
			if(HasRows(rsContact) = true) then
				FirmaID = rsContact("FirmaId")			
				rsContact.close
			else
				AddErrorMessage("Systemfeil: Ugyldig SOCUID, finnes ikke i Xis!")
				call RenderErrorMessage()			
			end if
			set rsContact = nothing
		End If

		' Default value is 'Aksept'
		selection = DEFAULT_OPPDRAGVIKAR_STATUS
		gamleMed = true
	Else
		' Add values FROM current page
		FirmaID = Request( "tbxFirmaID" )
		selection = Request( "rbnSelection" )
	End If

   strSQL = "SELECT S.OppdragsStatus, OVS.Status, " &_
	"O.OppdragID, O.Beskrivelse, O.Fradato, OV.notat, " &_
	"OV.TimeLoenn, OV.Timepris, " &_
	"V.VikarID, Vikar =  V.Fornavn + ' ' + V.Etternavn " &_
	"FROM " &_
	"OPPDRAG AS O " &_
	"INNER JOIN h_oppdrag_Status AS S ON O.statusID = S.OppdragsStatusID " &_
	"LEFT OUTER JOIN OPPDRAG_VIKAR AS OV ON O.OppdragID = OV.OppdragID " &_
	"LEFT OUTER JOIN H_OPPDRAG_VIKAR_STATUS AS OVS ON OV.StatusID = OVS.OppdragVikarStatusID " &_
	"LEFT OUTER JOIN Vikar AS V ON OV.VikarID = V.VikarID " &_
	"WHERE O.FirmaID = " & FirmaID

   ' Append selection of status if other than all selected
   If selection > 0 Then
      strSQL = strSQL & " AND O.statusID = " & selection
   End If

   ' Append sorting
   strSQL = strSQL & " ORDER BY S.OppdragsStatus, O.Oppdragid DESC, V.Vikarid, O.Fradato DESC "
   ' Get all records   
	set rsRapport = GetFirehoseRS(strSQL, Conn)

  ' No records found ?
   if (HasRows(rsRapport) = true) then
      foundRecords = true
   End If
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
		<title>Tidligere vikarer for kontakt <%=FirmaID%></title>
	</head>
<body>
	<div class="pageContainer" id="pageContainer">
		<div class="content">		
			<form name="formEn" ACTION="kunde-oppdrag.asp" METHOD="POST" ID="Form1">
				<input type="hidden" NAME="IsPostback" VALUE="1" ID="IsPostback">
				<input type="hidden" NAME="tbxFirmaID" VALUE="<%=FirmaID%>" ID="tbxFirmaID">
				<%
				strSQL = "SELECT [OppdragsStatusID], [OppdragsStatus] FROM [h_oppdrag_Status] ORDER BY [OppdragsStatus]"
				set rsStatus = GetFirehoseRS(strSQL, Conn)
				if(HasRows(rsStatus) = true) then
					%>
					<td>
						Oppdragstatus:&#160;<select id="rbnSelection" name="rbnSelection">
							<%
							if(selection = DEFAULT_OPPDRAGVIKAR_STATUS) then
								selected = "selected"
							end if
							%>			
							<option <%=selected%> value="0">(Alle)</option>
							<%
							while not rsStatus.EOF
								if(clng(selection) = rsStatus("OppdragsStatusID").value) then
									selected = "selected"
								else
									selected = ""
								end if
								%>
								<option <%=selected%> value="<%=rsStatus("OppdragsStatusID")%>"><%=rsStatus("OppdragsStatus")%></option>
								<%
								rsStatus.MoveNext
							wend
							%>
						</select>
					</td>
					<%
					rsStatus.close
				end if
				set rsStatus = nothing
				%>
				&#160;<input type="submit" name="pbnDataAction" value="Søk" ID="pbnDataAction"></td>
			</form>
			<%
			If  (foundRecords)  Then
				prevOppdragStatus = ""
				prevOppdragID = 0
				prevVikarID = 0
				%>
				<div class="listing">
					<table>	
						<%
						Do Until (rsRapport.EOF)
								if (prevOppdragStatus <> rsRapport("OppdragsStatus")) then
									Response.Write "<tr><th colspan='7'><strong>" & rsRapport("OppdragsStatus") & "</strong></th></tr>"
									%>
									<tr>
										<th>Oppdragnr</th>
										<th>Beskrivelse</th>
										<th>Vikar</th>
										<th>Status</th>
										<th>Lnn</th>
										<th>Timepris</th>
										<th>Notat</th>
									</tr>									
									<%
								end if
								if (prevVikarID <> rsRapport("vikarid") OR prevOppdragID <> rsRapport("OppdragID")) then
									Response.Write "<tr>"
									if (prevOppdragID <> rsRapport("OppdragID")) then
										Response.Write "<td>" & CreateSOLink(SUPEROFFICE_XIS_TASK_URL, "", "WebUI/OppdragView.aspx?OppdragID=" & rsRapport( "OppdragID" ), rsRapport( "OppdragID"), "Vis oppdrag """ & rsRapport( "Beskrivelse") & """" )  &  "</td>"
										Response.Write "<td>" & rsRapport( "Beskrivelse" ) & "&#160;</td>"
									else
										Response.Write "<td colspan='2'>&#160;</td>"
									end if	
									Response.Write "<td>" & CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & rsRapport( "VikarID" ), rsRapport( "Vikar" ).value, "Vis vikar " & rsRapport( "Vikar" ).value ) & "&#160;</td>"
									Response.Write "<td>" & rsRapport( "Status" ) & "&#160;</td>"
									loenn = rsRapport( "TimeLoenn" )
									if (not isnull(loenn)) then
										loenn = formatNumber(loenn, 2)
									end if
									
									pris = rsRapport( "Timepris" )
									if (not isnull(pris)) then
										pris = formatNumber(pris, 2)
									end if								
									Response.Write "<td class='right'>" & loenn & "&#160;</td>"
									Response.Write "<td class='right'>" & pris & "&#160;</td>"
									Response.Write "<td>" & rsRapport( "notat" ) & "&#160;</td>"
									Response.Write "</tr>"
								end if
							prevOppdragID = rsRapport("OppdragID")
							prevVikarID = rsRapport("vikarid")
							prevOppdragStatus = rsRapport("OppdragsStatus")
							rsRapport.MoveNext
						Loop
						rsRapport.close
						%>
					</table>
					<% 
					else
						Response.Write "<p>Ingen oppdrag funnet på kontakt.</p>"
					End If 
					CloseConnection(Conn)
					set rsRapport = Nothing
					set Conn = nothing		
					%>
				</div>
			</div>
		</div>
	</body>
</html>
