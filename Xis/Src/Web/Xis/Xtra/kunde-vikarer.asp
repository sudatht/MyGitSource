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
	
	const DEFAULT_OPPDRAGVIKAR_STATUS = 4
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
	dim ColumnValue
	dim SorttypeValue
	
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
   

   strSQL = "SELECT DISTINCT OV.VikarID, Navn=(V.Fornavn + ' ' + V.Etternavn), OV.OppdragID AS OppdragID, " &_
		"OV.Fradato, OV.Tildato, " &_
		"O.Beskrivelse, S.Status "&_
		" FROM OPPDRAG_VIKAR OV, OPPDRAG O, VIKAR V, H_OPPDRAG_VIKAR_STATUS S " &_
		" WHERE OV.FirmaID = " & FirmaID &_
		" AND O.OppdragID = OV.OppdragID" &_
		" AND V.VikarID   =* OV.VikarID" &_
		" AND S.OppdragVikarStatusID = OV.StatusID"

   ' Append selection of status if other than all selected
   If selection > 0 Then
      strSQL = strSQL & " AND OV.statusID = " & selection
   End If
   
   
   If gamleMed = true Then
		strSQL = strSQL & " UNION " &_
			"SELECT VikarID = o.vikarnr, Navn=(V.Fornavn + ' ' + V.Etternavn), OppdragID=o.Oppdragsnr, " &_
			"Fradato = o.StartDato, Tildato=o.Sluttdato, " &_
			"o.Beskrivelse, Status=S.OppdragsStatus " &_
			"FROM gamle_oppdrag o, Vikar V, H_OPPDRAG_STATUS S " &_
			"WHERE Kundenr = " & FirmaID &_
			" AND S.OppdragsStatusID = o.StatusID " &_
			"AND o.vikarnr = V.VikarID"'
	End If
	
   'Response.Write strSQL & " ------ " 
  
   ColumnValue = Request("column")
   SorttypeValue = Request("sorttype")
   
   'Response.Write ColumnValue & " ------ " 
   'Response.Write SorttypeValue & " ------ " 
   
   ' Append order by based on selection
   IF(ColumnValue <> "") Then
   		IF(ColumnValue = "Oppdrag" and SorttypeValue = "asc") Then
   			strSQL = strSQL & " ORDER BY OppdragID "
   		End If
   		IF(ColumnValue = "Oppdrag" and SorttypeValue = "des") Then
   			strSQL = strSQL & " ORDER BY OppdragID desc "
   		End If
   		IF(ColumnValue = "Vikar" and SorttypeValue = "asc") Then
   			strSQL = strSQL & " ORDER BY Navn"
   		End If
   		IF(ColumnValue = "Vikar" and SorttypeValue = "des") Then
   			strSQL = strSQL & " ORDER BY Navn desc "
   		End If   		
   Else
   		strSQL = strSQL & " ORDER BY OV.Fradato DESC"
   End If

   'Response.Write strSQL & " ------ " 

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
		<title>Tidligere vikarer for kontakt</title>
		
		<script language="javascript">
			function sort(column,sorttype) {				
				document.all.column.value=column;
				document.all.sorttype.value=sorttype;
				document.formEn.submit();
			}

		</script>
	</head>
<body>
	<div class="pageContainer" id="pageContainer">
		<div class="content">		
			<form name="formEn" ACTION="kunde-vikarer.asp" METHOD="POST" ID="Form1">
				<input type="hidden" NAME="IsPostback" VALUE="1" ID="IsPostback">
				<input type="hidden" NAME="tbxFirmaID" VALUE="<%=FirmaID%>" ID="tbxFirmaID">
				<input type="hidden" id="column" name="column" value="">
				<input type="hidden" id="sorttype" name="sorttype" value="">
				<%
				strSQL = "SELECT [OppdragVikarStatusID], [Status] FROM [H_OPPDRAG_VIKAR_STATUS] ORDER BY [Status]"
				set rsStatus = GetFirehoseRS(strSQL, Conn)
				if(HasRows(rsStatus) = true) then
					%>
					<td>
						Status:&#160;<select id="rbnSelection" name="rbnSelection">
							<%
							if(selection = DEFAULT_OPPDRAGVIKAR_STATUS) then
								selected = "selected"
							end if
							%>			
							<option <%=selected%> value="0">(alle)</option>
							<%
							while not rsStatus.EOF
								if(clng(selection) = rsStatus("OppdragVikarStatusID").value) then
									selected = "selected"
								else
									selected = ""
								end if
								%>
								<option <%=selected%> value="<%=rsStatus("OppdragVikarStatusID")%>"><%=rsStatus("Status")%></option>
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
				%>
				<div class="listing">
					<table>	
						<tr>
							<th>
								Vikar&nbsp;
								<img src="/xtra/images/sort_ascending.jpg" width="9" height="9" alt="sort by ascending" align="absmiddle" onclick="sort('Vikar','asc');">
								<img src="/xtra/images/sort_descending.jpg" width="9" height="9" alt="sort by descending" align="absmiddle" onclick="sort('Vikar','des');">
							</th>
							<th>
								Oppdragnr&nbsp;
								<img src="/xtra/images/sort_ascending.jpg" width="9" height="9" alt="sort by ascending" align="absmiddle" onclick="sort('Oppdrag','asc');">
								<img src="/xtra/images/sort_descending.jpg" width="9" height="9" alt="sort by descending" align="absmiddle" onclick="sort('Oppdrag','des');">
							</th>
							<th>Beskrivelse</th>
							<th>Status</th>
							<th>Fra dato</th>
							<th>Til dato</th>
							<th>Ant Dager</th>
						</tr>
						<%
						
						OPPD = rsRapport("OppdragID")
						VID = rsRapport("VikarID")
						antDager = 0
						nyttOppdrag = True
						tilDato = rsRapport("TilDato")
						fraDato = rsRapport("FraDato")
						antDager = (tilDato - fraDato)+1

						Do Until (rsRapport.EOF)
							'skjer hver gang det er likt for å samle til- og fradato
							If (rsRapport("VikarID") = VID AND rsRapport("OppdragID") = OPPD) Then

								If (tilDato < rsRapport("TilDato")) Then
									tilDato = rsRapport("tildato")
								end if
								If (fraDato > rsRapport("fraDato")) Then
									fraDato = rsRapport("fradato")
								end if
								antDager = (tildato - fradato) + 1

								If nyttOppdrag Then 'skjer hvergang det er nytt oppdrag eller nytt vikarnr

									nyttOppdrag = False
									Response.Write "<tr>"
									Response.Write "<td>" & CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "vikarvis.asp?VikarID=" & rsRapport( "VikarID" ), rsRapport( "Navn"), "Vis vikar " & rsRapport( "Navn") )  & "</td>"

									If rsRapport("Status") = "Ferdig" Then
										Response.Write "<td>" & rsRapport("OppdragID") & "</td>"
									Else
										Response.Write "<td>" & CreateSOLink(SUPEROFFICE_XIS_TASK_URL, "", "WebUI/OppdragView.aspx?OppdragID=" & rsRapport( "OppdragID" ), rsRapport( "OppdragID"), "Vis oppdrag med oppdragnr " & rsRapport( "OppdragID") )  &  "</td>"
									End If
									Response.Write "<td>" & rsRapport( "Beskrivelse" ) & "&#160;</td>"
									Response.Write "<td>" & rsRapport( "Status" ) & "&#160;</td>"

								End If
							Else 'det er nytt oppdrag eller vikarnr

								Response.Write "<td>" & fradato & "&#160;</td>"
								Response.Write "<td>" & tilDato & "&#160;</td>"
								Response.Write "<td class='right'>" & antDager & "&#160;</td>"

								Response.Write "<tr>"
								Response.Write "<td>" & CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "vikarvis.asp?VikarID=" & rsRapport( "VikarID" ), rsRapport( "Navn"), "Vis vikar " & rsRapport( "Navn") )  & "</td>"

								If rsRapport("Status") = "Ferdig" Then
									Response.Write "<td>" & rsRapport("OppdragID") & "</td>"
								Else
									Response.Write "<td>" & CreateSOLink(SUPEROFFICE_XIS_TASK_URL, "", "WebUI/OppdragView.aspx?OppdragID=" & rsRapport( "OppdragID" ), rsRapport( "OppdragID"), "Vis oppdrag med oppdragnr " & rsRapport( "OppdragID") )  &  "</td>"
								End If

								Response.Write "<td>" & rsRapport( "Beskrivelse" ) & "&#160;</td>"
								Response.Write "<td>" & rsRapport( "Status" ) & "&#160;</td>"

								nyttOppdrag = True
								OPPD = rsRapport("OppdragID")
								VID = rsRapport("VikarID")

								antDager = (rsRapport("Tildato") - rsRapport("Fradato")) + 1
								tilDato = rsRapport("TilDato")
								fraDato = rsRapport("FraDato")
							End If
							rsRapport.MoveNext
						Loop
						rsRapport.close
						Response.Write "<td>" & fradato & "&#160;</td>"
						Response.Write "<td>" & tilDato & "&#160;</td>"
						Response.Write "<td class='right'>" & antDager & "&#160;</td>"
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