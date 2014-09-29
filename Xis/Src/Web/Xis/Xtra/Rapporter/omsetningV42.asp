<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Reports.Utils.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Economics.Constants.inc"-->
<%
	If (HasUserRight(ACCESS_REPORT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim Conn			'Connection
	dim strSQL
	dim lngPrevAvd
	dim lngPrevMedID
	dim lngMedID
	dim lngTomID
	dim lngCurrentAvdID
	dim rsRapport
	dim rsDB
	dim rsOvertid
	dim strFradato
	dim strTildato
	dim strFraDatoUke
	dim strFraDatoAar
	dim strTilDatoUke
	dim strTilDatoAar
	dim strFraDatoAarsuke
	dim strTilDatoAarsuke
	dim strSelectedAvd
	dim strSQLSelectedAvd
	dim inttypeID
	dim dblDekningsbidrag
	dim dblDekningsbidragOvertid
	dim dblOvertimeSum
	dim dblSalarySum
	dim dblSalaryHours
	dim dblOmsetningSum
	dim dblOmsetningHours
	dim dblFactor
	dim lngOvertidProsent
	dim dblTmpSalarySum
	dim ATotTom()
	dim rsToms
	dim intTotToms
	dim intTOMTeller
	dim	dblConsultantOmsetningTotal
	dim	dblConsultantDekningsbidrag
	dim	dblConsultantSalarySum
	dim	intConsultantAntallOppdrag
	dim dblConsultantFactorBase
	dim intNOFConsultantRecords
	dim strCurrentAvdeling
	dim ATotAllToms()

	' Functions
	sub CalculateAndRenderSums()

			'Get DB for this consultantleader's current service area
			dblOmsetningSum 			= rsRapport("omsetning")
			dblOmsetningHours 			= rsRapport("TotFakturaTimer")
			dblSalarySum 	 			= rsRapport("loenn")
			dblSalaryHours 	 			= rsRapport("TotAntTimer")
			dblDekningsbidrag			= 0.00
			dblDekningsbidragOvertid	= 0.00
			lngMedID					= rsRapport("MedID")
			lngTomID					= rsRapport("TomID")
			lngCurrentAvdID				= rsRapport("AvdelingID")


			'Calculate db for current consultant, avdeling and servicearea, of type "1"
			strSQL = "exec sp_GET_DB_BY_DEPARTMENTAREA " & strFraDato & ", " & strTilDato & "," & lngMedID &  ", " & lngCurrentAvdID & ", " &  lngTomID  & ", 1"
			set rsDB = GetFirehoseRS(strSQL, Conn)
			if (NOT isnull(rsDB("DB").value)) then
				dblDekningsbidrag = rsDB("DB").value
			end if

			'Calculate db for current consultant, avdeling and servicearea, of all other types
			strSQL = "exec sp_GET_DB_BY_DEPARTMENTAREA " & strFraDato & ", " & strTilDato & "," & lngMedID &  ", " & lngCurrentAvdID & ", " & lngTomID  & ",0"
			set rsDB = GetFirehoseRS(strSQL, Conn)			
			if (NOT isnull(rsDB("DB").value)) then
				dblDekningsbidrag = dblDekningsbidrag + rsDB("DB").value
			end if

			'Calculate overtime, hours, salary and db for current consultant, avdeling and servicearea
			strSQL = "exec GET_OVERTIME_BY_DEPARTMENTAREA " & lngMedID &  " ," & lngCurrentAvdID & ", " & strFraDatoAarsuke & ", " & strTilDatoAarsuke & ", " & lngTomID 
			set rsOvertid = GetFirehoseRS(strSQL, Conn)
			if (HasRows(rsOvertid) = true) then
				while (NOT rsOvertid.EOF)
					'Overtime hours & total hours
					dblOmsetningHours 	= dblOmsetningHours + rsOvertid("FTimer")
					dblOvertimeSum 	 	= rsOvertid("FBelop")
					dblTmpSalarySum  	= rsOvertid("LBelop")
					inttypeID			= cint(rsOvertid("TypeID"))

					dblDekningsbidragOvertid = 0.00

					If (inttypeID = 1) Then
						dblDekningsbidragOvertid = dblOvertimeSum - ( dblTmpSalarySum * XIS_FACTOR )
					Else
						dblDekningsbidragOvertid = dblOvertimeSum -  dblTmpSalarySum
					End If

					dblDekningsbidrag 	= dblDekningsbidrag + dblDekningsbidragOvertid
					dblOmsetningSum 	= dblOmsetningSum + dblOvertimeSum
					dblSalarySum 		= dblSalarySum + dblTmpSalarySum

					rsOvertid.movenext
				wend
			end if
			set rsOvertid = nothing

			'Calculate factor..
			If dblOmsetningSum > 0 and (dblDekningsbidrag> 0) then
				dblFactor = dblOmsetningSum / ((dblOmsetningSum - dblDekningsbidrag) / XIS_FACTOR)
			else
				dblFactor = 0.0
			end if

			'Render sum for this service area for this consultant
			%>
				<tr>
					<td><%=rsRapport("tjenesteomrade")%></td>
					<td class="right"><%=FormatNumber( dblOmsetningSum, 2)%></TD>
					<td class="right"><%=FormatNumber( dblDekningsbidrag, 2)%></TD>
					<td class="right"><%=FormatNumber( dblSalarySum, 2)%></TD>
					<td class="right"><%=rsRapport("antallOppdrag")%></TD>
					<td class="right"><%=FormatNumber( dblFactor, 2)%></TD>
				</tr>
			<%
			'Add to consultant total
			intNOFConsultantRecords		= intNOFConsultantRecords + 1
			dblConsultantOmsetningTotal = dblConsultantOmsetningTotal + dblOmsetningSum
			dblConsultantDekningsbidrag = dblConsultantDekningsbidrag + dblDekningsbidrag
			dblConsultantSalarySum		= dblConsultantSalarySum + dblSalarySum
			intConsultantAntallOppdrag 	= intConsultantAntallOppdrag + rsRapport("antallOppdrag")
			dblConsultantFactorBase		= dblConsultantFactorBase + dblFactor

			'Add to department total
			intArrayIndex = getIndexByTomID(lngTomID)

			ATotTom(intArrayIndex,3) = ATotTom(intArrayIndex,3) + dblOmsetningSum
			ATotTom(intArrayIndex,5) = ATotTom(intArrayIndex,5) + dblSalarySum
			ATotTom(intArrayIndex,4) = ATotTom(intArrayIndex,4) + dblDekningsbidrag
			ATotTom(intArrayIndex,6) = ATotTom(intArrayIndex,6) + rsRapport("antallOppdrag")
	end sub


	sub renderConsultantTomSum()
		'Calculate factor for all TOM's for this consultant
		If dblConsultantOmsetningTotal > 0 and (dblConsultantDekningsbidrag> 0) then
			dblConsultantFactor = dblConsultantOmsetningTotal / ((dblConsultantOmsetningTotal - dblConsultantDekningsbidrag) / XIS_FACTOR)
		else
			dblConsultantFactor = 0.0
		end if
		%>
		<tr>
			<td>Sum</td>
			<td class="right"><%=FormatNumber( dblConsultantOmsetningTotal, 2)%></td>
			<td class="right"><%=FormatNumber( dblConsultantDekningsbidrag, 2)%></td>
			<td class="right"><%=FormatNumber( dblConsultantSalarySum, 2)%></td>
			<td class="right"><%=intConsultantAntallOppdrag%></td>
			<td class="right"><%=FormatNumber(dblConsultantFactor, 2)%></td>
		</tr>
		<%
		dblConsultantOmsetningTotal = 0.0
		dblConsultantDekningsbidrag = 0.0
		dblConsultantSalarySum		= 0.0
		intConsultantAntallOppdrag 	= 0
		dblConsultantFactor			= 0.0
		intNOFConsultantRecords		= 1
	end sub


	sub RenderValuesByDepartment
		dim intArrayRunner
		dim intArrayCount
		dim dblDeptTempOmsetningTotal
		dim dblDeptTempDekningsbidrag
		dim dblDeptTempSalarySum
		dim intDeptTempAntallOppdrag
		dim dblDeptTempFactor
		%>
		<h2>Totalt for avdeling <%=strCurrentAvdeling%></h2>
		<div class="listing">
			<table CELLPADDING="0" CELLSPACING="1" class="reportTable" ID="Table1">
				<TR>
					<TH>Tjenesteområde</TH>
					<TH>Omsetning</TH>
					<TH>Bidrag</TH>
					<TH>Lønn</TH>
					<TH>Oppdrag</TH>
				</TR>
				<%
				intArrayCount = ubound(ATotTom)
				intArrayRunner = 1
				while (intArrayRunner <= intArrayCount)
					%>
					<TR>
						<TD><%=ATotTom(intArrayRunner,2)%></TD>
						<td class="right"><%=FormatNumber( ATotTom(intArrayRunner,3), 2)%></TD>
						<td class="right"><%=FormatNumber( ATotTom(intArrayRunner,4), 2)%></TD>
						<td class="right"><%=FormatNumber( ATotTom(intArrayRunner,5), 2)%></TD>
						<td class="right"><%=ATotTom(intArrayRunner,6)%></TD>
					</TR>
					<%
					'Add to totals og all service areas for this department
					dblDeptTempOmsetningTotal 	= dblDeptTempOmsetningTotal + ATotTom(intArrayRunner,3)
					dblDeptTempDekningsbidrag 	= dblDeptTempDekningsbidrag + ATotTom(intArrayRunner,4)
					dblDeptTempSalarySum		= dblDeptTempSalarySum + ATotTom(intArrayRunner,5)
					intDeptTempAntallOppdrag 	= intDeptTempAntallOppdrag + ATotTom(intArrayRunner,6)

					'Add to all totals
					ATotAllToms(intArrayRunner,3) = ATotAllToms(intArrayRunner,3) + ATotTom(intArrayRunner,3)
					ATotAllToms(intArrayRunner,4) = ATotAllToms(intArrayRunner,4) + ATotTom(intArrayRunner,4)
					ATotAllToms(intArrayRunner,5) = ATotAllToms(intArrayRunner,5) + ATotTom(intArrayRunner,5)
					ATotAllToms(intArrayRunner,6) = ATotAllToms(intArrayRunner,6) + ATotTom(intArrayRunner,6)	
		
					'Reset values for all service areas for this department
					ATotTom(intArrayRunner,3) = 0.0
					ATotTom(intArrayRunner,4) = 0.0
					ATotTom(intArrayRunner,5) = 0.0
					ATotTom(intArrayRunner,6) = 0
					intArrayRunner = intArrayRunner + 1
				wend
				%>
				<tfoot>
					<tr>
						<Td>Sum alle tjenesteområder</Td>
						<td class="right"><%=FormatNumber( dblDeptTempOmsetningTotal, 2)%></TD>
						<td class="right"><%=FormatNumber( dblDeptTempDekningsbidrag, 2)%></TD>
						<td class="right"><%=FormatNumber( dblDeptTempSalarySum, 2)%></TD>
						<td class="right"><%=intDeptTempAntallOppdrag%></TD>
					</tr>
				</tfoot>
		</table>
	</div>
	<%
	end sub

	sub RenderTotalsForAllDepartments
		dim intArrayRunner
		dim intArrayCount
		dim dblDeptTempOmsetningTotal
		dim dblDeptTempDekningsbidrag
		dim dblDeptTempSalarySum
		dim intDeptTempAntallOppdrag
		dim dblDeptTempFactor
		%>
		<h2>Totalt for alle tjenesteområder for alle avdelinger</h2>
		<div class="listing">
			<table CELLPADDING="0" CELLSPACING="1" class="reportTable" ID="Table2">
				<TR>
					<TH>Tjenesteområde</TH>
					<TH>Omsetning</TH>
					<TH>Bidrag</TH>
					<TH>Lønn</TH>
					<TH>Oppdrag</TH>
				</TR>
				<%
				intArrayCount = ubound(ATotAllToms)
				intArrayRunner = 1
				while (intArrayRunner <= intArrayCount)
					%>
					<tr>
						<td><%=ATotAllToms(intArrayRunner,2)%></td>
						<td class="right"><%=FormatNumber( ATotAllToms(intArrayRunner,3), 2)%></TD>
						<td class="right"><%=FormatNumber( ATotAllToms(intArrayRunner,4), 2)%></TD>
						<td class="right"><%=FormatNumber( ATotAllToms(intArrayRunner,5), 2)%></TD>
						<td class="right"><%=ATotAllToms(intArrayRunner,6)%></td>
					</tr>
					<%
					'Add to totals og all service areas for this department
					dblDeptTempOmsetningTotal 	= dblDeptTempOmsetningTotal + ATotAllToms(intArrayRunner,3)
					dblDeptTempDekningsbidrag 	= dblDeptTempDekningsbidrag + ATotAllToms(intArrayRunner,4)
					dblDeptTempSalarySum		= dblDeptTempSalarySum + ATotAllToms(intArrayRunner,5)
					intDeptTempAntallOppdrag 	= intDeptTempAntallOppdrag + ATotAllToms(intArrayRunner,6)

					intArrayRunner = intArrayRunner + 1
				wend
				%>
				<tfoot>
					<tr>
						<td>Sum alle tjenesteområder</td>
						<td class="right"><%=FormatNumber( dblDeptTempOmsetningTotal, 2)%></TD>
						<td class="right"><%=FormatNumber( dblDeptTempDekningsbidrag, 2)%></TD>
						<td class="right"><%=FormatNumber( dblDeptTempSalarySum, 2)%></TD>
						<td class="right"><%=intDeptTempAntallOppdrag%></TD>
					</tr>
				</tfoot>
			</table>
		</div>
		<%
	end sub

	sub RenderDepartmentSum()
			call RenderValuesByDepartment()
			ATotTom(intArrayIndex,3) = 0.0
			ATotTom(intArrayIndex,5) = 0.0
			ATotTom(intArrayIndex,4) = 0.0
			ATotTom(intArrayIndex,6) = 0
	end sub

	function getIndexByTomID(lngTomID)
		dim intArrayRunner
		dim intArrayCount
		dim blnExitLoop

		blnExitLoop = false

		intArrayCount = ubound(ATotTom)
		intArrayRunner = 1
		while (intArrayRunner <= intArrayCount OR blnExitLoop=false)
			if cint(ATotTom(intArrayRunner,1)) = cint(lngTomID) then
				getIndexByTomID = intArrayRunner
				blnExitLoop = true
			end if
			intArrayRunner = intArrayRunner + 1
		wend
	end function

' /Functions

'Main body starts here
	if Request("POSTED") = 1 then

		strFradato  	 = Request.Form( "tbxFradato" )
		strTildato  	 = Request.Form( "tbxTildato" )
		strSelectedAvd 	 = Request.form("chkAvdeling")
		strSQLSelectedAvd = strSelectedAvd

		if strFradato <> "" then
			strFraDatoUke = strFradato
			strFraDatoAar = strFradato
			call KorrigerUke(strFraDatoUke)
			call KorrigerAar(strFraDatoAar)
			strFraDatoAarsuke = strFraDatoAar & strFraDatoUke
			strFradato     = DbDate(strFradato)
		else
			AddErrorMessage("Fradato må fylles ut!")
			call RenderErrorMessage()
		end if

		if strTildato <> "" then
			strTilDatoUke = strTildato
			strTilDatoAar = strTildato
			call KorrigerUke(strTilDatoUke)
			call KorrigerAar(strTilDatoAar)
			strTilDatoAarsuke = strTilDatoAar & strTilDatoUke
			strTildato     = DbDate(strTildato)
		else
			AddErrorMessage("Tildato må fylles ut!")
			call RenderErrorMessage()
		end if
		
		if(HasError() = true) then
			call RenderErrorMessage()
		end if		
	end if

	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	strSQL = "SELECT COUNT([tomID]) AS NofRec FROM [Tjenesteomrade] "
	set rsToms = GetFirehoseRS(strSQL, Conn)
	intTotToms = rsToms.fields("NofRec")
	redim ATotTom(intTotToms, 7)
	redim ATotAllToms(intTotToms, 7)

	strSQL = "SELECT [tomID], [Navn] FROM [Tjenesteomrade] ORDER BY [tomID]"
	set rsToms = GetFirehoseRS(strSQL, Conn)
	if (hasRows(rsToms) = true) then
		intTOMTeller = 1
		while not rsToms.EOF
			ATotTom(intTOMTeller,1) = rsToms.fields("tomid").value
			ATotTom(intTOMTeller,2) = rsToms.fields("navn").value
			ATotTom(intTOMTeller,3) = 0.0
			ATotTom(intTOMTeller,4) = 0.0
			ATotTom(intTOMTeller,5) = 0.0
			ATotTom(intTOMTeller,6) = 0
			ATotTom(intTOMTeller,7) = 0.0
			
			ATotAllToms(intTOMTeller,1) = rsToms.fields("tomid").value
			ATotAllToms(intTOMTeller,2) = rsToms.fields("navn").value
			ATotAllToms(intTOMTeller,3) = 0.0
			ATotAllToms(intTOMTeller,4) = 0.0
			ATotAllToms(intTOMTeller,5) = 0.0
			ATotAllToms(intTOMTeller,6) = 0
			ATotAllToms(intTOMTeller,7) = 0.0
			
			rsToms.movenext
			intTOMTeller = intTOMTeller + 1
		wend		
	end if
	rsToms.close
	set rsToms = nothing	
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
		<title>Omsetning pr. avdelingskontor/ansvarlig/tjenesteområde </title>
		<script language="javascript" src="../Js/javascript.js" type="text/javascript"></script>
		<script language="javaScript" type="text/javascript">
			function DeSelect(objAll)
			{

				if (objAll.checked==true)
				{
					selectAll()
				}else{
					deSelectAll()
				}
			}

			function selectAll()
			{
				for(i = 0; i < document.formEn.chkAvdeling.length; i++)
				{
					document.formEn.chkAvdeling[i].checked=true;
				}

				document.all.chkAvdeling.checked = true;
			};

			function deSelectAll()
			{
				for(i = 0; i < document.formEn.chkAvdeling.length; i++)
					{
						document.formEn.chkAvdeling[i].checked = false;
					}
			};
		</script>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Omsetning pr. avdelingskontor/ansvarlig/tjenesteområde </h1>
			</div>
			<div class="content">gir kun riktig overtid hvis datointervallet er en hel måned.<br>Faktor i totalene er gjennomsnittlige.
				<form name="formEn" ACTION="omsetningV42.asp" METHOD="POST" ID="Form1">
					<input type='HIDDEN' value='1' name='POSTED' id='POSTED'>
					<table>
						<tr>
							<td>Fra dato:</TD>
							<td><INPUT valign="top" NAME="tbxFraDato" TYPE='TEXT' SIZE='10' MAXLENGTH='10' Value="<%=Request.Form( "tbxFradato")%>" ONBLUR="dateCheck(this.form, this.name)" ID="Text1"> </TD>
							<td>Til dato:</TD>
							<td><INPUT valign="top" NAME="tbxTilDato" TYPE='TEXT' SIZE='10' MAXLENGTH='10' Value="<%=Request.Form( "tbxTildato" )%>" ONBLUR="dateCheck(this.form, this.name)" ID="Text2"> </TD>
							<td>Avdeling:</td>
							<td>
								<div class="container">
    								<input class="checkbox" TYPE='CHECKBOX' ID='chkAvdelingAll' NAME='chkAvdelingAll' onClick='javascript:DeSelect(this)' VALUE=' '>Alle<BR>
									<%
									dim strSel			'Contains "SELECTED" if current departement is selected
									dim rsAvdeling		'Recordset containing departments
									dim NofRecords
									dim lngAnsMedID

									if (not ISNULL(Session("medarbID")) ) AND (Not ISEMPTY(Session("medarbID")) ) then
										lngAnsMedID = Session("medarbID")
									end if

									'Get vikarid for this medarbeider
									strSQL = "SELECT vikarid FROM medarbeider WHERE medID =" & lngAnsMedID
									set rsAvdeling = GetFirehoseRS(strSQL, Conn)
									lngVikarid = rsAvdeling("vikarid")

		   							if (strSelectedAvd = "") then
		   								strSQL = "SELECT AvdelingskontorID FROM vikar_arbeidssted WHERE vikarid = " & lngVikarid
		   								set rsAvdeling = GetFirehoseRS(strSQL, Conn)
		   								strSelectedAvd = "," & rsAvdeling("AvdelingskontorID").value & ","
		   							else
		   								strSelectedAvd = "," & strSelectedAvd & ","
		   							end if

		   							strSQL = "SELECT id AS AvdelingID, navn AS Avdeling FROM AvdelingsKontor ORDER BY avdeling"
									set rsAvdeling = GetFirehoseRS(strSQL, Conn)
									strSelectedAvd = replace(strSelectedAvd," ","")

									Do Until (rsAvdeling.EOF)
										strSel = ""
										if (strSelectedAvd <> "") then
											if instr(strSelectedAvd,"," & rsAvdeling("AvdelingID") & ",") then
												strSel = "Checked"
											end if
										end if
										%>
										<div><input class="checkbox" TYPE='CHECKBOX' <%=strSel%> ID='chkAvdeling' NAME='chkAvdeling' VALUE='<%=rsAvdeling("AvdelingID")%>'><%=rsAvdeling("Avdeling")%></div>
										<%
										rsAvdeling.MoveNext
										NofRecords = NofRecords + 1
									Loop
									' Close and release recordset
									rsAvdeling.Close
									Set rsAvdeling = Nothing
									%>
								</div>
   							</td>
							<td valign="top"><input valign="top" type="submit" name="pbnDataAction" value="     Søk    " ID="Submit1"></td>
						</tr>
					</table>
				</form>
			</div>
			<%
			
			if Request.Form("POSTED") = 1 then
				%>
				<div class="contentHead"><h2>Rapport</h2></div>
				<div class="content">
					<%
					lngPrevAvd = 0
					strSQL = "EXEC sp_GET_OMSETNING_BY_DEPARTMENTAREA " & strFraDato & ", " & strTilDato & ",'" & strSQLSelectedAvd & "'"
					set rsRapport = GetFirehoseRS(strSQL, Conn)
					NofRecords = 0
					lngPrevMedID = 0
					if (HasRows(rsRapport) = true) then
						strCurrentAvdeling = rsRapport("Avdeling")
						%>
						<h1><%=strCurrentAvdeling%></h1>
						<%
						while not (rsRapport.EOF)
							'New department group..
							if ((lngPrevAvd <> 0) AND (lngPrevAvd <> clng(rsRapport("AvdelingID")) ) ) then
								call renderConsultantTomSum()
								%>
								</Table>
								<%
								call renderDepartmentSum()
								strCurrentAvdeling = rsRapport("Avdeling")
								%>
								</div>
								<div class="contentHead">
									<h2><%=strCurrentAvdeling%></h2>
									<p><%=rsRapport("etternavn") & ", " & rsRapport("fornavn")%></p>
								</div>
								<div class="listing">
								<TABLE CELLPADDING="0" CELLSPACING="1" class="reportTable" ID="Table4">
								<TR>
									<TH>Tjenesteområde</TH>
									<TH>Omsetning</TH>
									<TH>Bidrag</TH>
									<TH>Lønn</TH>
									<TH>Oppdrag</TH>
									<TH>Faktor</TH>
								</TR>
							<%
							'New consultant leader..
							elseif ( (lngPrevMedID <> clng(rsRapport("MedID"))) )  then
								if (lngPrevMedID <> 0) then
									call renderConsultantTomSum()
									response.write "</table></div>"
								end if
								%>
								<h3><%=rsRapport("etternavn") & ", " & rsRapport("fornavn")%></h3>
								<div class="listing">
									<table CELLPADDING="0" CELLSPACING="1" class="reportTable" ID="Table5">
										<tr>
											<TH>Tjenesteområde</TH>
											<TH>Omsetning</TH>
											<TH>Bidrag</TH>
											<TH>Lønn</TH>
											<TH>Oppdrag</TH>
											<TH>Faktor</TH>
										</tr>
								<%
							end if
							call CalculateAndRenderSums()

							lngPrevAvd 		= clng(rsRapport("AvdelingID"))
							lngPrevMedID	= clng(rsRapport("medid"))

							NofRecords = NofRecords + 1
							rsRapport.movenext
						wend
					end if

					call renderConsultantTomSum()
					response.write "</table></div>"
					call renderDepartmentSum()
					call RenderTotalsForAllDepartments()

					set rsRapport 	= nothing
					set rsDB 		= nothing
					set rsOvertid 	= nothing
					%>
					<P ID="sideskift" STYLE="page-break-after='always'">&nbsp;</P>
				</div>
				<%
			end if
			%>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>