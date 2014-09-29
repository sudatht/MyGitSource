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

	dim Conn		'Connection
	dim lngPrevFirmID
	dim lngPrevSubID
	dim rsRapport
	dim rsDB
	dim rsOvertid
	dim strCurrentHovedOppforing
	dim strFradato
	dim strTildato
	dim strFraDatoUke
	dim strFraDatoAar
	dim strTilDatoUke
	dim strTilDatoAar
	dim strFraDatoAarsuke
	dim strTilDatoAarsuke
	dim strSelectedOppf
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
	dim	dblFirmOmsetningTotal
	dim	dblFirmDekningsbidrag
	dim	dblFirmSalarySum
	dim	intFirmAntallOppdrag
	dim	dblTotalOmsetningTotal
	dim	dblTotalDekningsbidrag
	dim	dblTotalSalarySum
	dim	intTotalAntallOppdrag
	dim dblFirmFactorBase
	dim intNOFFirmRecords
	dim intNOFTotalsRecords 
	dim strCurrentOppforing
	dim ATotAllToms()

	sub RenderDepartmentAreaTotals()
		'Get DB for this  current service area
		dblOmsetningSum 			= rsRapport("omsetning")
		dblOmsetningHours 			= rsRapport("TotFakturaTimer")
		dblSalarySum 	 			= rsRapport("loenn")
		dblSalaryHours 	 			= rsRapport("TotAntTimer")
		dblDekningsbidrag			= 0.00
		dblDekningsbidragOvertid	= 0.00
		lngCurrentFirmaID			= rsRapport("FirmaID")			
		lngTomID					= rsRapport("TomID")			

		'Calculate db for current firm, of type "1"
		strSQL = "EXEC GET_DB_FOR_CUSTOMERANDDEPARTMENT " & strFraDato & ", " & strTilDato & "," & lngCurrentFirmaID & "," &  lngTomID & ", 1"
		Set rsDB = GetFirehoseRS(strSQL, Conn)
		if (NOT isnull(rsDB("DB").value)) then
			dblDekningsbidrag = rsDB("DB").value
		end if

		'Calculate db for current firm, of all other types
		strSQL = "EXEC GET_DB_FOR_CUSTOMERANDDEPARTMENT " & strFraDato & ", " & strTilDato & "," & lngCurrentFirmaID & "," &  lngTomID & ", 0"
		Set rsDB = GetFireHoseRS(strSQL, Conn)
		if (NOT isnull(rsDB("DB").value)) then
			dblDekningsbidrag = dblDekningsbidrag + rsDB("DB").value
		end if

		'Calculate overtime, hours, salary and db for current departmentarea
		strSQL = "EXEC GET_OVERTIME_FOR_CUSTOMERANDDEPARTMENT "  & strFraDatoAarsuke & ", " & strTilDatoAarsuke & ", " & lngCurrentFirmaID & "," & lngTomID
		Set rsOvertid = GetFirehoseRS(strSQL, Conn)

		if (HasRows(rsOvertid)) then
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

		'Render sum for this departmentarea 
		%>
			<tr>
				<td><%=rsRapport("tjenesteomrade")%></td>
				<td class="right"><%=FormatNumber( dblOmsetningSum, 2)%></td>
				<td class="right"><%=FormatNumber( dblDekningsbidrag, 2)%></td>
				<td class="right"><%=FormatNumber( dblSalarySum, 2)%></td>
				<td class="right"><%=rsRapport("antallOppdrag")%></td>
				<td class="right"><%=FormatNumber( dblFactor, 2)%></td>
			</tr>
		<%
		'Add to firm total
		intNOFFirmRecords		= intNOFFirmRecords + 1
		dblFirmOmsetningTotal = dblFirmOmsetningTotal + dblOmsetningSum
		dblFirmDekningsbidrag = dblFirmDekningsbidrag + dblDekningsbidrag
		dblFirmSalarySum		= dblFirmSalarySum + dblSalarySum
		intFirmAntallOppdrag 	= intFirmAntallOppdrag + rsRapport("antallOppdrag")
		dblFirmFactorBase		= dblFirmFactorBase + dblFactor

		'Add to firm total
		intArrayIndex = getIndexByTomID(lngTomID)

		ATotTom(intArrayIndex,3) = ATotTom(intArrayIndex,3) + dblFirmOmsetningTotal
		ATotTom(intArrayIndex,5) = ATotTom(intArrayIndex,5) + dblFirmSalarySum
		ATotTom(intArrayIndex,4) = ATotTom(intArrayIndex,4) + dblFirmDekningsbidrag
		ATotTom(intArrayIndex,6) = ATotTom(intArrayIndex,6) + intFirmAntallOppdrag
	end sub

	sub RenderDepartmentAreasSum
		'Render then sum for all department areas for this firm
		%>
			<tfoot>
				<tr>
					<td>sum</td>
					<td class="right"><%=FormatNumber( dblFirmOmsetningTotal, 2)%></td>
					<td class="right"><%=FormatNumber( dblFirmDekningsbidrag, 2)%></td>
					<td class="right"><%=FormatNumber( dblFirmSalarySum, 2)%></td>
					<td class="right"><%=intFirmAntallOppdrag%></td>
					<td class="right">&nbsp;</td>
				</tr>
			</tfoot>
		<%
		
		dblTotalOmsetningTotal = dblTotalOmsetningTotal + dblFirmOmsetningTotal
		dblTotalDekningsbidrag = dblTotalDekningsbidrag + dblFirmDekningsbidrag
		dblTotalSalarySum = dblTotalSalarySum + dblFirmSalarySum
		intTotalAntallOppdrag = intTotalAntallOppdrag + intFirmAntallOppdrag
		
		intNOFFirmRecords		= 0
		dblFirmOmsetningTotal	= 0.0
		dblFirmDekningsbidrag	= 0.0
		dblFirmSalarySum		= 0.0
		intFirmAntallOppdrag 	= 0
		dblFirmFactorBase		= 0.0
	end sub

	sub RenderHovedOppforingTotals
		dim intArrayRunner
		dim intArrayCount
		dim dblDeptTempOmsetningTotal
		dim dblDeptTempDekningsbidrag
		dim dblDeptTempSalarySum
		dim intDeptTempAntallOppdrag
		dim dblDeptTempFactor

		%>
		<h2>Totalt for hovedoppføring <%=strCurrentHovedOppforing%></h2>
		<div class="listing">
			<TABLE CELLPADDING="0" CELLSPACING="1" class="reportTable" ID="Table3">
				<tr>
					<th>Tjenesteområde</th>
					<th>Omsetning</th>
					<th>Bidrag</th>
					<th>Lønn</th>
					<th>Oppdrag</th>
				</tr>
			<%
			intArrayCount = ubound(ATotTom)
			intArrayRunner = 1
			while (intArrayRunner <= intArrayCount)
				%>
						<TR>
							<td><%=ATotTom(intArrayRunner,2)%></td>
							<td class="right"><%=FormatNumber( ATotTom(intArrayRunner,3), 2)%></td>
							<td class="right"><%=FormatNumber( ATotTom(intArrayRunner,4), 2)%></td>
							<td class="right"><%=FormatNumber( ATotTom(intArrayRunner,5), 2)%></td>
							<td class="right"><%=ATotTom(intArrayRunner,6)%></td>
						</TR>


					<%
					'Add to totals og all service areas for this firm
					dblDeptTempOmsetningTotal 	= dblDeptTempOmsetningTotal + ATotTom(intArrayRunner, 3)
					dblDeptTempDekningsbidrag 	= dblDeptTempDekningsbidrag + ATotTom(intArrayRunner, 4)
					dblDeptTempSalarySum		= dblDeptTempSalarySum + ATotTom(intArrayRunner, 5)
					intDeptTempAntallOppdrag 	= intDeptTempAntallOppdrag + ATotTom(intArrayRunner, 6)

					'Add to all totals
					ATotAllToms(intArrayRunner,3) = ATotAllToms(intArrayRunner,3) + ATotTom(intArrayRunner,3)
					ATotAllToms(intArrayRunner,4) = ATotAllToms(intArrayRunner,4) + ATotTom(intArrayRunner,4)
					ATotAllToms(intArrayRunner,5) = ATotAllToms(intArrayRunner,5) + ATotTom(intArrayRunner,5)
					ATotAllToms(intArrayRunner,6) = ATotAllToms(intArrayRunner,6) + ATotTom(intArrayRunner,6)	
		
					'Reset values for all service areas for this firm
					ATotTom(intArrayRunner,3) = 0.0
					ATotTom(intArrayRunner,4) = 0.0
					ATotTom(intArrayRunner,5) = 0.0
					ATotTom(intArrayRunner,6) = 0
					intArrayRunner = intArrayRunner + 1
			wend
			%>
				<tfoot>
				<TR>
					<td>Sum alle tjenesteområder</td>
					<td class="right"><%=FormatNumber( dblDeptTempOmsetningTotal, 2)%></td>
					<td class="right"><%=FormatNumber( dblDeptTempDekningsbidrag, 2)%></td>
					<td class="right"><%=FormatNumber( dblDeptTempSalarySum, 2)%></td>
					<td class="right"><%=intDeptTempAntallOppdrag%></td>
				</TR>
				</tfoot>
		</table>
	</div>
	<br>
	<%
	end sub
		
	sub RenderPageTotals
		dim intArrayRunner
		dim intArrayCount
		dim dblDeptTempOmsetningTotal
		dim dblDeptTempDekningsbidrag
		dim dblDeptTempSalarySum
		dim intDeptTempAntallOppdrag
		dim dblDeptTempFactor
		%>

		<h2>Totalt for alle tjenesteområder for alle hovedoppføringer</h2>
		<div class="listing">
		<TABLE CELLPADDING="0" CELLSPACING="1" class="reportTable" ID="Table1">
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
						<td class="right"><%=FormatNumber( ATotAllToms(intArrayRunner,3), 2)%></td>
						<td class="right"><%=FormatNumber( ATotAllToms(intArrayRunner,4), 2)%></td>
						<td class="right"><%=FormatNumber( ATotAllToms(intArrayRunner,5), 2)%></td>
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
					<td class="right"><%=FormatNumber( dblDeptTempOmsetningTotal, 2)%></td>
					<td class="right"><%=FormatNumber( dblDeptTempDekningsbidrag, 2)%></td>
					<td class="right"><%=FormatNumber( dblDeptTempSalarySum, 2)%></td>
					<td class="right"><%=intDeptTempAntallOppdrag%></td>
				</tr>
			</tfoot>
	</TABLE>
	</div>
	<%
	end sub

	function getIndexByTomID(lngTomID)
		dim intArrayRunner
		dim intArrayCount
		dim blnExitLoop

		blnExitLoop = false

		intArrayCount = ubound(ATotTom)
		intArrayRunner = 1
		while (intArrayRunner <= intArrayCount OR blnExitLoop = false)
			if cint(ATotTom(intArrayRunner,1)) = cint(lngTomID) then
				getIndexByTomID = intArrayRunner
				blnExitLoop = true
			end if
			intArrayRunner = intArrayRunner + 1
		wend
	end function
	
	
	if Request("POSTED") = 1 then

		strFradato  		= Request.Form( "tbxFradato" )
		strTildato  		= Request.Form( "tbxTildato" )
		strSelectedOppf		= Request.form( "chkOppforing")
		strSQLSelectedOppf	= strSelectedOppf

		if strFradato <> "" then
			strFraDatoUke = strFradato
			strFraDatoAar = strFradato
			call KorrigerUke(strFraDatoUke)
			call KorrigerAar(strFraDatoAar)
			strFraDatoAarsuke = strFraDatoAar & strFraDatoUke
			strFradato     = DbDate(strFradato)
		end if

		if strTildato <> "" then
			strTilDatoUke = strTildato
			strTilDatoAar = strTildato
			call KorrigerUke(strTilDatoUke)
			call KorrigerAar(strTilDatoAar)
			strTilDatoAarsuke = strTilDatoAar & strTilDatoUke
			strTildato     = DbDate(strTildato)
		end if
	end if

	Set Conn = GetConnection(GetConnectionstring(XIS, ""))
	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"
    "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script language="javascript" src="../Js/javascript.js" type="text/javascript"></script>
		<title>Omsetning pr. hovedoppføring</title>
		<script language="javaScript" type="text/javascript">
		<!--

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
				for(i=0;i<document.formEn.chkOppforing.length;i++)
					{
						document.formEn.chkOppforing[i].checked = true;
					}


				document.all.chkOppforing.checked = true;
			};

			function deSelectAll()
			{
				for(i=0;i<document.formEn.chkOppforing.length;i++)
					{
						document.formEn.chkOppforing[i].checked = false;
					}
			};
		//-->
		</script>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Omsetning pr. hovedoppføring/fagområde</h1>
			</div>
			<div class="content">
			<p>
				gir kun riktig overtid hvis datointervallet er en hel måned.<br>
				Faktor i totalene er gjennomsnittlige.
			</p>
			<form name="formEn" ACTION="omsetning_HovedOppforingTjenesteomrade.asp" METHOD="POST" ID="Form1">
				<INPUT TYPE='HIDDEN' VALUE='1' NAME='POSTED' ID='POSTED'>
				<table ID="Table2">
				<tr>
					<td>Fra dato:</td>
					<td><INPUT NAME="tbxFraDato" TYPE='TEXT' SIZE='10' MAXLENGTH='10' Value="<%=Request.Form( "tbxFradato")%>" ONBLUR="dateCheck(this.form, this.name)" ID="Text1"> </td>
					<td>Til dato:</td>
					<td><INPUT NAME="tbxTilDato" TYPE='TEXT' SIZE='10' MAXLENGTH='10' Value="<%=Request.Form( "tbxTildato" )%>" ONBLUR="dateCheck(this.form, this.name)" ID="Text2"> </td>
					<td>Hovedoppføring:</td>
					<td>
						<DIV class="container">
    					<INPUT class="checkbox" TYPE='CHECKBOX' ID='chkOppfAll' NAME='chkOppfAll' onClick='javascript:DeSelect(this)' VALUE=' '>Alle<br>
						<%
						dim strSel			'Contains "SELECTED" if current hoved oppføring is selected
						dim rsHovedOpp		'Recordset containing hoved oppføring
						dim strSQL			'SQL statement to EXECute, retriving coworkers
						dim strBgColor
						dim NofRecords
						
		   				strSQLSelectedOppf = "," & strSQLSelectedOppf & ","

		   				strSQL = "SELECT [FirmaID], [Firma] FROM [Firma] WHERE [ErHovedOppforing] = 1 ORDER BY [Firma]"
		   				Set rsHovedOpp = GetFirehoseRS(strSQL, Conn)

						strSQLSelectedOppf = replace(strSQLSelectedOppf," ","")

						Do Until (rsHovedOpp.EOF)
							strSel = ""
							if (strSQLSelectedOppf <> "") then
								if instr(strSQLSelectedOppf,"," & rsHovedOpp("FirmaID") & ",") then
									strSel = "Checked"
								end if
							end if
						%>
							<INPUT class="checkbox" TYPE="checkbox" <%=strSel%> ID="chkOppforing" NAME="chkOppforing" VALUE="<%=rsHovedOpp("FirmaID")%>"><%=rsHovedOpp("Firma")%><br>
						<%
							rsHovedOpp.MoveNext
							NofRecords = NofRecords + 1
						Loop

						' Close and release recordset
						rsHovedOpp.Close
						Set rsHovedOpp = Nothing
						%>
						</DIV>
   					</td>
					<td><input type="submit" name="pbnDataAction" value="Søk" ID="Submit1"></td>
				</tr>
			</table>
			</form>
			<%
				lngPrevAvd = 0

				IF Request.Form("POSTED") = 1 then

					strSQL = "SELECT COUNT([tomID]) AS NofRec FROM [Tjenesteomrade]"
					Set rsToms = Conn.execute(strSQL, 1)
					intTotToms = rsToms.fields("NofRec")
					redim ATotTom(intTotToms,7)
					redim ATotAllToms(intTotToms,7)

					strSQL = "SELECT [tomID], [Navn] FROM [Tjenesteomrade] ORDER BY [tomID]"
					Set rsToms = Conn.Execute(strSQL,1)
					if (NOT rsToms.EOF) then
						intTOMTeller = 1
						while not rsToms.EOF
							ATotTom(intTOMTeller,1) = rsToms.fields("tomID").value
							ATotTom(intTOMTeller,2) = rsToms.fields("navn").value
							ATotTom(intTOMTeller,3) = 0.0
							ATotTom(intTOMTeller,4) = 0.0
							ATotTom(intTOMTeller,5) = 0.0
							ATotTom(intTOMTeller,6) = 0
							ATotTom(intTOMTeller,7) = 0.0
							
							ATotAllToms(intTOMTeller,1) = ATotTom(intTOMTeller,1)
							ATotAllToms(intTOMTeller,2) = ATotTom(intTOMTeller,2)
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

					dblTotalOmsetningTotal = 0.0
					dblTotalDekningsbidrag = 0.0
					dblTotalSalarySum = 0.0
					intTotalAntallOppdrag = 0

					strSQL = "EXEC GET_OMSETNINGDEPARTMENT_BY_MAINCUSTOMER " & strFraDato & ", " & strTilDato & ",'" & strSQLSelectedOppf & "'"
					Set rsRapport = GetFirehoseRS(strSQL, Conn)
					if (HasRows(rsRapport)) then
						NofRecords = 0
						lngPrevFirmID = 0
						lngPrevSubID = 0
						%>
						</div>
						<div class="contentHead">
							<h2>Rapport</h2>
						</div>
						<div class="content">
						<%
						while not (rsRapport.EOF)
							'New underoppføring
							if ( (lngPrevSubID > 0) AND (lngPrevSubID <> clng(rsRapport("FirmaID"))) )  then
								call RenderDepartmentAreasSum()
								%>
								</Table>
								<br>
								<%
							end if
							'Break on new hovedoppføring
							if ( (lngPrevFirmID <> clng(rsRapport("HovedOppforingID"))) )  then
								if (lngPrevSubID > 0)  then
									call RenderHovedOppforingTotals()
								end if
								strCurrentHovedOppforing = rsRapport("HovedOppforing")					
								%>					
									<h1><%=strCurrentHovedOppforing%></h1>
								<%
							end if
							if ( (lngPrevSubID <> clng(rsRapport("FirmaID"))) )  then
								%>
								<h3><%=rsRapport("Firma")%></h3>
								<div class="listing">
								<table cellpadding='0' cellspacing='1' class="reportTable" ID="Table4">
									<tr>
											<th>Tjenesteområde</th>
											<th>Omsetning</th>
											<th>Bidrag</th>
											<th>Lønn</th>
											<th>Oppdrag</th>
											<th>Faktor</th>
									</tr>
								<%
							end if				
							Call RenderDepartmentAreaTotals()
							
							'Prepare for next loop
							lngPrevFirmID = clng(rsRapport("HovedOppforingID"))
							lngPrevSubID = clng(rsRapport("FirmaID"))
							NofRecords = NofRecords + 1
							rsRapport.movenext
						wend
					end if
					call RenderDepartmentAreasSum()
					%>
					</Table>
					<%		
					call RenderHovedOppforingTotals()
					call RenderPageTotals()
					rsRapport.close
					set rsRapport 	= nothing
					set rsDB 		= nothing
					set rsOvertid 	= nothing
				end if
			%>
			<p id="sideskift" class="pageBreakAfter">&nbsp;</p>
 			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>