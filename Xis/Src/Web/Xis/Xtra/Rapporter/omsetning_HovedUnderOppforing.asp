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
<!--#INCLUDE FILE="..\includes\Xis.Excel.Utils.inc"-->
<%
	If (HasUserRight(ACCESS_REPORT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim Conn			'Connection
	dim lngPrevAvd
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
	dim strSelectedOppf
	dim strSQLSelectedAvd

	dim xlsResult

	dim dblFactor
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
	dim rsFirma
	dim firmLink	

	dim previousIDs : previousIDs = ""
	dim previousHIDs  : previousHIDs = ""

	sub CalculateAndRenderSums(rsRapport, level, strFraDato, strTilDato, strFraDatoAarsuke, strTilDatoAarsuke)
		dim dblOmsetningSum
		dim dblOmsetningHours
		dim dblSalarySum
		dim dblSalaryHours
		dim dblDekningsbidrag
		dim dblDekningsbidragOvertid
		dim lngCurrentFirmaID
		dim inttypeID
		dim dblOvertimeSum
		dim lngOvertidProsent

		lngCurrentFirmaID = rsRapport("FirmaID")

		'Get DB for this Firmleader's current service area
		dblOmsetningSum 			= rsRapport("omsetning")
		dblOmsetningHours 			= rsRapport("TotFakturaTimer")
		dblSalarySum 	 			= rsRapport("loenn")
		dblSalaryHours 	 			= rsRapport("TotAntTimer")
		dblDekningsbidrag			= 0.00
		dblDekningsbidragOvertid	= 0.00

		'Calculate db for current firm, of type "1"

		strSQL = "EXEC GET_DB_FOR_CUSTOMER " & strFraDato & ", " & strTilDato & "," & lngCurrentFirmaID & ", 1"
'Response.Write "<tr><td colspan='6'>" & strSQL & "</td></tr>"
		Set rsDB = GetFirehoseRS(strSQL, Conn)
		if (NOT isnull(rsDB("DB").value)) then
			dblDekningsbidrag = rsDB("DB").value
		end if
		set rsDB  = nothing

		'Calculate db for current firm, of all other types
		strSQL = "EXEC GET_DB_FOR_CUSTOMER " & strFraDato & ", " & strTilDato & "," & lngCurrentFirmaID & ", 0"
'Response.Write "<tr><td colspan='6'>" & strSQL & "</td></tr>"
		Set rsDB = GetFirehoseRS(strSQL, Conn)
		if (NOT isnull(rsDB("DB").value)) then
			dblDekningsbidrag = dblDekningsbidrag + rsDB("DB").value
		end if
		set rsDB  = nothing

		'Calculate overtime, hours, salary and db for current firm
		strSQL = "EXEC GET_OVERTIME_FOR_CUSTOMER "  & strFraDatoAarsuke & ", " & strTilDatoAarsuke & ", " & lngCurrentFirmaID
'Response.Write "<tr><td colspan='6'>" & strSQL & "</td></tr>"
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
		If dblOmsetningSum > 0 and (dblDekningsbidrag > 0) then
			dblFactor = dblOmsetningSum / ((dblOmsetningSum - dblDekningsbidrag) / XIS_FACTOR)
		else
			dblFactor = 0.0
		end if

		dim influx
		
		influx = Replace(Space(level * 3), " ", "&nbsp;")

		strSQL = "SELECT [SOCuID] FROM [Firma] WHERE [FirmaID] = " & rsRapport("FirmaID")
		set rsFirma = GetFirehoseRS(strSQL, Conn)
		if(isnull(rsFirma("SOCuID"))) then
			firmLink = rsRapport("Firma")
		else
			firmLink = CreateSONavigationLink(SUPEROFFICE_PANEL_CONTACT_URL, SUPEROFFICE_PANEL_CONTACT_URL, rsFirma("SOCuID"), rsRapport("Firma"), "Vis kunde '" & rsRapport("Firma") & "'")		
		end if
		rsFirma.Close
		set rsFirma = nothing
		
		'Render sum for this firm
		%>
			<tr>
				<td><%=influx%><%=firmLink%></td>
				<td class="right"><%=FormatNumber( dblOmsetningSum, 2)%></td>
				<td class="right"><%=FormatNumber( dblDekningsbidrag, 2)%></td>
				<td class="right"><%=FormatNumber( dblSalarySum, 2)%></td>
				<td class="right"><%=rsRapport("antallOppdrag")%></td>
				<td class="right"><%=FormatNumber( dblFactor, 2)%></td>
			</tr>
		<%
		'Add to firm total
		intNOFFirmRecords		= intNOFFirmRecords + 1
		dblFirmOmsetningTotal   = dblFirmOmsetningTotal + dblOmsetningSum
		dblFirmDekningsbidrag   = dblFirmDekningsbidrag + dblDekningsbidrag
		dblFirmSalarySum		= dblFirmSalarySum + dblSalarySum
		intFirmAntallOppdrag 	= intFirmAntallOppdrag + rsRapport("antallOppdrag")
		dblFirmFactorBase		= dblFirmFactorBase + dblFactor

		'Add to excel result
		xlsResult = xlsResult & ToCSVString(rsRapport("FirmaID"), false) & ";"
		xlsResult = xlsResult & ToCSVString(rsRapport("Firma"), false) & ";"
		xlsResult = xlsResult & ToCSVString( dblOmsetningSum, false) & ";"
		xlsResult = xlsResult & ToCSVString( dblDekningsbidrag, false) & ";"
		xlsResult = xlsResult & ToCSVString( dblSalarySum, false) & ";"
		xlsResult = xlsResult & ToCSVString(rsRapport("antallOppdrag"), false) & ";"
		xlsResult = xlsResult & ToCSVString( dblFactor, false) & vbcrlf

		if( not IsNull(rsRapport("hovedoppforingID"))  ) then

			dim rsSubRapport
			dim id
			dim Hid

			Hid = "|" & trim(lngCurrentFirmaID) & "|"
'Response.Write "<tr><td colspan='6'>previousHIDs:" & previousHIDs & "</td></tr>"
'Response.Write "<tr><td colspan='6'>instr:" & instr(1, previousHIDs, Hid) & "</td></tr>"
			if(instr(1, previousHIDs, Hid) = 0) then
				previousHIDs = previousHIDs & Hid
				strSQL = "EXEC GET_OMSETNING_BY_MAINCUSTOMER_ALL " & strFraDato & ", " & strTilDato & ",'" & lngCurrentFirmaID & "'"
'Response.Write "<tr><td colspan='6'>" & strSQL & "</td></tr>"
				Set rsSubRapport = GetFirehoseRS(strSQL, Conn)

				if (HasRows(rsSubRapport)) then
					while not (rsSubRapport.EOF)
						id = "|" & trim(rsSubRapport("FirmaID")) & "|"
'Response.Write "<tr><td colspan='6'>previousIDs:" & previousIDs & "</td></tr>"
'Response.Write "<tr><td colspan='6'>instr:" & instr(1, previousIDs, id) & "</td></tr>"
'Response.Write "<tr><td colspan='6'>id:" & id & "</td></tr>"
						if(instr(1, previousIDs, id) = 0) then
							call CalculateAndRenderSums(rsSubRapport, level + 1, strFraDato, strTilDato, strFraDatoAarsuke, strTilDatoAarsuke)
						end if
						previousIDs = previousIDs & id
						rsSubRapport.movenext
					wend
				end if
				rsSubRapport.close
				set rsSubRapport = nothing
			end if
		end if
	end sub

	sub renderHovedOppforingTotals
			'Render total sum for this firm
			%>
			<tfoot>
				<tr>
					<td>Totalt for firma</td>
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

			intNOFFirmRecords       = 0
			dblFirmOmsetningTotal   = 0.0
			dblFirmDekningsbidrag   = 0.0
			dblFirmSalarySum		= 0.0
			intFirmAntallOppdrag 	= 0
			dblFirmFactorBase		= 0.0
	end sub


	sub RenderPageTotals
		%>
		<br>
		<table cellpadding='0' cellspacing='1' class="reportTable" ID="Table1">
			<tr>
					<th>Totalt</th>
					<th>Omsetning</th>
					<th>Bidrag</th>
					<th>Lønn</th>
					<th>Oppdrag</th>
			</tr>
			<tfoot>
			<tr>
				<td>Totalt alle oppføringer</td>
				<td class="right"><%=FormatNumber( dblTotalOmsetningTotal, 2)%></td>
				<td class="right"><%=FormatNumber( dblTotalDekningsbidrag, 2)%></td>
				<td class="right"><%=FormatNumber( dblTotalSalarySum, 2)%></td>
				<td class="right"><%=intTotalAntallOppdrag%></td>
			</tr>
			</tfoot>
		</Table>
		<%
	end sub

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
			strTildato		  = DbDate(strTildato)
		end if
	end if

	Set Conn = GetConnection(GetConnectionstring(XIS, ""))
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
		<script language="javascript" src="../Js/javascript.js" type="text/javascript"></script>
		<title>Omsetning pr. hovedoppføring</title>
		<script language="javaScript" type="text/javascript">
		<!--
			function DeSelect(objAll)
			{

				if (objAll.checked == true)
				{
					selectAll()
				}
				else
				{
					deSelectAll()
				}
			}

			function selectAll()
			{
				for(i = 0; i < document.formEn.chkOppforing.length; i++)
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
				<h1>Omsetning pr. hovedoppføring / underoppføring</h1>
			</div>
			<div class="content">
			<p>
				gir kun riktig overtid hvis datointervallet er en hel måned.<br>
				Faktor i totalene er gjennomsnittlige.
			</p>
			<form name="formEn" ACTION="omsetning_HovedUnderOppforing.asp" METHOD="POST" ID="Form1">
				<INPUT TYPE='HIDDEN' VALUE='1' NAME='POSTED' ID='POSTED'>
				<table ID="Table2">
				<tr>
					<td>Fra dato:</td>
					<td><INPUT NAME="tbxFraDato" TYPE='TEXT' SIZE='10' MAXLENGTH='10' Value="<%=Request.Form( "tbxFradato")%>" ONBLUR="dateCheck(this.form, this.name)" ID="Text1"> </td>
					<td>Til dato:</td>
					<td><INPUT NAME="tbxTilDato" TYPE='TEXT' SIZE='10' MAXLENGTH='10' Value="<%=Request.Form( "tbxTildato" )%>" ONBLUR="dateCheck(this.form, this.name)" ID="Text2"> </td>
					<td>Rotoppføring:</td>
					<td>
						<DIV class="container">
    					<INPUT class="checkbox" TYPE='CHECKBOX' ID='chkOppfAll' NAME='chkOppfAll' onClick='javascript:DeSelect(this)' VALUE=' '>Alle<br>
						<%
						dim strSel			'Contains "SELECTED" if current hoved oppføring is selected
						dim rsHovedOpp		'Recordset containing hoved oppføring
						dim strSQL			'SQL statement to execute, retriving coworkers
						dim strBgColor

						dim level : level = 0

		   				strSQLSelectedOppf = "," & strSQLSelectedOppf & ","

						strSQL = "SELECT [FirmaID], [Firma] FROM [Firma] WHERE [ErHovedOppforing] = 1 AND [HovedOppforingId] = [FirmaId] ORDER BY [Firma]"
		   				Set rsHovedOpp = getFirehoseRS(strSQL, Conn)

						strSQLSelectedOppf = replace(strSQLSelectedOppf," ","")

						Do Until (rsHovedOpp.EOF)
							strSel = ""
							if (strSQLSelectedOppf <> "") then
								if instr(strSQLSelectedOppf, "," & rsHovedOpp("FirmaID") & ",") then
									strSel = "Checked"
								end if
							end if
							%>
							<INPUT class="checkbox" TYPE="checkbox" <%=strSel%> ID="chkOppforing" NAME="chkOppforing" VALUE="<%=rsHovedOpp("FirmaID")%>"><%=rsHovedOpp("Firma")%><br>
							<%
							rsHovedOpp.MoveNext
						Loop

						' Close and release recordset
						rsHovedOpp.Close
						Set rsHovedOpp = Nothing
						%>
						</DIV>
   					</td>
					<TD><input type="submit" name="pbnDataAction" value="Søk" ID="Submit1"></td>
				</tr>
			</table>
			</form>
			<%
				dim strFeedback
				lngPrevAvd = 0

				if Request.Form("POSTED") = 1 then

					if(len(strFraDato) = 0 or Len(strTilDato) = 0 ) then
						strFeedback = "Du må skrive inn fra- og tildato."
					else
						dblTotalOmsetningTotal = 0.0
						dblTotalDekningsbidrag = 0.0
						dblTotalSalarySum = 0.0
						intTotalAntallOppdrag = 0
						strSQL = "EXEC GET_OMSETNING_BY_MAINCUSTOMER_TOP " & strFraDato & ", " & strTilDato & ",'" & strSQLSelectedOppf & "'"
						Set rsRapport = GetFirehoseRS(strSQL, Conn)
						lngPrevFirmID = 0
						if (HasRows(rsRapport)) then
							%>
							</div>
							<div class="contentHead">
								<h2>Rapport</h2>
							</div>
							<div class="content">
								<a href="rapportExport.asp?report=hovedunder" target="_blank"><img src="../images/icon_xls.gif" align="absmiddle"> Åpne csv-fil</a>
							<%
							xlsResult = xlsResult & ToCSVString("KundeID", false) & ";"
							xlsResult = xlsResult & ToCSVString("Kunde", false) & ";"
							xlsResult = xlsResult & ToCSVString("Omsetning", false) & ";"
							xlsResult = xlsResult & ToCSVString("Bidrag", false) & ";"
							xlsResult = xlsResult & ToCSVString("Lønn", false) & ";"
							xlsResult = xlsResult & ToCSVString("Oppdrag", false) & ";"
							xlsResult = xlsResult & ToCSVString("Faktor", false) & vbcrlf

							while not (rsRapport.EOF)
								'New firm..
								if ( (lngPrevFirmID <> clng(rsRapport("HovedOppforingID"))) )  then
									if (lngPrevFirmID <> 0) then
										renderHovedOppforingTotals()
										response.write "</table></div>"
										xlsResult = xlsResult & ";;;;;;" & vbcrlf
									end if
								%>
								<h2><%=rsRapport("HovedOppforing")%></h2>
								<div class="listing">
								<table border='1' cellpadding='0' cellspacing='1' class="reportTable" ID="Table3">
								<tr>
										<th>Kunde</th>
										<th>Omsetning</th>
										<th>Bidrag</th>
										<th>Lønn</th>
										<th>Oppdrag</th>
										<th>Faktor</th>
								</tr>
								<%
								end if
								id = "|" & trim(rsRapport("HovedOppforingID")) & "|"
								if(instr(1, previousIDs, id) = 0) then
									previousIDs = previousIDs & id
									call CalculateAndRenderSums(rsRapport, level, strFraDato, strTilDato, strFraDatoAarsuke, strTilDatoAarsuke)
								end if

								lngPrevFirmID = clng(rsRapport("HovedOppforingID"))

								rsRapport.movenext
							wend
							rsRapport.close
						end if
						call renderHovedOppforingTotals()
						response.write "</table>"
						call RenderPageTotals()
						response.write "</div>"
						session("hovedunder") = xlsResult
						set rsRapport 	= nothing
						set rsDB 		= nothing
						set rsOvertid 	= nothing
					end if
				end if
				%>
				<p id="sideskift" class="pageBreakAfter">&nbsp;</p>
				<%=strFeedback%>
 			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>