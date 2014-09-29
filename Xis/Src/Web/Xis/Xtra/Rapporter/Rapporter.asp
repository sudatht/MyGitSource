<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<%
dim strSQL 'As String
dim rsPerm
dim bQuestbackPerm

	If (HasUserRight(ACCESS_REPORT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
%> 
<%
	'--------------------------------------------------------------------------------------------------
	' Connect to database
	'--------------------------------------------------------------------------------------------------
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
	Conn.CommandTimeOut = Session("xtra_CommandTimeout")
	Conn.Open Session("xtra_ConnectionString"),Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")
	
	strSQL = "SELECT Questback FROM BRUKER WHERE (ID = " &  Session("brukerID") & " )"
	
	Set rsPerm = conn.Execute(strSQL)
	If Not rsPerm.EOF Then
		bQuestbackPerm = rsPerm("Questback").Value
	Else 'ingen linjer 
		rsPerm.Close
		set rsPerm = nothing
	End If	

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<title>Rapporter</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script language='javascript' src='../js/navigation.js' id='navigationScripts'></script>
	</head>
	<body onLoad="fokus()">
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Rapporter</h1>
			</div>
			<div class="content content2">
			<table>
			<tr>
			<td>
			<ol>
					<li><a id="lnk1"  href="OmsetningVikar.asp">Omsetning pr. vikar ( fakturert )</a></li>
					<li><a id="lnk2"  href="OmsetningKunde.asp">Omsetning pr. kunde ( fakturert )</a></li>
					<li><a id="lnk3"  href="OmsetningAvdeling.asp">Omsetning pr. avdeling ( fakturert )</a></li>
					<li><a id="lnk4"  href="../reports/DepartmentTurnoverByResponsible.aspx">Omsetning pr. avdeling/Ansvarlig/Oppdrag</a></li>
					<li><a id="lnk5"  href="../reports/DepartmentTurnoverByCustomer.aspx">Omsetning pr. avdeling/Kunde/Oppdrag</a></li>
					<li><a id="lnk6"  href="../reports/RevenueReportByOffice.aspx">Omsetning pr. avdelingskontor/Ansvarlig/tjenesteområde</a></li>
					<li><a id="lnk7"  href="../reports/RevenueReportByDepartment.aspx">Omsetning pr. avdelingskontor/Ansvarlig/avdeling</a></li>
					<!--<li><a id="lnk7"  href="Omsetning_v4.asp">Omsetning pr. regnskapsavdeling/Kunde/Ansvarlig</a></li>
					<li><a id="lnk8"  href="omsetning_HovedUnderOppforing.asp">Omsetning pr. Hovedoppføring/Underoppføring</a></li>
					<li><a id="lnk9"  href="omsetning_HovedOppforingTjenesteomrade.asp">Omsetning pr. Hovedoppføring/Fagområde</a></li>-->
					<li><a id="lnk10"  href="OmsetningAnsvarlig.asp">Omsetning pr. ansvarlig medarbeider ( fakturert )</a></li>
					<li><a id="lnk11"  href="NyeVikarer.asp">Nye vikarer</a></li>
					<li><a id="lnk12"  href="foedselsdager.asp">Fødselsdager</a></li>
					<li><a id="lnk13"  href="FaktGrlagkundeStart.asp">Fakturagrunnlag kunde</a></li>
					<li><a id="lnk14"  href="loennFaktGrStart.asp">Lønn-/ fakturagrunnlag vikar</a></li>
					<!--<li><a id="lnk15"  href="estimertOmsetningAnsMed.asp">Estimert omsetning pr. avdeling/ansvarlig/oppdrag</a></li>
					<li><a id="lnk16"  href="estimertOmsetningKunder.asp">Estimert omsetning pr. kunder/ansvarlig/oppdrag</a></li>-->
					<li><a id="lnk17"  href="rapportSkattekort.asp">Manglende skattekort vikarer</a></li>
					<li><a id="lnk19"  href="../reports/ActivityList.aspx">Aktivitetsrapport for vikarledere</a></li>
					<li><a id="lnk20"  href="../reports/ActivityListforDepartmentLeader.aspx">Aktivitetsrapport for avdelingsledere</a></li>
					<li><a id="lnk21"  href="../reports/CustomerInquiriesByDepartment.aspx">Forespørsler etter avdeling</a></li>
					<li><a id="lnk22"  href="../reports/ListsForFAgreement.aspx">Omsetning - rammeavtaler</a></li>
					<li><a id="lnk23"  href="../reports/FAgreementExtendedReport.aspx">Analyse - rammeavtaler</a></li>
					<li><a id="lnk24"  href="../reports/CustomerTimesheetReport.aspx">Customer timesheet report</a></li>
				</ol>
			</td>
			<td>			
			<ol style='<% If(bQuestbackPerm) Then Response.Write("visibility:visible") Else  Response.Write("visibility:hidden") End If %>'>
				<li><a id="Qblnk1"  href="../reports/QuestbackReportPage.aspx">Questback - oppdrag</a></li>
				<li><a id="Qblnk2"  href="../reports/QuestbackCommOmitted.aspx">Questback - reservasjonsliste</a></li>
			</ol>
			</td>
				
				</tr>
				</table>
			</div>
		</div>
	</body>
</html>

