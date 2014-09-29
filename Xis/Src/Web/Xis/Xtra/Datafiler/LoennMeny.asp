<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<% 
If (HasUserRight(ACCESS_ADMIN, RIGHT_READ) = false) Then 
	call Response.Redirect("/xtra/IngenTilgang.asp")	
End If 
%>
<html>
	<head>
		<title>Lønn</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Lønn</h1>
			</div>
			<div class="content content2">
				<ul>
					<li><a href="~\..\..\WebUI\Admin\Salary\TransferSalary.aspx">Kjør lønn</a></li>
					<li><a href="~\..\..\WebUI\Admin\Salary\TransferVikar.aspx">Nye og endrede vikarer</a></li>						
					<li><a href="../reports/AccruedReportByDepartment.aspx?vikarTypeID=1">Periodiseringsrapport - vanlig</a></li>
					<li><a href="../reports/AccruedReportByDepartment.aspx?vikarTypeID=3">Periodiseringsrapport - AS</a></li>					
					<!--<li><a href="As_vis.asp">Endre lønnstatus på AS - vikarer</a></li>-->
					<li><a href="Loen_kontroll_vis.asp">Kontrollrapport - Lønn</a></li>
					
					
				</ul>
				<!--<p><a href="Rutiner.html" target="_new">Rutinebeskrivelse</a></p>-->
			</div>
		</div>
	</body>
</html>