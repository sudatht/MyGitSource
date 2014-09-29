<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<% 
If (HasUserRight(ACCESS_ADMIN, RIGHT_READ) = false) Then 
	call Response.Redirect("/xtra/IngenTilgang.asp")	
End If 
%>
<html>
	<head>
		<title>Faktura</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<H1>Faktura</H1>
			</div>
			<div class="content content2">
				<ul>
					<li><a href="../WebUI/Admin/Invoice/InvoiceList.aspx">Behandle fakturagrunnlag</a></li>
					<li><a href="../reports/AccruedReportByDepartment.aspx?vikarTypeID=1">Periodiseringsrapport - vanlig</a></li>
					<li><a href="../reports/AccruedReportByDepartment.aspx?vikarTypeID=3">Periodiseringsrapport - AS</a></li>
					<li><a href="Faktura_kontroll_vis.asp">Kontrollrapport - faktura</a></li>
				</ul>
				<!-- <p><a href="Rutiner.html" target="_new">Rutinebeskrivelse</a></p> -->
			</div>
		</div>
	</body>
</html>