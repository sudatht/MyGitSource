<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<% 
If (HasUserRight(ACCESS_ADMIN, RIGHT_READ) = false) Then 
	call Response.Redirect("/xtra/IngenTilgang.asp")	
End If 
%>
<html>
	<head>
		<title>Timelister</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Timelister</h1>
			</div>
			<div class="content content2">
				<ul>
					<li><a href="~\..\..\WebUI\Admin\TimeSheets\TimeSheetListing.aspx">Behandling av timelister</a></li>
					<!--<% 
					If (HasUserRight(ACCESS_ADMIN, RIGHT_WRITE) = true) Then
						%>
						<li><a href="Timeliste_lag_gamle_01.asp">Fjerning av gamle timelister fra visning</a></li>
						<% 
					End If 
					%>-->					
					<li><a href="Vikar_timeliste_list_gml.asp?viskode=0">Søk i gamle timelister</a></li>					
				</ul>
				<!-- <p><a href="Rutiner.html" target="_new">Rutinebeskrivelse</a></p> -->
			</div>
		</div>
	</body>
</html>