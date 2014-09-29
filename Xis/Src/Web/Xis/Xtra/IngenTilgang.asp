<%@ LANGUAGE="VBSCRIPT" %>
<%Option explicit%>
<%
	dim referer
	dim hasReferer : hasReferer = false
	dim mailTo : mailTo = "mailto:christian@xtra.no"
	
	if len(Request.ServerVariables("HTTP_REFERER")) > 0 then
		referer = mid(Request.ServerVariables("HTTP_REFERER"), InStrRev(Request.ServerVariables("HTTP_REFERER"), "/") + 1)
		hasReferer = true
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<title>Ikke tilgang til ressursen</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Ikke tilgang</h1>
			</div>			
			<div class="content">
				<p>Du har ikke tilgang til denne ressursen.<br>
					<%
					if (hasReferer) then
						mailTo = "mailto:christian@xtra.no?subject=" & "Vedr. tilgang til siden '" & referer & "'"
					end if
					%>
					Kontakt <a href="<%=mailTo%>">Christan Willoch</a> for nærmere informasjon.<br>
					<%
					if (session("debug") = "true") then
						%>
						<a href="/xtra/admin/debug.asp">Debug informasjon</a>
						<%
					end if
					%>
				</p>
			</div>
		</div>
	</body>
</html>
