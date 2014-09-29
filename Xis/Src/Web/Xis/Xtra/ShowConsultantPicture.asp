<%option explicit%>
<%
	dim lngVikarID
	
	lngVikarID = trim(request("vikarID"))
	if len(lngVikarID)=0 then
		response.Write "VikarID mangler!"
		Response.End
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<title>Viser bilde av vikar</title>
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead">
				<h1>Vikar foto</h1>
			</div>
			
			<div class="content content2">
				<img src="<%=Application("URLConsultantImages") & lngVikarID & ".jpg"%>">
			</div>
		</div>
	</body>
</html>
