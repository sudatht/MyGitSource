<% profil = Session("Profil") %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <meta http-equiv="Content-Style-Type" content="text/css">
    <meta http-equiv="Content-Script-Type" content="text/javascript">
    <meta name="Developer" content="Electric Farm ASA">
    <title></title>
    <link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
	<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
</head>
	<body>
		<div class="pageContainer" id="pageContainer">
		<div class="contentHead1">
			<H1>Admin</H1>
		</div>
		<div class="content">
		<p>
		<% 
		If  Mid(profil,6,1) > 2 Then 
			%>
			<a href="tilgang_frame.asp">Rettigheter</a> | 
			<% 
		End If 
		If  Mid(profil,6,3) > 2 Then 
			%>
			<a href="../internettAdmin/tilgangadmin.asp">Internett-tilgang</a> | 
			<a href="Omsetning-ansvarlig.asp">Omsetning pr. ansvarlig medarbeider ( fakturert )</a>
		<% 
		End If 
		%>
		</p>
	</div>
    </div>
</body>
</html>
