<%@ Language=VBScript %>
<%
	response.Clear
	if request("dest") = "Application" then
		Response.ContentType ="application/msword"
		Response.AddHeader "Content-Disposition", "inline; filename=fnu.doc"
	elseif request("dest") = "Disk" then
		Response.ContentType ="application/html"
		Response.AddHeader "Content-Disposition", "attachment; filename=fnark.html"
	elseif request("dest") = "Browser" then
		Response.ContentType ="application/html"
		Response.AddHeader "Content-Disposition", "inline; filename=fnark.html"
	end if

%>
<!doctype html public "-//w3c//dtd html 4.01 transitional//en" "http://www.w3.org/tr/html4/loose.dtd">
<html>
<head>
	<title>Curriculum Vitae</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<meta http-equiv="Content-Style-Type" content="text/css">
	<meta http-equiv="Content-Script-Type" content="text/javascript">
	<style>
		BODY		{font-size: .7em; margin: 0; padding: 0 10px 10px 10px;}
		BODY, TABLE	{font-family: tahoma, sans-serif;}
		H1, H2, H3	{margin:0 0 0 0; color:#666666; background:transparent;}
		H1			{font-size:2.4em; font-weight:normal;}
		H2			{font-size:1.5em; font-weight:normal;}
		P           {font-size:1em; margin:0 0 1em 0;}
		TABLE 		{width:100%;}
		IMG			{border:none;}
		.newWindow	{color:#000000; background:#ffffff; padding:0 2em 2em 2em;}
		.showCV		{margin:1.5em 0 0 0;}
		.newWindow .logo	{position:absolute; top:0; left:84%; width:63;}
		.showCV H1	{border-bottom:double #336666;}
		.showCV H2	{font-size: 1.6em; margin-top:.5em; /*border-bottom:1px solid #336666;*/}
		.showCV TD, .showCV TH	{border-bottom:1px solid #cccccc;}
		.showCV TABLE	{width:100%;}
		.showCV TR	{vertical-align:top;}
		.left		{text-align:left;}
		.center		{text-align:center;}
		.right		{text-align:right;}
		.warning	{color:#ff0000; background:transparent;}
		.top		{vertical-align:top;}
	</style>
</head>
	<body class="newWindow">
		<div class="showCV">
			<h1>TEST!</h1>
			Dette er no html! Testa testa!<br>
			<p>fun fnu!</p>
		</div>
	</body>
</html>