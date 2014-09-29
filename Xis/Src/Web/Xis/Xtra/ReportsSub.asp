<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<title>vikarSubMeny</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script language='javascript' src='js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='js/navigation.js' id='navigationScripts'></script>
		<base target="bottom">
		<script language="javaScript" type="text/javascript">	
			function shortKey(e) 
			{	
				var keyChar = String.fromCharCode(event.keyCode);
				var modKey  = event.ctrlKey;
				var modKey2 = event.shiftKey;
				<% 
				If HasUserRight(ACCESS_TASK, RIGHT_READ) Then
					%>				
					if (modKey && modKey2 && keyChar=="S")
					{	
						parent.frames[funcFrameIndex].location=("/xtra/OppdragSoek.asp");
					}
					<% 
				End If 
				%>				
			}
			//her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;
		</script>
	</head>
	<body class="navMenuBody">
		<div class="navMenu">
			<% 
			If HasUserRight(ACCESS_TASK, RIGHT_READ) Then
				%>		
				<a href="rapporter\Rapporter.asp"><u>R</u>eports</a>
				<% 
			End If 
			%>		
		</div>
	</body>
</html>