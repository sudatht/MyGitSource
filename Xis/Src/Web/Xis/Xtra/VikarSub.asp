<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"  "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<title>VikarSubMeny</title>
		<meta name="Microsoft Border" content="none">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script language='javascript' src='/xtra/js/menu.js' id='menuScripts'></script>		
		<script language="javaScript" type="text/javascript">
		
			function openMaximus(){
				window.open("http://xposlapp3/Site/Search/SearchMain.aspx");
			}		
		
			function shortKey(e) 
			{
				var keyChar = String.fromCharCode(event.keyCode);
				var modKey  = event.ctrlKey;
				var modKey2 = event.shiftKey;
			
				<% 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>				
					if (modKey && modKey2 && keyChar == "S")
					{	
						parent.frames[funcFrameIndex].location=("VikarSoek.asp");
					}
					<% 
				End If 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
					%>					
					if (modKey && modKey2 && keyChar == "Y")
					{	
						parent.frames[funcFrameIndex].location=("VikarNy.asp");
					}
					<% 
				End If 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>					
					if (modKey && modKey2 && keyChar == "W")
					{	
						parent.frames[funcFrameIndex].location = ("SuspectList.asp");
					}
					<% 
				End If 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>					
					if (modKey && modKey2 && keyChar == "M")
					{	
						parent.frames[funcFrameIndex].location = ("http://xposlapp3/Site/Search/SearchMain.aspx");
					}
					<% 
				End If 
				%>					
			}
			//Her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;
		</script>
		<base target="bottom">
	</head>
	<body class="navMenuBody">
		<div class="navMenu">
			<p>
				<% 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>		
					<a href="VikarSoek.asp"><u>S</u>øk</a>
					<% 
				End If 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
					%>
					<a href="VikarDuplikatSoek.asp">N<U>y</U> vikar</a>
					<% 
				End If 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>
					<a href="jobb/SuspectList.asp">Søkere fra <u>w</u>eb</a>
					<% 
				End If 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>
					<a style="cursor:hand;" onclick="openMaximus();">Search <u>M</u>aximus</a>
					<% 
				End If 
				%>
			</p>
		</div>
	</body>
</html>