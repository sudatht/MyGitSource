<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Library.inc"-->
<%
profil = Session("Profil")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<title>hotlistSub</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script language='javascript' src='js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='js/navigation.js' id='navigationScripts'></script>
		<script type="text/javascript">
			//lager felles variabler
			function shortKey(e) 
			{
				var keyChar = String.fromCharCode(event.keyCode);
				var modKey  = event.ctrlKey;
				var modKey2 = event.shiftKey;

				<%
				If HasUserRight(ACCESS_TASK, RIGHT_READ) Then
					%>				
					if (modKey && modKey2 && keyChar=="P")
					{
						parent.frames[funcFrameIndex].location=("/xtra/hotlist/hotlistOppdrag.asp");
					}
					<%
				end if
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>
					if (modKey && modKey2 && keyChar=="V")
					{
						parent.frames[funcFrameIndex].location=("/xtra/hotlist/hotlistVikar.asp");
					}
					<%
				end if
				%>
			}
			
			//Her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;

		</script>
		<base target="bottom">
	</head>
	<body class="navMenuBody" onLoad="javascript:LoadMainPage('/xtra/hotlist/hotlistHoved.asp')">
		<div class="navMenu">
			<p>
				<% 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) OR HasUserRight(ACCESS_TASK, RIGHT_READ)  Then
					%>		
					<a href="hotlist/HotlistHoved.asp">Min startside</a>
					<% 
				End If 
				If HasUserRight(ACCESS_TASK, RIGHT_WRITE) Then
					%>		
					<a href="hotlist/HotlistOppdrag.asp">Hotlist-o<u>p</u>pdrag</a>
					<% 
				End If 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>		
					<a href="/xtra/hotlist/hotlistVikar.asp">Hotlist-<u>v</u>ikarer</a>
					<%
				End If 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
					%>		
					<a href="/xtra/hotlist/hotlistSkattekort.asp">Manglende skattekort</a>
					<a href="/xtra/hotlist/hotlistTimelister.asp">Mine timelister</a>
					<%
				end if
				%>
			</p>
		</div>
	</body>
</html>