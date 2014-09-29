<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<% 
If  HasUserRight(ACCESS_ADMIN, RIGHT_ADMIN) Then 
	%>
	<html>
		<head>
			<title></title>
		</head>
		<frameset cols="250,*" framespacing="0" frameborder="0">
			<frame src="BrukerVis.asp" frameborder="0" scrolling="Auto" marginwidth="0" marginheight="0">
			<frame src="tomt.htm" name="RIGHT_WINDOW" id="RIGHT_WINDOW" frameborder="0" scrolling="auto" noresize marginwidth="0" marginheight="0">
		</frameset>
	</html>
	<%
End If 
%>