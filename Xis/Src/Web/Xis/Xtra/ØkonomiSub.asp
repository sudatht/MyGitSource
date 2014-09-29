<%@ Language=VBScript%>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes/xis.rights.inc"-->
<!--#INCLUDE FILE="includes/Library.inc"-->
<%
	If (HasUserRight(ACCESS_ADMIN, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/admin/tomt.htm")
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<title>adminSubMeny</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script language='javascript' src='js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='js/navigation.js' id='navigationScripts'></script>
		<base target="bottom">
	</head>
	<body <% If HasUserRight(ACCESS_ADMIN, RIGHT_READ) Then %> onLoad="javascript:LoadMainPage('datafiler/TimelisteMeny.asp')" <%end if%> class="navMenuBody">
		<div class="navMenu">
			<p>
			<% 
			If HasUserRight(ACCESS_ADMIN, RIGHT_READ) Then 
				%>
				<a href="datafiler/TimelisteMeny.asp?viskode=1">Timelister</a>
				<a href="datafiler/LoennMeny.asp">Lønn</a>
				<a href="datafiler/FakturaMeny.asp">Faktura</a>
				<% 
			End If
			If HasUserRight(ACCESS_ADMIN, RIGHT_ADMIN) Then 
				%>
				<a href="datafiler/AdministrasjonsMeny.asp">Administrasjon</a>
				<% 
			End If 
			%>
			</p>
		</div>
	</body>
</html>