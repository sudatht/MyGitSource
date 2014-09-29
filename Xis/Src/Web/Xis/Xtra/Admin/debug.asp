<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Response.Expires = 0
%>
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<%
	dim debugLabel
	dim debugState
	dim sv
	dim debugStateChanged : debugStateChanged = false

	If (HasUserRight(ACCESS_ADMIN, RIGHT_ADMIN) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	
	if (request("IsPostback") = "true") then
		if (LenB(request("debug")) <> 0) then
			debugState = request("debug")
			debugStateChanged = true
			if(debugState = "debug på") then
				session("debug") = true
			elseif(debugState = "debug av") then
				session("debug") = false
			end if
		end if
	end if
	if(session("debug") = true) then
		debugLabel = "debug av"
	else
		debugLabel = "debug på"
	end if	
%>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en" "http://www.w3.org/tr/rec-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<title>Debug</title>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Debug</h1>
			</div>			
			<div class="content">
				<p>
				<form action="Debug.asp"  method="post" id="frmPost" >
					<p>
					<input type="hidden" id="IsPostback" name="IsPostback" value="true">
					<input type="submit" id="btnDisplayServerVars" name="btnDisplayServerVars" value="Server variable">
					<input type="submit" id="btnDisplaySessionVars" name="btnDisplaySessionVars" value="Sesjons variable">
					<input type="submit" id="btnDisplayApplicationVars" name="btnDisplayApplicationVars" value="Applikasjons variable">
					<input type="submit" id="debug" name="debug" value="<%=debugLabel%>">					
					</p>
					<%
					'Session vars:
					if (LenB(Request("btnDisplayServerVars") ) <> 0) then
						%>
						<h1>Server variable</h1>
						<div class="listing">
							<table>
								<%
								for each sv in Request.ServerVariables
									Response.Write "<tr><th>" & sv & "</th><td>"  & Request.ServerVariables.Item(sv) & "</td></tr>"
								next
								%>
							</table>
						</div>
						<%
					elseif (LenB(Request("btnDisplaySessionVars") ) <> 0) then
						%>
						<h1>Session variable</h1>
						<div class="listing">
							<table>
								<%
								for each sv in Session.Contents
									Response.Write "<tr><th>" & sv & "</th><td>"  &  Session.Contents.Item(sv) & "</td></tr>"
								next
								%>
							</table>
						</div>
						<%
					elseif (LenB(Request("btnDisplayApplicationVars") ) <> 0) then
						%>
						<h1>Application variable</h1>
						<div class="listing">
							<table id="Table1">
								<%
								for each sv in Application.Contents
									Response.Write "<tr><th>" & sv & "</th><td>"  &  Application.Contents.Item(sv) & "</td></tr>"
								next
								%>
							</table>
						</div>
						<%					
					elseif (debugStateChanged = true) then
						%>
						<h1>Debug er endret </h1>
						<div class="listing">
							<table>
								<tr>
									<th>Debugstatus:</th>
									<td><% if (session("debug") = true ) then Response.Write "På" else Response.Write "Av" end if%></td>
								</tr>
							</table>
						</div>
						<%					
					end if
					%>					
				</form>
			</div>
		</div>
	</body>
</html>