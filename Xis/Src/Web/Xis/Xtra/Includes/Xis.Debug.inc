<%
sub WriteUserSessionInfo()
	Response.write Request.ServerVariables("LOGON_USER") & "<br>"
end sub
%>