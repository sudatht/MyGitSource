<!-- #include file = "Include_GetString.asp" -->

<%
    dim ua
    ua = lcase(Request.ServerVariables("HTTP_USER_AGENT"))
    
    if instr(ua,"msie") > 0 then    
		Server.Transfer("colorpicker_more_ie.asp")
	else
		Server.Transfer("colorpicker_more_ns.asp")	
    end if
%>
