<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<%
Dim FrameSrc
FrameSrc = "colorpicker.frame.asp"
If Request.ServerVariables("QUERY_STRING") <> "" Then
    FrameSrc = FrameSrc & "?" & Request.ServerVariables("QUERY_STRING")
End If
%>
<html>
	<head>
		<title><%= GetString("MoreColors") %>  </title>		
	</head>
	<body style="margin:0px;padding:0px;border-width:0px;overflow:hidden;" scroll="no">
		<iframe id='frame' frameborder='0' style='width:100%;height:100%;' src='<%=FrameSrc%>'></iframe>
	</body>
</html>