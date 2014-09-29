<html>
	<head>
		<title>Posted Values</title>
<%
	Dim path
    path=Request.ServerVariables("SCRIPT_NAME")
    path=left(path,len(path)-12)    
%>
		 <link href="<%=path%>Style/SyntaxHighlighter.css" rel="stylesheet" type="text/css">
	</head>
	<body>
	</body>
</html>