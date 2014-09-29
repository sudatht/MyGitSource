<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<html>
	<head>		
		<title>Upload Page</title>		
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
	</head>
	<body>
	<%
		Dim FilePath,file_Type
		file_Type = Lcase(Trim(Request.QueryString("Type")))		
		If file_Type <> "" then
			FilePath = Request.QueryString("FP")		
		End If
	%>	
	<form action="filePost.asp?<%=setting %>&FP=<%=FilePath%>&Type=<%=file_Type%>&Theme=<%=Theme%>" enctype="multipart/form-data" method="post" name="f" id="f">
			<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td>
						<input type="file" name="test" size="35"/>
					</td>
					<td>&nbsp;&nbsp;</td>
					<td id="uploader">
						<input type="submit" value="upload" id="Submit1" name="Submit1"/>
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
