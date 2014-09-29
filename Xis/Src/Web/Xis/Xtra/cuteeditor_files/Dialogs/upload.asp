<%@  codepage="65001" %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<!-- #include file = "include_Security.asp" -->
<!-- #include file = "aspuploader/include_aspuploader.asp" -->
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
	<form action="upload_handler.asp?<%=setting %>&FP=<%=FilePath%>&Type=<%=file_Type%>&Theme=<%=Theme%>"
	 method="post" name="f" id="f">
		<%	
		Dim uploader
		Set uploader=new AspUploader
		uploader.Name="UploadControl"
		uploader.MultipleFilesUpload="true"
		
		'Dim re
		'Set re=new RegExp
		're.Pattern="\\[^\\]+\\\.\."
		'uploader.LicenseFile=re.Replace(Server.MapPath(".") & "\..\license\aspedit.lic","")
		Dim dir,pos
		dir=Server.MapPath(".")
		pos=InStrRev(dir,"\")
		dir=Left(dir,pos) ' end with the '\'
		uploader.LicenseFile=dir & "license\aspedit.lic"
		
		If file_Type = "image" Then
			uploader.AllowedFileExtensions=ImageFilters
		ElseIf file_Type = "flash" Then
			uploader.AllowedFileExtensions="*.swf,*.flv"
		ElseIf file_Type = "media" Then
			uploader.AllowedFileExtensions=MediaFilters
		ElseIf file_Type = "template" Then
			uploader.AllowedFileExtensions=TemplateFilters
		ElseIf file_Type = "document" Then
			uploader.AllowedFileExtensions=DocumentFilters
		Else
			uploader.AllowedFileExtensions="*.jpg"
		End If
		
		%>
		<%=uploader.GetString() %>
	</form>
</body>
</html>
