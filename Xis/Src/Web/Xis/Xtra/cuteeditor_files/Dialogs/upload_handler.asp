<%@  codepage="65001" %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<!-- #include file = "include_Security.asp" -->
<!-- #include file = "aspuploader/include_aspuploader.asp" -->
<%
Function ValidFileName(str_FileName)
	Set RegularExpressionObject = New RegExp

	With RegularExpressionObject
		.Pattern = "^[a-zA-Z0-9\._-]+$"
		.IgnoreCase = True
		.Global = True
	End With
	
	ValidFileName = RegularExpressionObject.Test(str_FileName)
End Function

Function ValidFileExtension(str_FileName)
	
	dim temp
	temp = LCase(Right(str_FileName,(Len(str_FileName)-InStrRev(str_FileName,"."))))
	
	dim Array_DocumentFilters
	Array_DocumentFilters	= split(DocumentFilters,",")
	dim i
	for i = 0 to ubound(Array_DocumentFilters)
		if lcase(trim(Array_DocumentFilters(i))) = "."&temp then
			ValidFileExtension = true
			exit for
		else
			ValidFileExtension = false
		end if	
	next
End Function

Function ValidFlash(str_FileName)
	dim temp, TemplateFilters
	temp = LCase(Right(str_FileName,(Len(str_FileName)-InStrRev(str_FileName,"."))))
	FlashFilters = ".swf,.flv"
	dim Array_FlashFilters
	Array_FlashFilters	= split(FlashFilters,",")
	dim i
	for i = 0 to ubound(Array_FlashFilters)
		if trim(Array_FlashFilters(i)) = "."&temp then
			ValidFlash = true
			exit for
		else
			ValidFlash = false
		end if	
	next
End Function

Function ValidTemplate(str_FileName)

	dim temp, TemplateFilters
	temp = LCase(Right(str_FileName,(Len(str_FileName)-InStrRev(str_FileName,"."))))
	TemplateFilters = ".html,.htm,"
	dim Array_TemplateFilters
	Array_TemplateFilters	= split(TemplateFilters,",")
	dim i
	for i = 0 to ubound(Array_TemplateFilters)
		if trim(Array_TemplateFilters(i)) = "."&temp then
			ValidTemplate = true
			exit for
		else
			ValidTemplate = false
		end if	
	next
End Function

Function ValidMedia(str_FileName)

	dim temp
	temp = LCase(Right(str_FileName,(Len(str_FileName)-InStrRev(str_FileName,"."))))
	
	dim Array_MediaFilters
	Array_MediaFilters	= split(MediaFilters,",")
	dim i
	for i = 0 to ubound(Array_MediaFilters)
		if lcase(trim(Array_MediaFilters(i))) = "."&temp then
			ValidMedia = true
			exit for
		else
			ValidMedia = false
		end if	
	next
End Function


Function ValidImage(str_FileName)

	dim temp
	temp = LCase(Right(str_FileName,(Len(str_FileName)-InStrRev(str_FileName,"."))))
	
	dim Array_ImageFilters
	Array_ImageFilters	= split(ImageFilters,",")
	dim i
	for i = 0 to ubound(Array_ImageFilters)
		if lcase(trim(Array_ImageFilters(i))) = "."&temp then
			ValidImage = true
			exit for
		else
			ValidImage = false
		end if	
	next
End Function

Dim FilePath,MaxSize, file_Type
FilePath = Request.QueryString("FP")
file_Type = Lcase(Trim(Request.QueryString("Type")))			

Select Case lcase(file_Type)
	case "image":
		MaxSize =  MaxImageSize	
	case "flash":
		MaxSize =  MaxFlashSize		
	case "media":
		MaxSize =  MaxMediaSize	
	case "template":
		MaxSize =  MaxTemplateSize			
	Case Else			
		MaxSize =  MaxDocumentSize			
End Select

Dim guidlist,i
guidlist=Request.Form("UploadControl") & ""

' Check if any files were uploaded
If guidlist="" Then
	Response.Write "File(s) not uploaded."
Else
	
	Dim Uploader, File
	Set Uploader = New AspUploader
	
	guidlist=Split(guidlist,"/")
	
	For i=0 to Ubound(guidlist)
		Set File=uploader.GetUploadedFile(guidlist(i))

		If ValidFileName(File.FileName) = false then
			Response.Write "<font color=red><b>File name not allowed!</b></font><br><br>Please keep the file name one word with no spaces or special characters."
		else
			If File.FileSize >= MaxSize*1024 then	
				Response.Write "<font color=red><b><font color=red><b>File size exceeds "& MaxSize &" KB limit: "& Formatnumber(File.FileSize/1024,2) &" KB</b></font>"
			ElseIf file_Type = "image" and not ValidImage(File.FileName)Then
				Response.Write "<font color=red><b>File format not allowed!</b></font>"									
			ElseIf file_Type = "flash" and not ValidFlash(File.FileName) Then
				Response.Write "<font color=red><b>File format not allowed!</b></font>"
			ElseIf file_Type = "media" and not ValidMedia(File.FileName)Then
				Response.Write "<font color=red><b>File format not allowed!</b></font>"							
			ElseIf file_Type = "template" and not ValidTemplate(File.FileName)Then
				Response.Write "<font color=red><b>File format not allowed!</b></font>"						
			ElseIf file_Type = "document" and not ValidFileExtension(File.FileName)Then
				Response.Write "<font color=red><b>File format not allowed!</b></font>"									
			ElseIf file_Type = "flash" or  file_Type = "image" or file_Type = "document" or file_Type = "media" or file_Type = "template" then
				
				Dim fso,newpath
				Set fso=CreateObject("Scripting.FileSystemObject")
				newpath=Server.MapPath(FilePath)
				newpath=fso.BuildPath(newpath,File.FileName)
				If fso.FileExists(newpath) Then
					fso.DeleteFile newpath
				End If
				
				File.MoveTo newpath
				
				' Output the file details to the browser
				Response.Write File.FileName&" uploaded successfully!<br>"
				Response.Write "Size: " & Formatnumber(File.FileSize/1024,2) & " KB<br>"
				' Response.Write "Type: " & File.ContentType  Server.MapPath(FilePath &"/" & File.FileName) &"<br>"	
				Response.Write "<script language=javascript>parent.UploadSaved('" & FilePath &"/" & File.FileName & "','"&FilePath&"');</script>"
			End If
		end if 
	Next
End If

%>
<html>
<head>
    <title>Upload</title>
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
</head>
<body>
	<div style="vertical-align:top">
		<a href="upload.asp?<%=setting %>&FP=<%=FilePath%>&MaxSize=<%=MaxSize%>&Theme=<%=Theme%>&Type=<%=file_Type%>&BG=<%= Server.URLEncode(Request.QueryString("BG")) %>">Upload again</a>
	</div>
</body>
</html>