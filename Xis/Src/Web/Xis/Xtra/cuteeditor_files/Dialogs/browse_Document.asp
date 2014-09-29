<%@ CODEPAGE=65001 %>
<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<%
	Response.Expires = -1
	Dim folpath, goingup, current_Path, action, fso, newname, exMessage
	folpath = Request.QueryString("loc") 
	goingup = Request.QueryString("u")
	current_Path = Request.QueryString("MP")

	If Right(current_Path,1) <> "/" Then
		current_Path = current_Path & "/"
	End If
	
	If InStr(Lcase(current_Path),Lcase(trim(FilesGalleryPath))) <= 0 or FilesGalleryPath ="" then
		Response.Write "The area you are attempting to access is forbidden"	
		Response.End
	End If

	If folpath <> "" And goingup <> "y" AND Right(folpath,1) <> "/" Then
		folpath = folpath & "/"
	End If
	
	action = Request.QueryString("action")

    If DemoMode <> "true" Then
    
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        Select Case action
	        Case "deletefile"  
		        fso.DeleteFile Server.MapPath(Request.QueryString("filename")), True
                Response.Write "<script language=""javascript"">parent.ResetFields();</script>"	  
	        Case "renamefile"  
		        fso.MoveFile Server.MapPath(Request.QueryString("filename")), Server.MapPath(Request.QueryString("newname"))
		        Response.Write "<script language=""javascript"">parent.row_click("""&Request.QueryString("newname")&""");</script>"	  
	        Case "renamefolder"  
		        fso.MoveFolder Server.MapPath(Request.QueryString("filename")), Server.MapPath(Request.QueryString("newname"))
	        Case "downloadfile"  
		        filename = Request.QueryString("filename")
		        Call downloadfile(Server.MapPath(filename))
		        Response.End
	        Case "deletefolder"  
		        fso.DeleteFolder Server.MapPath(Request.QueryString("foldername")), True
	        Case "createfolder"  
		        dim folderPath 
		        folderPath = Request.QueryString("foldername")
		        folderPath = Server.MapPath(current_Path + folpath + folderPath)
		        If fso.FolderExists(folderPath) = false and fso.FileExists(folderPath) = false Then 
			        fso.CreateFolder(folderPath)
			        If err.Number <> 0 Then 
				        exMessage = "Unable to create the folder """ & folderPath & """, an error occured..." 
			        Else
				        exMessage = "Created the folder """ & folderPath & """..."
			        End If
		        Else
			        exMessage = "Unable to create the folder """ & folderPath & """, there exists a file or a folder with the same name..."
		        End If
        End Select	
	
	Else
	   CheckDemo(action)
	End If
		
	
	If(exMessage<>"") Then 
	'	Response.Write "<script language=""javascript"">alert('" & exMessage & "');</script>"
	End If

Function Showbrowse_Img(spec)

	Dim f, sf, fol, fc, fil, s, ext, counter
	Dim fso
	
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	
	Set f = fso.GetFolder(spec)
	Set sf = f.SubFolders
	s = s & "<table border=""0"" cellspacing=""1"" cellpadding=""1"" id=""FoldersAndFiles"" style=""width:100%"" class=sortable>"
	s = s & "<tr onMouseOver=""row_over(this)"" onMouseOut=""row_out(this)"" bgcolor=""#f0f0f0"">"
	s = s & "<td width=16 nowrap><img src=""../Images/refresh.gif"" title=""refresh"" onclick=""parent.Refresh('"&folpath&"');"" onMouseOver=""parent.CuteEditor_ColorPicker_ButtonOver(this);"" style=""VERTICAL-ALIGN: middle""></td>"
	s = s & "<td width=150 Class=""filelistHeadCol""><b>Name</b></td>"
	s = s & "<td width=50 nowrap Class=""filelistHeadCol""><b>Size</b></td>"
	s = s & "<td width=50 nowrap Class=""filelistHeadCol""><b>Modified</b></td>"
	s = s & "<td width=50 nowrap Class=""filelistHeadCol""><b>Created</b></td>"
	s = s & "<td width=15 nowrap Class=""filelistHeadCol""></td>"
	s = s & "<td width=80 nowrap Class=""filelistHeadCol""><b>Type</b></td>"
	If CBool(AllowDelete) Then
		s = s & "<td nowrap></td>"	
	End If
	If CBool(AllowRename) Then
		s = s & "<td nowrap></td>"	
	End If
	If CBool(AllowDelete) Then
		s = s & "<td nowrap></td>"	
	End If
	s = s & "</tr>"
	s = s & "<tr onMouseOver=""row_over(this)"" onMouseOut=""row_out(this)"" onclick=""Editor_upfolder();"">"
	s = s & "<td><img src=""../Images/parentfolder.gif"" title=""Go up one level"" style=""VERTICAL-ALIGN: middlel;cursor:hand""></td>"
	s = s & "<td>...</td>"
	s = s & "<td></td>"
	s = s & "<td></td>"
	s = s & "<td></td>"	
	s = s & "<td></td>"	
	s = s & "<td></td>"	
	If CBool(AllowDelete) Then
		s = s & "<td nowrap></td>"	
	End If
	If CBool(AllowRename) Then
		s = s & "<td nowrap></td>"	
	End If
	If CBool(AllowDelete) Then
		s = s & "<td nowrap></td>"	
	End If
	s = s & "</tr>"	
	
	For Each fol In sf 'add the html for the folders
		dim str_openfolderEvent
		
	    dim f_Name,f_Size,f_Modified,f_Created,f_Accessed,f_Type, f_Attributes,f_DisplayAttribute,f_Tooltip
	    f_Name=fol.name
	    f_Size=fol.size
	    f_Modified=fol.DateLastModified
	    f_Created=fol.DateCreated
	    f_Accessed=fol.DateLastAccessed
	    f_Type=fol.Type
	    f_Attributes=fol.Attributes
	    f_Tooltip=""
	    f_Tooltip=f_Tooltip & "<nobr>Name: "&f_Name&"</nobr><br>"
	    f_Tooltip=f_Tooltip & "<nobr>Size: "&FormatSize(f_Size)&"</nobr><br>"
	    f_Tooltip=f_Tooltip & "<nobr>Date created: "&f_Created&"</nobr><br>"
	    f_Tooltip=f_Tooltip & "<nobr>Date modified: "&f_Modified&"</nobr><br>"
	  '  f_Tooltip=f_Tooltip & "<nobr>Date accessed: "&f_Accessed&"</nobr><br>"
	    f_Tooltip=f_Tooltip & "<nobr>Read Only: "&CBool(f_Attributes And    1)&"</nobr><br>"
	    f_Tooltip=f_Tooltip & "<nobr>Hidden: "&CBool(f_Attributes And    2)&"</nobr><br>"
	  ' f_Tooltip=f_Tooltip & "<nobr>System: "&CBool(f_Attributes And    4)&"</nobr><br>"
	  ' f_Tooltip=f_Tooltip & "<nobr>Volume: "&CBool(f_Attributes And    8)&"</nobr><br>"
	  ' f_Tooltip=f_Tooltip & "<nobr>Directory: "&CBool(f_Attributes And   16)&"</nobr><br>"
	    f_Tooltip=f_Tooltip & "<nobr>Archive: "&CBool(f_Attributes And  32)&"</nobr><br>"
	  ' f_Tooltip=f_Tooltip & "<nobr>Alias: "&CBool(f_Attributes And 1024) &"</nobr><br>"
	  ' f_Tooltip=f_Tooltip & "<nobr>Compressed: "&CBool(f_Attributes And 2048)&"</nobr><br>"
	  
	    
	    if CBool(f_Attributes And  1) then
	        f_DisplayAttribute = "R" 
	    elseif CBool(f_Attributes And   2) then
	        f_DisplayAttribute = "H"	
	    elseif CBool(f_Attributes And   32) then
	        f_DisplayAttribute = "A"
	    else
	        f_DisplayAttribute = ""			        		    
	    end if 
			    
			    
		str_openfolderEvent = "onclick=""parent.SetUpload_FolderPath('"&current_Path&folpath&fol.name&"');location.href='browse_Document.asp?"&setting&"&loc=" & folpath & fol.name & "&Theme="&Theme&"&MP="&current_Path&"';"""
		s = s & "<tr onMouseOver=""row_over(this)"" onMouseOut=""row_out(this);"">"
		s = s & "<td "&str_openfolderEvent&"><img vspace=""0"" hspace=""0"" src=""../Images/closedfolder.gif"" style=""VERTICAL-ALIGN: middle""></td>" & vbcrlf
		s = s & "<td "&str_openfolderEvent&">" & f_Name & "&nbsp;&nbsp;</td>" & vbcrlf
		's = s & "<td "&str_openfolderEvent&">" & FormatSize(f_Size) & "</td>" 
		s = s & "<td "&str_openfolderEvent&"></td>" 
		s = s & "<td "&str_openfolderEvent&">" & FormatDateTime(f_Modified,2) & "</td>" 
		s = s & "<td "&str_openfolderEvent&">" & FormatDateTime(f_Created,2) & "</td>" 
		s = s & "<td "&str_openfolderEvent&" align=center>" & f_DisplayAttribute & "</td>" 
		s = s & "<td>Folder</td>" 
		if CBool(AllowDelete) Then
			s = s & "<td NOWRAP style=""cursor:pointer; border:1px"" ><img vspace=""0"" hspace=""0"" src=""../Images/delete.gif"" onclick=""deletefolder('" & current_Path & folpath  & fol.name & "')"" title=""Delete""></td>" 
		End If
		if CBool(AllowRename) Then	
			s = s & "<td NOWRAP style=""cursor:pointer; border:1px"" ><img vspace=""0"" hspace=""0"" src=""../Images/edit.gif"" title=""Rename"" onclick=""renamefolder('" & current_Path & folpath  & fol.name & "')""></td>" 
		End If		
		if CBool(AllowDelete) Then
			s = s & "<td></td>" 
		End If
		s = s & "</tr>" & vbcrlf
	Next
	Set fc = f.Files
	
	For Each fil In fc 'add the html for the files
		If (InStr(fil.name, "'" ) = 0) Then
			dim filename
			filename=fil.name
			If ValidMedia(filename) Then	
			    
			    f_Name=fil.name
			    f_Size=fil.size
			    f_Modified=fil.DateLastModified
			    f_Created=fil.DateCreated
			    f_Accessed=fil.DateLastAccessed
			    f_Type=fil.Type
			    f_Attributes=fil.Attributes
			    f_Tooltip=""
			    f_Tooltip=f_Tooltip & "<nobr>Name: "&f_Name&"</nobr><br>"
			    f_Tooltip=f_Tooltip & "<nobr>Size: "&FormatSize(f_Size)&"</nobr><br>"
			    f_Tooltip=f_Tooltip & "<nobr>Date created: "&f_Created&"</nobr><br>"
			    f_Tooltip=f_Tooltip & "<nobr>Date modified: "&f_Modified&"</nobr><br>"
			  '  f_Tooltip=f_Tooltip & "<nobr>Date accessed: "&f_Accessed&"</nobr><br>"
			    f_Tooltip=f_Tooltip & "<nobr>Read Only: "&CBool(f_Attributes And    1)&"</nobr><br>"
			    f_Tooltip=f_Tooltip & "<nobr>Hidden: "&CBool(f_Attributes And    2)&"</nobr><br>"
			  ' f_Tooltip=f_Tooltip & "<nobr>System: "&CBool(f_Attributes And    4)&"</nobr><br>"
			  ' f_Tooltip=f_Tooltip & "<nobr>Volume: "&CBool(f_Attributes And    8)&"</nobr><br>"
			  ' f_Tooltip=f_Tooltip & "<nobr>Directory: "&CBool(f_Attributes And   16)&"</nobr><br>"
			    f_Tooltip=f_Tooltip & "<nobr>Archive: "&CBool(f_Attributes And  32)&"</nobr><br>"
			  ' f_Tooltip=f_Tooltip & "<nobr>Alias: "&CBool(f_Attributes And 1024) &"</nobr><br>"
			  ' f_Tooltip=f_Tooltip & "<nobr>Compressed: "&CBool(f_Attributes And 2048)&"</nobr><br>"
			  
			    
			    if CBool(f_Attributes And  1) then
			        f_DisplayAttribute = "R" 
			    elseif CBool(f_Attributes And   2) then
			        f_DisplayAttribute = "H"	
			    elseif CBool(f_Attributes And   32) then
			        f_DisplayAttribute = "A"
			    else
			        f_DisplayAttribute = ""			        		    
			    end if 
			    
				s = s & "<tr onMouseOver=""row_over(this);"" onMouseOut=""row_out(this);"" onclick=""parent.row_click('" & current_Path & folpath & fil.name & "','" & f_Tooltip & "'); "">"
				s = s & "<td><img hspace=""3"" vspace=""1"" src=""../Images/" & GetExtension(fil.Name) & ".gif"" border=""0"" alt="""" style=""VERTICAL-ALIGN: middle""></td>" & vbcrlf
				s = s & "<td nowrap>" & f_Name & "</td>" & vbcrlf
				s = s & "<td nowrap>" & FormatSize(f_Size) & "</td>" 
				s = s & "<td nowrap>" & FormatDateTime(f_Modified,2) & "</td>" 
				s = s & "<td nowrap>" & FormatDateTime(f_Created,2) & "</td>" 
				s = s & "<td nowrap align=center>" & f_DisplayAttribute & "</td>" 
				s = s & "<td>" & f_Type & "</td>" 
				if CBool(AllowDelete) Then
					s = s & "<td><img vspace=""0"" hspace=""0"" src=""../Images/delete.gif"" title=""Delete"" onclick=""deletefile('" & current_Path & folpath  & fil.name & "')""></td>" 
				End If
				if CBool(AllowRename) Then	
					s = s & "<td><img vspace=""0"" hspace=""0"" src=""../Images/edit.gif"" title=""Rename"" onclick=""renamefile('" & current_Path & folpath  & fil.name & "','"&fil.name&"')""></td>" 
				End If
				if CBool(AllowDelete) Then
				    s = s & "<td><img vspace=""0"" hspace=""0"" src=""../Images/download.gif"" title=""Download"" onclick=""downloadfile('" & current_Path & folpath  & fil.name & "')""></td>" 
				End If
				s = s & "</tr>" & vbcrlf
			End If
		End If
	Next
	s = s & "</table>"
	Showbrowse_Img = s
   	
   	set f=nothing
	set fso=nothing
		
End Function

Function GetExtension(str_FileName)
	GetExtension = LCase(Right(str_FileName,(Len(str_FileName)-InStrRev(str_FileName,"."))))
End Function

Function ValidMedia(str_FileName)
	dim temp
	temp = LCase(Right(str_FileName,(Len(str_FileName)-InStrRev(str_FileName,"."))))
	
	dim Array_DocumentFilters
	Array_DocumentFilters	= split(DocumentFilters,",")
	dim i
	for i = 0 to ubound(Array_DocumentFilters)
		if trim(Array_DocumentFilters(i)) = "."&temp then
			ValidMedia = true
			exit for
		else
			ValidMedia = false
		end if	
	next
End Function


function FormatSize(fileSize) ' need to check later ........adam
	if Isnumeric(fileSize) then
		if fileSize < 1024 then
			FormatSize = fileSize &" B"
		elseif fileSize < 1024*1024  then
			FormatSize = FormatNumber(fileSize/1024,2) &" KB"
		else
			FormatSize = FormatNumber(fileSize/(1024*1024),2)&" MB"
		end if		
	else
		FormatSize = ""
	end if 

end function

'' this is a simple download file function. Make sure that you have MDAC 2.5+ installed in order for this to work........adam
function downloadfile(file)

	dim objFSO,objFile,objStream
	
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.GetFile(file)
    dim name, size
    
    name = objFile.Name
    size = objFile.Size
   
    Set objFile = Nothing
	Set objFSO = Nothing
	
	Response.Clear   
	Response.AddHeader "Content-Disposition", "attachment; filename=" & name
    Response.AddHeader "Content-Length", size
    Response.ContentType = "application/octet-stream"
    Set objStream = Server.CreateObject("ADODB.Stream")
	objStream.Open
	objStream.Type = 1
	Response.CharSet = "UTF-8"
	objStream.LoadFromFile(file)
	Response.BinaryWrite(objStream.Read)
	objStream.Close
	Set objStream = nothing	
end function
%>
<html>
<head>
    <title>Browse</title>
	<script type="text/javascript">
		var folderpath = "browse_Document.asp?<%=setting %>&Theme=<%=Theme%>&MP=<%=current_Path %>";
		
		function Editor_upfolder(path) {
			var arrloc = curloc.split("/"); 
			str = "";
			for (i=0;i<arrloc.length-2;i++) {
				str += arrloc[i] + "/";
			}
			
			str="browse_Document.asp?<%=setting %>&Theme=<%=Theme%>&MP=<%=current_Path %>&loc=" + str + "&u=y";
			location.href = str;
			parent.SetUpload_FolderPath('<%=current_Path %>');
		}
		
		
		function deletefile(path)
		{
			if (confirm("Delete File " + path + "?")) {
				self.location.replace(folderpath+"&loc=<%=folpath%>&action=deletefile&filename=" + path + "");
			}	
		}
		
		function deletefolder(path)
		{
			if (confirm("Delete Folder " + path + "?")) {
				self.location.replace(folderpath+"&loc=<%=folpath%>&action=deletefolder&foldername=" + path + "");
			}	
		}
		function downloadfile(path)
		{
			self.location.replace(folderpath+"&loc=<%=folpath%>&Theme=<%=Theme%>&MP=<%=current_Path%>&action=downloadfile&filename=" + path + "");
		}
		function renamefile(path,oldname)
		{
			var i=oldname.lastIndexOf('.'); 
			var ext=oldname.substr(i);
			var t= oldname.substr(0,oldname.lastIndexOf('.'));
			
			if(parent.Browser_IsIE7())
	        {
		        parent.IEprompt(promptCallback,'Type the new name for this file:', "");		
		        function promptCallback(newName)
		        {					 
	               if ((newName) && (newName!=""))
	               {
						newName=newName + ext;    	
		                self.location.replace(folderpath+"&loc=<%=folpath%>&action=renamefile&filename=" + path + "&newname=<%=current_Path%><%=folpath%>" + newName + "");
	               }	
		        }
	        }
	        else
	        {
			    var newName=prompt("Type the new name for this file:",t);
    						 
			    if ((newName) && (newName!=""))
			    {
				    newName=newName + ext;
				    self.location.replace(folderpath+"&loc=<%=folpath%>&action=renamefile&filename=" + path + "&newname=<%=current_Path%><%=folpath%>" + newName + "");
			    }
	        }	
		}
		
		
		function renamefolder(path)
		{
		    if(parent.Browser_IsIE7())
	        {
		        parent.IEprompt(promptCallback,'Type the new name for this folder:', "");		
		        function promptCallback(newName)
		        {
			        if ((newName) && (newName!=""))
			        {
				        self.location.replace(folderpath+"&loc=<%=folpath%>&action=renamefolder&filename=" + path + "&newname=<%=current_Path%><%=folpath%>" + newName + "");
			        }	
		        }
	        }
	        else
	        {
			    var newName = prompt('Type the new name for this folder:','');
			    if ((newName) && (newName!=""))
			    {
				    self.location.replace(folderpath+"&loc=<%=folpath%>&action=renamefolder&filename=" + path + "&newname=<%=current_Path%><%=folpath%>" + newName + "");
			    }	
	        }	
		}
		function row_over(row)
		{
			row.style.backgroundColor='#eeeeee';
		}
		function row_out(row)
		{
			row.style.backgroundColor='';
		}
		
	</script>
	<script type="text/javascript">var curloc = "<%=folpath%>";</script> 
	<script type="text/javascript" src="filebrowserpage.js"></script>
	<script type="text/javascript" src="sorttable.js"></script>
    <link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
</head>
<body style="overflow:auto;background-color:white">
    <% 
		Response.Write Showbrowse_Img(Server.MapPath(current_Path & folpath)) 
	%>
</body>
</html>