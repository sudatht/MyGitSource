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
	
	If InStr(Lcase(current_Path),Lcase(trim(FlashGalleryPath))) <= 0 or FlashGalleryPath ="" then
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
	s = s & "<table border=""0"" cellspacing=""1"" cellpadding=""1"" valign=""top"" id=""FoldersAndFiles"" style=""width:100%"" class=sortable>"
	s = s & "<tr onMouseOver=""row_over(this)"" onMouseOut=""row_out(this)"" bgcolor=""#f0f0f0"">"
	s = s & "<td width=16 nowrap><img src=""../Images/refresh.gif"" title=""refresh"" onclick=""parent.Refresh('"&folpath&"');"" onMouseOver=""parent.CuteEditor_ColorPicker_ButtonOver(this);"" style=""VERTICAL-ALIGN: middle;border:#f0f0f0 1px solid;""></td>"
	s = s & "<td width=""60%"" Class=""filelistHeadCol""><b>Name</b></td>"
	s = s & "<td nowrap Class=""filelistHeadCol""><b>Size</b></td>"
	If CBool(AllowDelete) Then
		s = s & "<td nowrap></td>"	
	End If
	If CBool(AllowRename) Then
		s = s & "<td nowrap></td>"	
	End If
	s = s & "</tr>"
	s = s & "<tr onMouseOver=""row_over(this)"" onMouseOut=""row_out(this)"" onclick=""Editor_upfolder();"">"
	s = s & "<td><img src=""../Images/parentfolder.gif"" title=""Go up one level"" style=""VERTICAL-ALIGN: middlel;cursor:hand""></td>"
	s = s & "<td>...</td>"
	s = s & "<td></td>"
	If CBool(AllowDelete) Then
		s = s & "<td nowrap></td>"	
	End If
	If CBool(AllowRename) Then
		s = s & "<td nowrap></td>"	
	End If
	s = s & "</tr>"	
	
	For Each fol In sf 'add the html for the folders
		dim str_openfolderEvent
		str_openfolderEvent = "onclick=""parent.SetUpload_FolderPath('"&current_Path&folpath&fol.name&"');location.href='browse_flash.asp?"&setting&"&loc=" & folpath & fol.name & "&Theme="&Theme&"&MP="&current_Path&"';"""
		s = s & "<tr onMouseOver=""row_over(this)"" onMouseOut=""row_out(this)"">"
		s = s & "<td "&str_openfolderEvent&"><img vspace=""0"" hspace=""0"" src=""../Images/closedfolder.gif"" style=""VERTICAL-ALIGN: middle""></td>" & vbcrlf
		s = s & "<td valign=""top"" style=""cursor:pointer;VERTICAL-ALIGN: middle"" "&str_openfolderEvent&">" & vbcrlf
		s = s & fol.name & "&nbsp;&nbsp;</td>" & vbcrlf
		's = s & "<td NOWRAP style=""cursor:pointer;"" "&str_openfolderEvent&">" & FormatSize(fol.size) & "</td>" 
		s = s & "<td "&str_openfolderEvent&"></td>" 
		if CBool(AllowDelete) Then
			s = s & "<td NOWRAP style=""cursor:pointer; border:1px;VERTICAL-ALIGN: middle"" ><img vspace=""0"" hspace=""0"" src=""../Images/delete.gif"" onclick=""deletefolder('" & current_Path & folpath  & fol.name & "')"" title=""Delete""></td>" 
		End If
		if CBool(AllowRename) Then	
			s = s & "<td NOWRAP style=""cursor:pointer; border:1px;VERTICAL-ALIGN: middle"" ><img vspace=""0"" hspace=""0"" src=""../Images/edit.gif"" title=""Rename"" onclick=""renamefolder('" & current_Path & folpath  & fol.name & "')""></td>" 
		End If		
		s = s & "</tr>" & vbcrlf
	Next
	Set fc = f.Files
	
	For Each fil In fc 'add the html for the files
		If (InStr(fil.name, "'" ) = 0) Then
			dim filename
			filename=fil.name
			If ValidImage(filename) Then
				
				s = s & "<tr onMouseOver=""row_over(this)"" onMouseOut=""row_out(this)"" onclick=""parent.row_click('" & current_Path & folpath & fil.name & "'); "">"
				s = s & "<td style=""cursor:pointer"" ><img hspace=""3"" vspace=""1"" src=""../Images/" & GetExtension(fil.Name) & ".gif"" border=""0"" alt="""" style=""VERTICAL-ALIGN: middle""></td>" & vbcrlf
				s = s & "<td align=""left"" valign=""top"" style=""cursor:pointer;VERTICAL-ALIGN: middle"" >" & vbcrlf
				s = s & fil.name & "</td>" & vbcrlf
				s = s & "<td NOWRAP>" & FormatSize(fil.size) & "</td>" 
				if CBool(AllowDelete) Then
					s = s & "<td><img vspace=""0"" hspace=""0"" src=""../Images/delete.gif"" title=""Delete"" onclick=""deletefile('" & current_Path & folpath  & fil.name & "')""></td>" 
				'	s = s & "<td><img vspace=""0"" hspace=""0"" src=""../Images/download.gif"" title=""Download"" onclick=""downloadfile('" & current_Path & folpath  & fil.name & "')""></td>" 
				End If
				if CBool(AllowRename) Then	
					s = s & "<td><img vspace=""0"" hspace=""0"" src=""../Images/edit.gif"" title=""Rename"" onclick=""renamefile('" & current_Path & folpath  & fil.name & "','"&fil.name&"')""></td>" 
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

Function ValidImage(str_FileName)
	ValidImage = InStr(lcase(str_FileName), ".swf" ) <> 0 or InStr(lcase(str_FileName), ".flv" ) <> 0 
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

%>
<html>
<head>
    <title>Browse</title>
	<script type="text/javascript">
		var folderpath = "browse_flash.asp?<%=setting %>&MP=<%=current_Path %>&Theme=<%=Theme%>";
				
		function Editor_upfolder() {
			arrloc = curloc.split("/"); 
			str = "";
			for (i=0;i<arrloc.length-2;i++) {
				str += arrloc[i] + "/";
			}
			location.href = folderpath+"&loc=" + str + "&u=y";
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
			    if(newName)
				    newName=newName + ext;
    						 
			    if ((newName) && (newName!=""))
			    {
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
    <link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
	<script type="text/javascript">var curloc = "<%=folpath%>";</script> 
	<script type="text/javascript" src="filebrowserpage.js"></script>
	<script type="text/javascript" src="sorttable.js"></script>
</head>
<body style="overflow:auto;background-color:white">
    <% 
		Response.Write Showbrowse_Img(Server.MapPath(current_Path & folpath)) 
	%>  
</body>
</html>