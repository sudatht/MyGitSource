<%@ CODEPAGE=65001 %>
<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<%
	Response.Expires = -1	
	Dim folpath, goingup, c_ImageGalleryPath, action, fso, newname, bkP,startimage,cp
	bkP = 15 
	folpath = Request.QueryString("loc")
	goingup = Request.QueryString("u")
	c_ImageGalleryPath = Request.QueryString("GP")

	If Right(c_ImageGalleryPath,1) <> "/" Then
		c_ImageGalleryPath = c_ImageGalleryPath & "/"
	End If
	
	If InStr(Lcase(c_ImageGalleryPath),Lcase(trim(ImageGalleryPath))) <= 0 or ImageGalleryPath ="" then
		Response.Write ImageGalleryPath & "The area you are attempting to access is forbidden"	
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
		    Case "renamefile"  
			    fso.MoveFile Server.MapPath(Request.QueryString("filename")), Server.MapPath(Request.QueryString("newname"))
		    Case "renamefolder"  
			    fso.MoveFolder Server.MapPath(Request.QueryString("filename")), Server.MapPath(Request.QueryString("newname"))
		    Case "deletefolder"  
			    fso.DeleteFolder Server.MapPath(Request.QueryString("foldername")), True
		End Select	
	
	Else
	   CheckDemo(action)
	End If

	Function Showbrowse_Img(spec)
		Dim f, sf, fol, fc, fil, s, ext, counter
		Dim fso
	
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
	
		Set f = fso.GetFolder(spec)
		Set sf = f.SubFolders
		s = s & "<div style='overflow: auto; HEIGHT: 120px;'>"
		s = s & "<table border=""0"" cellspacing=""0"" cellpadding=""1"" width=""100%"" align=""center"" valign=""top"">"
		For Each fol In sf 'add the html for the folders
			dim p
			p=c_ImageGalleryPath&folpath&fol.name
			s = s & "<tr><td NOWRAP valign=""top"" style=""cursor:pointer"" onclick=""SetUpload_imagePath('"&p&"');location.href='browse_Img_gallery.asp?"&setting&"&loc=" & folpath & fol.name & "&Theme="&Theme&"&GP="&c_ImageGalleryPath&"';"">" & vbcrlf
			s = s & "<img vspace=""0"" hspace=""0"" src=""../Images/closedfolder.gif"" style=""VERTICAL-ALIGN: middle"">&nbsp;" & fol.name & "&nbsp;&nbsp;</td>"
			s = s & "</tr>" & vbcrlf	
		Next
		
		s = s & "<tr><td valign=""top"">" & vbcrlf
		
		Set fc = f.Files
		
		dim bkFileCount
		bkFileCount =fc.Count
		
		' Data source specific vars
		Dim rstTableData
		Const adVarChar = 200
		Const adInteger = 3
		Const adDate = 7

		' Table Building Vars
		Dim int_Columns, int_Rows
		Dim int_WidthLooper, int_HeightLooper

		Dim int_PageCount
		Dim int_CurrentPage

		Dim blnOutOfData
		Dim I

		' Read in parameters + set defaults if none entered.
		' You could save chosen values to a DB or cookies or whatever, but it
		' all sort of depends on what you're using the script for.
		int_Columns  = CInt(Request.QueryString("width"))
		If int_Columns <= 0 Then int_Columns = 5

		int_Rows = CInt(Request.QueryString("height"))
		If int_Rows <= 0 Then int_Rows = 4

		int_CurrentPage = CInt(Request.QueryString("page"))


		' Make up some sample data to display.  I'm using an in-memory recordset
		' simply to keep things fast and because I assume most users will be using
		' this script with some sort of database.  This way you should just be
		' able to drop in your own DB code and edit the display section and you
		' should have everything working pretty quickly.
		Set rstTableData = Server.CreateObject("ADODB.Recordset")
		rstTableData.Fields.Append "name", adVarChar, 150
		rstTableData.Fields.Append "fullpath", adVarChar, 100
		rstTableData.Fields.Append "size", adInteger
		rstTableData.Fields.Append "DateCreated", adVarChar, 50
		rstTableData.Fields.Append "DateLastModified", adVarChar, 50

		rstTableData.Open
		
		For Each fil In fc 'add the html for the files			
			If (InStr(fil.name, "'" ) = 0) Then
				dim filename
				filename = lcase(fil.name)
				dim fullpath
				fullpath=c_ImageGalleryPath & folpath & fil.name
				'If (InStr(filename, ".gif" ) <> 0) Or (InStr(filename, ".jpg" ) <> 0) Or (InStr(filename, ".png" ) <> 0 )  Or (InStr(filename, ".bmp" ) <> 0 )Then
				If ValidImage(filename) Then 
					rstTableData.AddNew
					rstTableData.Fields("name").Value       = fil.name
					rstTableData.Fields("fullpath").Value   = fullpath
					rstTableData.Fields("size").Value       = fil.size
					rstTableData.Fields("DateCreated").Value = fil.DateCreated
					rstTableData.Fields("DateLastModified").Value       = fil.DateLastModified
					rstTableData.Update
				End If
			End If
		Next 'I


		' Display table using parameters for table size and page # and using data from recordset.

		' Calculate # of pages - Divide items by page size and find the next largest whole number.
		int_PageCount = -(Int(-(rstTableData.RecordCount / (int_Columns * int_Rows))))
		' Sorry for the above line... it doesn't make much logical sense, but it was the easiest
		' one line implementation I could come up with.  There must be a more logical one...
		' and no it's not Int(...) + 1

		' If page size is larger then # of records, reset page count to 1
		If int_Columns * int_Rows >= rstTableData.RecordCount Then int_PageCount = 1

		' If current page falls outside acceptable range, default to page 1
		If 0 >= int_CurrentPage Or int_CurrentPage > int_PageCount Then int_CurrentPage = 1

		' Move int_o recordset the appropriate number of pages
		If Not rstTableData.EOF Then
			rstTableData.MoveFirst
			rstTableData.Move (int_CurrentPage - 1) * (int_Columns * int_Rows)		
		End If
		
		s = s & "</td></tr></table></div>"
		
		' Show our table
		s = s & "<div style='HEIGHT: 280px;'><table width='100%' height='250' CellSpacing='0' valign='top'>"
		For int_HeightLooper = 1 To int_Rows
			If Not rstTableData.EOF Then
				s = s & "<tr>"
				For int_WidthLooper = 1 to int_Columns
					If Not rstTableData.EOF Then
						s = s & vbTab & "<td valign=top>"
						's = s & rstTableData.Fields("name").Value
						s = s & "<img onclick=""parent.insert(this.src)"" src=""" & rstTableData.Fields("fullpath").Value & """ width=""80"" height=""56"" onMouseover=""Check(this,1); showTooltip('<nobr>Name: "&rstTableData.Fields("name").Value&"</nobr><br><nobr>Size: "&FormatSize(rstTableData.Fields("size").Value)&"</nobr><br><nobr>Date created: "&rstTableData.Fields("DateCreated").Value&"</nobr><br><nobr>Date modified: "&rstTableData.Fields("DateLastModified").Value&"</nobr>', this, event)"" onMouseout='Check(this,0); delayhidetip()' style='BORDER: white 1px solid' align='center'>"
						s = s & "</td>"
						rstTableData.MoveNext
					Else
						' Out of data... bail out of cells early!
						blnOutOfData = True
						Exit For
					End If
				Next

				' Clean up odd table cells if any and bail out of rows!
				If blnOutOfData Then
					If rstTableData.RecordCount Mod int_Columns > 0 And int_HeightLooper > 1 Then
						For I = (rstTableData.RecordCount Mod int_Columns) + 1 To int_Columns
							s = s & "<td>&nbsp;</td>"
						Next ' I
					End If
					s = s & "</tr>"
					Exit For
				End If

				s = s & "</tr>"
			End If
		Next ' int_HeightLooper
		s = s & "</table></div>"
		s = s & "<center>"
		' Show paging indicator if multiple pages:
		If int_Columns * int_Rows < rstTableData.RecordCount Then
			s = s & "Pages:&nbsp;"
			
			If int_CurrentPage > 1 Then
				s = s & "<a href=""?"&setting&"&Theme="&Theme&"&GP="&c_ImageGalleryPath & folpath&"&width=" & int_Columns & "&height=" & int_Rows & "&page=" & int_CurrentPage - 1 & """>&lt;&lt; Prev</a>&nbsp;"
			End If

			' You can also show page numbers:
			For I = 1 To int_PageCount
				If I = int_CurrentPage Then
					s = s & I & "&nbsp;"
				Else
					s = s & "<a href=""?"&setting&"&Theme="&Theme&"&GP="&c_ImageGalleryPath & folpath&"&width=" & int_Columns & "&height=" & int_Rows & "&page=" & I & """>" & I & "</a>&nbsp;"
				End If
			Next 'I

			If int_CurrentPage < int_PageCount Then
				s = s & "<a href=""?"&setting&"&Theme="&Theme&"&GP="&c_ImageGalleryPath & folpath&"&width=" & int_Columns & "&height=" & int_Rows & "&page=" & int_CurrentPage + 1 & """>Next &gt;&gt;</a>"
			End If
		End If
		
		s = s & "</center>"

		' Close down RS
		rstTableData.Close
		Set rstTableData = Nothing
		
		
		Showbrowse_Img = s
		
		'Set the start image
		If startimage = "" then
			startimage = 1
		else
			startimage = request("startimage")
		end if
		'If there is no page number set, then set it to 1, otherwise get the page number
		If Request("cp") = "" then
			cp = "1"
		else
			cp = Request("cp")
		end if
	
	   	set f=nothing
		set fso=nothing
	End Function
		
	Function GetExtension(str_FileName)
		GetExtension = LCase(Right(str_FileName,(Len(str_FileName)-InStrRev(str_FileName,"."))))
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
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Browse</title>
<script type="text/javascript">

	function highlightcell(c) {
		var allcells = document.getElementsByTagName("TD");
		for (i=0;i<allcells.length;i++) {
			allcells[i].style.backgroundColor = "white"; allcells[i].style.color = "black";
		}
		c.style.backgroundColor = "#0a246a"; c.style.color = "white";
	}
	
	var folderpath = "browse_Img_gallery.asp?<%=setting%>&Theme=<%=Theme%>&GP=<%=c_ImageGalleryPath %>";
	
	function Editor_upfolder() {
		arrloc = curloc.split("/"); 
		str = "";
		for (i=0;i<arrloc.length-2;i++) {
			str += arrloc[i] + "/";
		}
		parent.currentfolder = folderpath+"&loc=" + str + "&u=y";
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
		newName = prompt('Type the new name for this file:',oldname);
		if ((newName) && (newName!=""))
		{
			self.location.replace(folderpath+"&loc=<%=folpath%>&action=renamefile&filename=" + path + "&newname=<%=c_ImageGalleryPath%><%=folpath%>" + newName + "");
		}	
	}
	
	function renamefolder(path)
	{
		newName = prompt('Type the new name for this folder:','');
		if ((newName) && (newName!=""))
		{
			self.location.replace(folderpath+"&loc=<%=folpath%>&action=renamefolder&filename=" + path + "&newname=<%=c_ImageGalleryPath%><%=folpath%>" + newName + "");
		}	
	}
	
			
	function Check(t,n)	
	{
		if(n==1) {
			t.style.border = "1px solid #00107B";
			t.style.background = "#F1EEE7";
		}
		else  {
			t.style.border = "1px solid #d7d3cc";
			t.style.background = "#d7d3cc";
		}
	}
	
	function UploadSaved(sFileName,path){
		document.getElementById("TargetUrl").value = sFileName;
	}
	
	
</script>
    <script language="JavaScript">var curloc = "<%=folpath%>";</script> 
	<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
	

	<style type="text/css">
		INPUT { BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; FONT-SIZE: 8pt; VERTICAL-ALIGN: middle; BORDER-LEFT: black 1px solid; CURSOR: pointer; BORDER-BOTTOM: black 1px solid; FONT-FAMILY: MS Sans Serif }
		A:link { COLOR: #000099 }
		A:visited { COLOR: #000099 }
		A:active { COLOR: #000099 }
		A:hover { COLOR: darkred }
		#tooltipdiv{
		position:absolute;
		padding: 2px;
		border:1px solid black;
		font:menu;
		z-index:100;
		}
		body
		{
			background-color:#eeeeee;	
			overflow:hidden;margin:0px; border:0px;
		}
		</style>
	<script type="text/javascript">

				
				var tipbgcolor='lightyellow'  //tooltip bgcolor
				var disappeardelay=250  //tooltip disappear speed onMouseout (in miliseconds)
				var vertical_offset="0px" //horizontal offset of tooltip from anchor link
				var horizontal_offset="-3px" //horizontal offset of tooltip from anchor link

				/////No further editting needed

				var ie4=document.all
				var ns6=document.getElementById&&!document.all

				if (ie4||ns6)
				document.write('<div id="tooltipdiv" style="visibility:hidden;background-color:'+tipbgcolor+'" ></div>')

				function getposOffset(what, offsettype){
				var totaloffset=(offsettype=="left")? what.offsetLeft : what.offsetTop;
				var parentEl=what.offsetParent;
				while (parentEl!=null){
				totaloffset=(offsettype=="left")? totaloffset+parentEl.offsetLeft : totaloffset+parentEl.offsetTop;
				parentEl=parentEl.offsetParent;
				}
				return totaloffset;
				}


				function showhide(obj, e, visible, hidden){
				if (ie4||ns6)
				dropmenuobj.style.left=dropmenuobj.style.top=-500;
				if (e.type=="click" && obj.visibility==hidden || e.type=="mouseover")
				obj.visibility=visible
				else if (e.type=="click")
				obj.visibility=hidden
				}

				function iecompattest(){
				return (document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body
				}

				function clearbrowseredge(obj, whichedge){
				var edgeoffset=(whichedge=="rightedge")? parseInt(horizontal_offset)*-1 : parseInt(vertical_offset)*-1
				if (whichedge=="rightedge"){
				var windowedge=ie4 && !window.opera? iecompattest().scrollLeft+iecompattest().clientWidth-15 : window.pageXOffset+window.innerWidth-15
				dropmenuobj.contentmeasure=dropmenuobj.offsetWidth
				if (windowedge-dropmenuobj.x < dropmenuobj.contentmeasure)
				edgeoffset=dropmenuobj.contentmeasure-obj.offsetWidth
				}
				else{
				var windowedge=ie4 && !window.opera? iecompattest().scrollTop+iecompattest().clientHeight-15 : window.pageYOffset+window.innerHeight-18
				dropmenuobj.contentmeasure=dropmenuobj.offsetHeight
				if (windowedge-dropmenuobj.y < dropmenuobj.contentmeasure)
				edgeoffset=dropmenuobj.contentmeasure+obj.offsetHeight
				}
				return edgeoffset
				}

				function showTooltip(menucontents, obj, e){
				if (window.event) 
					event.cancelBubble=true
				else 
					if (e.stopPropagation) e.stopPropagation()
						clearhidetip()
				dropmenuobj=document.getElementById? document.getElementById("tooltipdiv") : tooltipdiv
				dropmenuobj.innerHTML=menucontents

				if (ie4||ns6){
					showhide(dropmenuobj.style, e, "visible", "hidden")
					dropmenuobj.x=getposOffset(obj, "left")
					dropmenuobj.y=getposOffset(obj, "top")
					dropmenuobj.style.left=dropmenuobj.x-clearbrowseredge(obj, "rightedge")+"px"
					dropmenuobj.style.top=dropmenuobj.y-clearbrowseredge(obj, "bottomedge")+obj.offsetHeight+"px"
					}
				}

				function hidetip(e){
				if (typeof dropmenuobj!="undefined"){
				if (ie4||ns6)
					dropmenuobj.style.visibility="hidden"
				}
				}

				function delayhidetip(){
				if (ie4||ns6)
				delayhide=setTimeout("hidetip()",disappeardelay)
				}

				function clearhidetip(){
					if (typeof delayhide!="undefined")
						clearTimeout(delayhide)
				}	
				function SetUpload_imagePath(path)
				{
					if(document.getElementById("upload_image")!=null)
						document.getElementById("upload_image").src="upload.asp?<%=setting %>&Theme=<%=Theme%>&FP="+path+"&Type=Image"
				}				

	</script>
</head>
<body>
	<% Response.Write Showbrowse_Img(Server.MapPath(c_ImageGalleryPath & folpath)) %>	
	<%
		if CBool(AllowUpload) then
	%>
	<div style="padding-top:5px;">
	<input type="hidden" id="Hidden1" name="Hidden1">
	<iframe src="upload.asp?<%=setting %>&Theme=<%=Theme%>&FP=<%=c_ImageGalleryPath & folpath%>&Type=Image" id="upload_image" frameborder="0" scrolling="auto" style="width:100%;height:65px"></iframe>
    <%
		end if
	%>	
	</div>
	<input type="hidden" id="TargetUrl" name="TargetUrl" />		
</body>
</html>

