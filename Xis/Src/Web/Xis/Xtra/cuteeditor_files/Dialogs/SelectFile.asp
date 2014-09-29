<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
		<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<!--[if IE]>
			<link href="../Style/IE.css" type="text/css" rel="stylesheet" />
		<![endif]-->
		<style type="text/css">
	    #upload_image {height:60; VISIBILITY: inherit; Z-INDEX: 2}
		</style>
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>
	    <title><%= GetString("Browse") %> </title>
	    <%
		    Response.Expires = -1
    				
		    Dim Current_FilesGalleryPath
		    Current_FilesGalleryPath=FilesGalleryPath
		    if Request.QueryString("MP") <> "" then
			    Current_FilesGalleryPath = Request.QueryString("MP")
		    End if
		 %>
	</head>
	<body>
		<div id="container">
			<table border="0" cellspacing="0" cellpadding="0" width="100%">
	            <tr>
		            <td style="white-space:nowrap; width:250px">
		            </td>
		            <td valign="bottom" style="width:200px">
<img src="../Images/newfolder.gif" id="btn_CreateDir" onclick="CreateDir_click();" title="<%= GetString("Createdirectory") %>" <% if CBool(AllowCreateFolder) then %> class="dialogButton" <% else %> class="CuteEditorButtonDisabled" <% end if%> onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
<img src="../Images/zoom_in.gif" id="btn_zoom_in" onclick="Zoom_In();" title="<%= GetString("ZoomIn") %>" class="dialogButton" onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
<img src="../Images/zoom_out.gif" id="btn_zoom_out" onclick="Zoom_Out();" title="<%= GetString("ZoomOut") %>" class="dialogButton" onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
<img src="../Images/Actualsize.gif" id="btn_Actualsize" onclick="Actualsize();" title="<%= GetString("ActualSize") %>" class="dialogButton" onmouseover="CuteEditor_ColorPicker_ButtonOver(this);" /> 
		            </td>
		            <td align="right">
		            </td>	
	            </tr>
            </table>
			<table border="0" cellspacing="0" cellpadding="0" width="100%">
				<tr>
					<td valign="top" style="width:540px">
<iframe src="browse_Document.asp?<%=setting %>&Theme=<%=Theme%>&MP=<%=Current_FilesGalleryPath%>" id="browse_Frame" frameborder="0" scrolling="auto" style="border: buttonshadow 1px solid; width:540px;height:280px"></iframe>		
					</td>
					<td>&nbsp;
					</td>
					<td valign="top" style="width:220px">
						<div style="border: buttonshadow 1px solid; vertical-align: top; overflow: auto; width:100%; HEIGHT: 280px; BACKGROUND-COLOR: white;">
							<div id="divpreview" style="BACKGROUND-COLOR: white;Padding:3;">
							</div>
						</div>
					</td>
				</tr>
				<tr>
					<td colspan="3" style="height:2">
					</td>
				</tr>
			</table>
			<div style="text-align:center">
			<table border="0" cellspacing="0" cellpadding="0" width="100%">
				<tr>
					<td valign="top">
							 <table border="0" cellpadding="5" cellspacing="0" width="100%">
								    <tr>
									    <td valign="top">
			    			              <fieldset>
							                <legend><%= GetString("Insert") %></legend>
							                        <table border="0" cellpadding="5" cellspacing="0" width="100%">
								                        <tr>
									                    <td valign="middle"><%= GetString("Url") %>:</td>
									                    <td>
									                        <input type="text" id="TargetUrl" style="WIDTH:360px" name="TargetUrl" />
								                        </td>
								                    </tr>
							                    </table>
						                    </fieldset>	    	    
						                        <%
						                            Dim Style_Display_None
							                        if Not CBool(AllowUpload) then
							                            Style_Display_None="Style='Display:none'"
						                            end if
						                        %>
						                    <fieldset id="fieldsetUpload" <%= Style_Display_None %> >
							                    <legend><%= GetString("Upload") %> (Max size <%=MaxDocumentSize%>K)</legend>
<iframe src="upload.asp?<%=setting %>&Theme=<%=Theme%>&FP=<%=Current_FilesGalleryPath%>&Type=document" id="upload_image" frameborder="0" scrolling="auto" style="width:100%;height:50px"></iframe>
									        </fieldset> 
						                    </td>
								    </tr>
								</table>
					</td>
				</tr>
			</table>
			</div>
						<div style="padding-top:10px; text-align:center">
<input class="formbutton" type="button" value="   <%= GetString("Insert") %>   " onclick="do_insert()" id="Button1" /> 
&nbsp;&nbsp;&nbsp; 
<input class="formbutton" type="button" value="   <%= GetString("Cancel") %>  " onclick="do_Close()" id="Button2" />
						</div>
		</div>
			<script type="text/javascript">	
	            var OK = "<%= GetString("OK")%>";
	            var Cancel = "<%= GetString("Cancel")%>";
	            var InputRequired = "<%= GetString("InputRequired")%>";
	            var ValidID = "<%= GetString("ValidID")%>";
	            var ValidColor = "<%= GetString("ValidColor")%>";
	            var SelectImagetoInsert = "<%= GetString("SelectImagetoInsert")%>";
	            
	            function UploadSaved(sFileName,path){
		            ResetFields();
		            try{
		    browse_Frame.location.reload();
		    }
		    catch(x)
		    {}
		            TargetUrl.value = sFileName;
		            var ext=sFileName.substring(sFileName.lastIndexOf('.')).toLowerCase();
	                switch(ext)
	                {
		                case ".jpeg":case ".jpg":case ".gif":case ".png":case ".bmp":
		                    setTimeout(function(){do_preview();}, 100); 
			                break;
		                case ".swf":
		                    setTimeout(function(){do_preview();}, 100); 
			                break;
		                case ".avi":case ".mpg":case ".mp3":case ".mpeg":case ".wav":
		                    setTimeout(function(){do_preview();}, 100); 
			                break;
			            default:
		                    setTimeout(function(){do_preview(sFileName);}, 100); 
			                break;
			            
	                }
		            row_click(sFileName,"");
	            }
            	
	            function Refresh(path)
	            {
		            browse_Frame.location="browse_Document.asp?<%=setting %>&Theme=<%=Theme%>&MP=<%=FilesGalleryPath%>&loc="+path+"";
	            }
	            function CreateDir_click()
	            {
		            <%
			            if not CBool(AllowCreateFolder) then
		            %>
			            alert("<%= GetString("Disabled")%>");
			            return;
		            <%
			            End If
		            %>		    
	                if(Browser_IsIE7())
	                {
		                IEprompt(promptCallback,'<%= GetString("SpecifyNewFolderName")%>', "");		
		                function promptCallback(dir)
		                {
			                var tempPath = browse_Frame.location;	
			                tempPath = tempPath + "&action=createfolder&foldername="+dir;
			                browse_Frame.location = tempPath;		
		                }
	                }
	                else
	                {
		                var dir=prompt("<%= GetString("SpecifyNewFolderName")%>","")
		                if(dir)
		                {
			                var tempPath = browse_Frame.location;	
			                tempPath = tempPath + "&action=createfolder&foldername="+dir;
			                browse_Frame.location = tempPath;			
		                }
	                }
	            }
	            function row_click(path,html)
	            {	
		            ResetFields();
		            TargetUrl.value=path;
		            do_preview(html);
	            }	    
        		
	            function SetUpload_FolderPath(path)
	            {	if(path.substring(path.length-1, path.length)=='/')
		            {
			            path=path.substring(0, path.length-1);
		            }
		            upload_image.src="upload.asp?<%=setting %>&Theme=<%=Theme%>&FP="+path+"&Type=document";
	            }	
	        </script>
		<script type="text/javascript" src="../Scripts/Dialog/DialogFoot.js"></script>
		<script type="text/javascript" src="../Scripts/Dialog/Dialog_SelectFile.js"></script>
	</body>
</html>