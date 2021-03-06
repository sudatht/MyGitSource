<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>		
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<!--[if IE]>
			<link href="../Style/IE.css" type="text/css" rel="stylesheet" />
		<![endif]-->
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>
		<style type="text/css">
		#framepreview {
			width: 100%;
			height: 100%;
			overflow:hidden;
			text-align: left;
			vertical-align: top;
			padding: 0px;
			margin: 0px;
			zoom: 50%;
			background-color: white;
		}	
		.row { HEIGHT: 22px }
		.cb { VERTICAL-ALIGN: middle }
		.itemimg { VERTICAL-ALIGN: middle }
		.editimg { VERTICAL-ALIGN: middle }
		.cell1 { VERTICAL-ALIGN: middle }
		.cell2 { VERTICAL-ALIGN: middle }
		.cell3 { PADDING-RIGHT: 4px; VERTICAL-ALIGN: middle; TEXT-ALIGN: right }
		.cb { }
		</style>
		
	<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
	<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
	<title><%= GetString("InsertTemplate") %></title>
	<%
		Response.Expires = -1
		
		Dim Current_TemplateGalleryPath
		Current_TemplateGalleryPath=TemplateGalleryPath
		if Request.QueryString("MP") <> "" then
			Current_TemplateGalleryPath = Request.QueryString("MP")
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
					<td valign="top" style="width:260px;height:240px;">
						<iframe src="browse_template.asp?<%=setting %>&Theme=<%=Theme%>&MP=<%=Current_TemplateGalleryPath%>" id="browse_Frame" frameborder="0" scrolling="auto" style="border:1.5pt inset;width:100%;height:246px"></iframe>		
					</td>
					<td>&nbsp;&nbsp;</td>
					<td valign="top">
						<div style="BORDER: 1.5pt inset; VERTICAL-ALIGN: top; OVERFLOW: auto; WIDTH: 380px; HEIGHT: 250px; BACKGROUND-COLOR: white; TEXT-ALIGN: center">
						    <iframe id="framepreview" src="../template.asp"></iframe>
						</div>
					</td>
				</tr>
				<tr>
					<td colspan="2" style="height:2">
					</td>
				</tr>
			</table>
			<table border="0" cellspacing="0" cellpadding="0" width="100%">
				<tr>
					<td valign="top">
					</td>
					<td style="width:10">
					</td>
					<td valign="top">
					    <input type="hidden" id="hiddenHTML" name="hiddenDirectory" /> 
						<input type="hidden" id="TargetUrl" name="TargetUrl" />						
						<%
						    Dim Style_Display_None
							if Not CBool(AllowUpload) then
							    Style_Display_None="Style='Display:none'"
						    end if
						%>
						<fieldset id="fieldsetUpload" <%= Style_Display_None %> >
							<legend><%= GetString("Upload") %> (Max size <%=MaxTemplateSize%>K)</legend>
							<table border="0" cellspacing="2" cellpadding="0" width="100%" class="normal">
								<tr>
									<td style="width:8">
									</td>
								</tr>
								<tr>
									<td valign="top">
<iframe src="upload.asp?<%=setting %>&Theme=<%=Theme%>&FP=<%=Current_TemplateGalleryPath%>&Type=template" id="upload_image" frameborder="0" scrolling="auto" style="width:100%; height:60px"></iframe>

									</td>
								</tr>
							</table>
						</fieldset>
						<div style="padding-top:10px; text-align:center">
<input class="formbutton" type="button" value="   <%= GetString("Insert") %>   " onclick="do_insert()" id="Button1" /> 
&nbsp;&nbsp;&nbsp; 
<input class="formbutton" type="button" value="   <%= GetString("Cancel") %>  " onclick="do_Close()" id="Button2" />
						</div>
					</td>
				</tr>
			</table>
		</div>
		<script type="text/javascript" src="../Scripts/Dialog/DialogFoot.js"></script>
		<script type="text/javascript" src="../Scripts/Dialog/Dialog_InsertTemplate.js"></script>
	    <script type="text/javascript">	
	        var OK = "<%= GetString("OK")%>"
	        var Cancel = "<%= GetString("Cancel")%>"
	        var InputRequired = "<%= GetString("InputRequired")%>"
	        var ValidID = "<%= GetString("ValidID")%>"
	        var ValidColor = "<%= GetString("ValidColor")%>"
	        var SelectImagetoInsert = "<%= GetString("SelectImagetoInsert")%>";
	        
	        function UploadSaved(sFileName,path){
		        ResetFields();
		        try{
		            browse_Frame.location.reload();
		            }
		            catch(x)
		            {}
		        TargetUrl.value = sFileName;
		        setTimeout(function(){do_preview(sFileName);}, 100); 
		        row_click(sFileName);
	        }
        	
	        function Refresh(path)
	        {
		        browse_Frame.location="browse_template.asp?<%=setting %>&Theme=<%=Theme%>&MP=<%=TemplateGalleryPath%>&loc="+path+"";
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
	        function row_click(path)
	        {
		        framepreview.location=path;
		        setTimeout(do_preview,500);
	        }	    
    		
	        function SetUpload_FolderPath(path)
	        {	if(path.substring(path.length-1, path.length)=='/')
		        {
			        path=path.substring(0, path.length-1);
		        }
		        upload_image.location="upload.asp?<%=setting %>&FP="+path+"&Theme=<%=Theme%>&Type=template";
	        }	
	    </script>	
	</body>
</html>