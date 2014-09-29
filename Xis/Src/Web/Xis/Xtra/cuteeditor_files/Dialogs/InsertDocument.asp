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
	    <title><%= GetString("InsertDocument") %> </title>
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
					<td valign="top" style="width:520px; height:300px">
						<iframe src="browse_Document.asp?<%=setting %>&Theme=<%=Theme%>&MP=<%=Current_FilesGalleryPath%>" id="browse_Frame" frameborder="0" scrolling="auto" style="border:1.5pt inset;overflow: auto;width:100%;height:100%"></iframe>		
					</td>
					<td>&nbsp;
					</td>
					<td valign="top" style="width:240px;HEIGHT: 300px;">
						<div style="border: buttonshadow 1px solid; vertical-align: top; overflow: auto; width:100%; HEIGHT: 100%; BACKGROUND-COLOR: white;">
							<div id="divpreview" style="Padding:2px;">
							</div>
						</div>
					</td>
				</tr>
				<tr>
					<td colspan="3" style="height:2">
					</td>
				</tr>
			</table>
			<table border="0" cellspacing="0" cellpadding="0" width="100%">
				<tr>
					<td valign="top">
						<fieldset>
							<legend>
								<%= GetString("Properties") %></legend>
							<table border="0" cellpadding="4" cellspacing="0" id="Table3">
								<tr>
									<td>
										<table border="0" cellpadding="1" cellspacing="0" id="Table4" class="normal">
											<tr>
												<td><%= GetString("Target") %></td>
												<td><select id="sel_target">
														<option value=""><%= GetString("NotSet") %></option>
														<option value="_blank"><%= GetString("Newwindow") %></option>
														<option value="_self"><%= GetString("Samewindow") %></option>
														<option value="_top"><%= GetString("Topmostwindow") %></option>
														<option value="_parent"><%= GetString("Parentwindow") %></option>
													</select>
												</td>
											</tr>
								            <tr>
									            <td><%= GetString("Color") %>:</td>
									            <td style="text-align:left">
										            <input type="text" id="inp_color" name="inp_color" size="7" style="WIDTH:57px;" />
										            <img title="<%= GetString("Preview") %>" src="../Images/colorpicker.gif" id="inp_color_preview" style="vertical-align:top;" />
									            </td>
								            </tr>
											<tr>
												<td><%= GetString("CssClass") %>:</td>
												<td><input id="inc_class" type="text" size="14" style="WIDTH:80px" name="inc_class" /></td>
											</tr>
											<tr>
												<td><%= GetString("ID") %>:</td>
												<td><input id="inp_id" type="text" size="14" style="WIDTH:80px" name="inp_id" /></td>
											</tr>
											<tr>
												<td><%= GetString("TabIndex") %>:</td>
												<td><input id="inp_index" type="text" size="14" maxlength="5" style="WIDTH:80px" name="inp_index"
														onkeypress="return CancelEventIfNotDigit()" /></td>
											</tr>
											<tr>
												<td><%= GetString("AccessKey") %>:</td>
												<td><input id="inp_access" type="text" size="14" maxlength="1" style="WIDTH:80px" name="inp_access"
														onkeypress="return CancelEventIfNotDigit()" /></td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</fieldset>
					</td>
					<td style="width:5">
					</td>
					<td valign="top">
						<fieldset>
							<table border="0" cellpadding="4" cellspacing="0">
								<tr>
									<td>
										<table id="Table8" cellspacing="0" cellpadding="1" border="0" class="normal">
											<tr>
												<td valign="middle"><%= GetString("Url") %>:</td>
												<td><input id="TargetUrl" onpropertychange="do_preview()" type="text" size="60" name="TargetUrl"
														/>
												</td>
											</tr>
											<tr>
												<td valign="middle"><%= GetString("Title") %>:</td>
												<td valign="middle"><input id="inp_title" type="text" size="60" name="inp_title" /></td>
											</tr>
										</table>
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
							    <table border="0" cellspacing="2" cellpadding="0" width="100%" class="normal">
								    <tr>
									    <td style="width:8">
									    </td>
								    </tr>
								    <tr>
									    <td valign="top">
<iframe src="upload.asp?<%=setting %>&FP=<%=Current_FilesGalleryPath%>&Type=document&Theme=<%=Theme%>" id="upload_image" frameborder="0" scrolling="auto" style="width:100%;"></iframe>
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
	        <script type="text/javascript" src="../Scripts/Dialog/Dialog_InsertDocument.js"></script>
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
		            upload_image.location="upload.asp?<%=setting %>&FP="+path+"&Type=document";
	            }	
	        </script>
	</body>
</html>