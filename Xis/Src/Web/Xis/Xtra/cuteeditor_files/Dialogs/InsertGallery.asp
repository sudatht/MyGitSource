<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>		
	    <title><%= GetString("InsertImage") %> </title>
	    <%
		    Response.Expires = -1
    		
		    Dim Current_ImageGalleryPath
		    Current_ImageGalleryPath=ImageGalleryPath
		    if Request.QueryString("GP") <> "" then
			    Current_ImageGalleryPath = Request.QueryString("GP")
		    End if
	    %>
		<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
		<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<!--[if IE]>
			<link href="../Style/IE.css" type="text/css" rel="stylesheet" />
		<![endif]-->
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>
		<style type="text/css">
	        #browse_Img_gallery {border:0;height:470px;width:100%;vertical-align: top;overflow: auto;}
			A:link { COLOR: #000099 }	
			A:visited { COLOR: #000099 }	
			A:active { COLOR: #000099 }	
			A:hover { COLOR: darkred }
			#tooltipdiv { border: black 1px solid; padding: 2px; z-index: 100; font: menu; position: absolute }	
			#ThumbList1_MyList IMG { border:solid 2px #cccccc;}	
			IMG.spring_image { BEHAVIOR: url(htc.htc); CURSOR: pointer; -moz-binding: url("xbl.xml#MyComponent") }	
			IMG.spring_image_popup { border:solid 2px #ff0000; DISPLAY: none; Z-INDEX: 99999; LEFT: 0px; FLOAT: none; MARGIN: 0px; CURSOR: pointer; POSITION: absolute; TOP: 0px }	
		</style>
	</head>
	<body>
		<div id="container">    
		    		
	        <table width="100%" cellspacing="0" cellpadding="2" border="0" align="center">
				<tr>
					<td valign="middle" style="text-align:right;padding-right:10px;">	  
						<img class="dialogButton" id="upbtn" src="../Images/parentfolder.gif" style="VERTICAL-ALIGN: middlel;cursor:hand" title="Go up one level" />
					</td>
				</tr>
				<tr>
					<td valign="top">     
					    <iframe src="browse_Img_gallery.asp?<%=setting %>&GP=<%=Current_ImageGalleryPath%>&Theme=<%=Theme%>" id="browse_Img_gallery" frameborder="0" scrolling="auto"></iframe>
				        <input type="hidden" id="TargetUrl" onpropertychange="do_preview()" name="TargetUrl" style="width:300px;"/>
					</td>
				</tr>
				<tr>
					<td align="right" style="padding-top:10px; text-align:center">
<input type="button" value="<%= GetString("Cancel") %>" class="formbutton" onclick="cancel();">
					</td>
				</tr>
			</table>
		</div>
	</body>
	<script type="text/javascript">
	    var OxO6c84=[];var OxO51a5=[];var currentfolder='browse_Img_gallery.asp?<%=setting %>&GP=<%=Current_ImageGalleryPath %>&Theme=<%=Theme%>';var setting='<%=setting %>';	
	</script>
	<script type="text/javascript" src="../Scripts/Dialog/DialogFoot.js"></script>
	<script type="text/javascript" src="../Scripts/Dialog/Dialog_InsertGallery.js"></script>
</html>