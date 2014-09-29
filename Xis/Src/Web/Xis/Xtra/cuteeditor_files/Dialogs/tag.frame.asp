<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<%
    function IsTagPattern(tagname,pattern)
        tagname=lcase(tagname)
        pattern=lcase(pattern)
		IsTagPattern= false
        if tagname = pattern then
            IsTagPattern= true
        else
            dim tagArray,j, str
            tagArray = Split(pattern,",")
			for j = 0 to Ubound(tagArray)
			    str	= trim(tagArray(j))	
			    if str="*" then
			        IsTagPattern= true
			    elseif str=tagname then
			        IsTagPattern= true
			    elseif str="-"&tagname then
			        IsTagPattern= false    
			    end if
			    if IsTagPattern then
			        Exit for	
			    end if
			next			
        end if    
    end function
    SUB ReadDisplayFile(FileToRead)
        whichfile=server.mappath(FileToRead)
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set thisfile = fs.OpenTextFile(whichfile, 1, False)
        tempSTR=thisfile.readall
        response.write tempSTR
        thisfile.Close
        set thisfile=nothing
        set fs=nothing
    END SUB

    function GetTagDisplayName(tagname)
        Select Case lcase(tagname)
			case "img":
			    GetTagDisplayName=GetString("Image")
			case "object":
			    GetTagDisplayName=GetString("ActiveXObject")
			case "inserttable":
			    GetTagDisplayName=GetString(tagname)
			Case else
			    GetTagDisplayName=tagname
		End Select    
    end function

    dim nocancel,tagName,doc,tabcontrol,tabtext
    nocancel=false

	if Request.QueryString("NoCancel")="True" then
		nocancel=true
    end if		
	tagName=Request.QueryString("Tag")
	tabName=Request.QueryString("Tab")	
	tabName=""&tabName
	
	if tabName= "" then
	    if lcase(tagName) = "table" then
	        tabName = "InsertTable"
	    else
	        tabName=tagName
	    end if
	End if
	if lcase(tabName)= "textarea" then
	    tabName="TextBox"
	End if
	
	set doc = server.CreateObject("Microsoft.XMLDOM")	
	
	' For ASP, wait until the XML is all ready before continuing
	doc.async = False
	doc.Load(Server.MapPath("tag.config"))
%>
<html>
	<head>
		<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
		<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>
		<%
		    if nocancel then
		        response.Write "<script type=text/javascript>top.nocancel=true;</script>"
		    else
		        response.Write "<script type=text/javascript>top.nocancel=false;</script>"		    
		    End if		
		%>
		<script type="text/javascript" src="../Scripts/Dialog/Dialog_TagHead.js"></script>
	</head>
	
	<body>
		
		<div id="container">
			<div class="tab-pane-control tab-pane" id="controlparent">
				<div class="tab-row">
							<% 
						'	Response.Write Request.ServerVariables("QUERY_STRING")
							Dim index, isactive
							index = 0
							Dim Nodes,objNode,objText,objPattern,objTab,objControl
		                    set Nodes = doc.DocumentElement.selectNodes("//configuration/*")		                    
		                    
		                    For Each objNode in Nodes
			                    With objNode.Attributes
				                    set objText = .GetNamedItem("text") 
				                    set objPattern = .GetNamedItem("pattern") 
				                    set objTab = .GetNamedItem("tab") 
				                    set objControl = .GetNamedItem("control") 
				                    
				                    isactive = false
				                    if IsTagPattern(tagName,objPattern.Text) then
				                   ' Response.Write objTab.Text &"<br>"
				                   ' Response.Write tabName&"<br>"
				                   ' Response.Write tagName&"<br>"
				                        if index=0 and request.QueryString("Tab")&"" ="" then
				                            isactive = true
				                        end if
				                        if objTab.Text=tabName then
				                            isactive = true
			                            end if
				                        dim url				                        
				                        url=request.ServerVariables("URL")				                        
				                        url=url&"?Tag="&request.QueryString("Tag")		                        
				                        url=url&"&Tab="&objTab.Text	                        
				                        url=url&"&UC="&request.QueryString("UC")	           
				                        url=url&"&Theme="&request.QueryString("Theme")	
				                        url=url&"&setting="&request.QueryString("setting")	
				                        if isactive then
				                            tabcontrol=objControl.Text
				                            tabtext=objText.Text
				                            Response.Write "<h2 class='tab selected'>"
				                            Response.Write "<a tabindex='-1' href='"&url&"'>"
								            Response.Write "<span style='white-space:nowrap;' >"
									        Response.Write tabtext
								            Response.Write "</span>"
							                Response.Write "</a>"
							                Response.Write "</h2>"
				                        else				   
				                            tabtext=objText.Text        
				                            Response.Write "<h2 class='tab'>"         
				                            Response.Write "<a tabindex='-1' href='"&url&"'>"
								            Response.Write "<span style='white-space:nowrap;' >"
									        Response.Write tabtext
								            Response.Write "</span>"
							                Response.Write "</a>"
							                Response.Write "</h2>"
				                        end if 		
			                            index = index + 1	
				                    end if				                    		                    
			                    End With
			                    
		                    Next 
		                    %>
						</div>
		                <br>
			            <div class="tab-page" style="WIDTH:450px">
	                    <%
	                        if tabcontrol <> "" then			        		 
	                            Server.Execute "Tag/"&tabcontrol
	                        else       		 
	                            Server.Execute "Tag/tag_common.asp"			    
	                        end if
	                    %>			
		                </div>
		          </div>
		        <br>
		        <div id="container-bottom">
			        <input type="button" id="btn_editinwin" class="formbutton" value="<%= GetString("EditHtml") %>">
			        &nbsp;&nbsp;&nbsp; <input type="button" id="btnok" class="formbutton" value="<%= GetString("OK") %>" style="WIDTH:80px">&nbsp;
			        <input type="button" id="btncc" class="formbutton" value="<%= GetString("Cancel") %>" style="WIDTH:80px">
		        </div>
		</div>
	</body>
	<script type="text/javascript" src="../Scripts/Dialog/DialogFoot.js"></script>
	<script type="text/javascript" src="../Scripts/Dialog/Dialog_TagFoot.js"></script>
</html>