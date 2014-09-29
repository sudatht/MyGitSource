<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<% 

dim ps
ps=array("Font","Text","Border","Layout","Background","Other") 

dim activepanel
activepanel=request.QueryString("Style") &""
if activepanel ="" then
	activepanel="Font"
end if

dim activetext
activetext=""
%>

<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_Style.js"></script>
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td style="width:94" valign="top" id="navtd">
			<%
			
			dim px, iscurrent
			iscurrent = false
			For Each px in ps
                dim url				                        
                url=request.ServerVariables("URL")
                				                        
                url=url&"?Tag="&request.QueryString("Tag")     
                url=url&"&Tab=Style"			                        
                url=url&"&Style="&px	                        
                url=url&"&UC="&request.QueryString("UC")	           
                url=url&"&Theme="&request.QueryString("Theme")	
                url=url&"&setting="&request.QueryString("setting")	
			
			    if activepanel = px then
			        activetext=px
                    Response.Write "<a tabindex='-1' class='ActiveStyleNav' href='"&url&"'>"
		            Response.Write "<img alt='' align='absmiddle' src='../Images/style."&px&".gif' style='border:0; vertical-align:inherit;'>"
			        Response.Write px
	                Response.Write "</a>"
                else				   
                    Response.Write "<a tabindex='-1' class='StyleNav' href='"&url&"'>"
		            Response.Write "<img alt='' align='absmiddle' src='../Images/style."&px&".gif' style='border:0; vertical-align:inherit;'>"
			        Response.Write px
	                Response.Write "</a>"
			    end if 			    
			Next 
			%>
		</td>
		<td style="white-space:nowrap;width:10" ></td>
		<td valign="top">
		    <%
             ' response.Write Server.MapPath("Tag/Tag_Style_"&activepanel&".asp")
                Server.Execute "Tag/Tag_Style_"&activepanel&".asp"          
            %>	
		</td>
	</tr>
</table>
<script type="text/javascript">
var OxOc1e3=["ondblclick","navtd","shiftKey","style"];Window_GetElement(window,OxOc1e3[1],true)[OxOc1e3[0]]=function (){if(event[OxOc1e3[2]]){alert(element[OxOc1e3[3]].cssText);} ;} ;
</script>