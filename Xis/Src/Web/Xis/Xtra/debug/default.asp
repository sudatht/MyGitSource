<%option explicit%>
<%
dim var
%>
<html>
    <head>
        <title>Application/session variables</title>
		<style type="text/css">
			SPAN.MenuCategorySelected {margin:0px;padding-left:3px; border-top: 1px solid black; border-left:1px solid black; border-right:2px solid black; border-bottom:none }
			SPAN.MenuCategory		 {margin:0px;padding-right:3px; border-top: 1px solid black; border-left:1px solid black; border-right:2px solid black; border-bottom:1px solid black }
		</style>
    </head>
    <body>
		<!--#INCLUDE FILE="top_menu.asp"-->	    
		<h3>Application/session variables</h3>
		<p>
		<strong>Web file-root is:&#160;</strong><%=server.MapPath("\")%><br>
		</p>
		<p>
			<strong>There are&#160;<%=Session.Contents.Count%>&#160;session variables.</strong><br>
			<%
			if (Session.Contents.Count > 0) then		
				%>
				<strong>Session variables:</strong><br>		
				<%		
				for each var in session.Contents
					Response.Write  "- " & var & " er """ & session(var) & """<br>"
				next 
			end if
			%>
		</p>
		<p>
			<strong>There are&#160;<%=application.Contents.Count%>&#160;application variables.</strong><br>
			<%
			if (application.Contents.Count > 0) then
				%>
				<strong>Application variables:</strong><br>		
				<%
				for each var in application.Contents
					Response.Write  "- " & var & " er """ & application(var) & """<br>"
				next 
			end if
			%>
		</p>
		<p>
			<strong>There are&#160;<%=Request.ServerVariables.Count%>&#160;server variables.</strong><br>
			<%
			if (Request.ServerVariables.Count > 0) then
				%>
				<strong>Server variables:</strong><br>
				<table border="1" cellspacing="0" cellpadding="2">
					<%
					for each var in Request.ServerVariables
						Response.Write "<tr><td>" &  var & "</td><td>" & Request.ServerVariables(var).Item & "</td></tr>"
					next
					%>
				</table>
				<%
			end if								
			%>
		</p>
	</body>
</html>