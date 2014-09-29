<HTML>
	<HEAD>
		<TITLE>Cookies</TITLE>
		<style type="text/css">
			SPAN.MenuCategorySelected {margin:0px;padding-left:3px; border-top: 1px solid black; border-left:1px solid black; border-right:2px solid black; border-bottom:none }
			SPAN.MenuCategory		 {margin:0px;padding-right:3px; border-top: 1px solid black; border-left:1px solid black; border-right:2px solid black; border-bottom:1px solid black }
		</style>
	</HEAD>
	<BODY BACKGROUND="/gfx/offwhite.gif" BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#008000" VLINK="#008000" ALINK="#808000">
		<!--#INCLUDE FILE="top_menu.asp"-->
		<h3>Cookies</h3>
		<%
		if (Request.Cookies.Count > 0) then
			%>
			<TABLE WIDTH="500" BORDER="1">
				<TR>
					<TD WIDTH="100"></TD>
					<TD WIDTH="400"></TD>
				</TR>
				<% 
				For Each x in Request.Cookies 
					%>
					<TR>
						<TD><FONT FACE="arial" SIZE=-1><% RW x %></FONT></TD>
						<TD><FONT FACE="arial" SIZE=-1><% RW Request.Cookies(x) %></FONT></TD>
					</TR>
					<% 
				Next 
				%>
			<TABLE>
			<%
		else
			%>
			<p>No cookies.</p>
			<%
		end if
		%>
	</BODY>
</HTML>

