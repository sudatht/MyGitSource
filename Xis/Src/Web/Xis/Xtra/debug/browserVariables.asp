<%option explicit%>
<HTML>
	<HEAD>
		<TITLE>Browser Properties</TITLE>
		<style type="text/css">
			SPAN.MenuCategorySelected {margin:0px;padding-left:3px; border-top: 1px solid black; border-left:1px solid black; border-right:2px solid black; border-bottom:none }
			SPAN.MenuCategory		 {margin:0px;padding-right:3px; border-top: 1px solid black; border-left:1px solid black; border-right:2px solid black; border-bottom:1px solid black }
		</style>		
	</HEAD>
	<BODY BGCOLOR=#FFFFFF>
		<!--#INCLUDE FILE="top_menu.asp"-->
		<% 
		dim bc
		Set bc = Server.CreateObject("MSWC.BrowserType") 
		%>
		<H3>The following is a list of properties of your browser:</H3>
		<TABLE border="1">
			<TR><TD>Browser Type</TD>		<TD><%= bc.Browser %></TD>
			<TR><TD>What Version</TD>		<TD><%= bc.Version %></TD>
			<TR><TD>Major Version</TD>		<TD><%= bc.Majorver %></TD>
			<TR><TD>Minor Version</TD>		<TD><%= bc.Minorver %></TD>
			<TR><TD>Frames</TD>			<TD><%= CStr((bc.Frames)) %></TD>
			<TR><TD>Tables</TD>			<TD><%= CStr((bc.Tables)) %></TD>
			<TR><TD>Cookies</TD>			<TD><%= CStr((bc.cookies)) %></TD>
			<TR><TD>Background Sounds</TD>		<TD><%= CStr((bc.BackgroundSounds)) %></TD>
			<TR><TD>VBScript</TD>			<TD><%= CStr((bc.VBScript)) %></TD>
			<TR><TD>JavaScript</TD>			<TD><%= CStr((bc.Javascript)) %></TD>
		</TABLE>
	</BODY>
</HTML>