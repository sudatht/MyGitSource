<%
'HTML Editor: File 3/3.
'The HTML editor consists of 3 files:
' - inc_HTML_editor_javascript  (all the javascripts necessary, events, initializations etc.)
' - inc_HTML_editor_main		(the editor object, menu bars and buttons)
' - inc_HTML_editor_end			(Menubar initialization, a couple of forms temp forms)
'This include file must be included last in the asp-file, just before the </BODY> tag.
%>
<!-- HTML Editor specific forms-->
<form name="linkform">
	<input type="hidden" name="linkname" value="null">
	<input type="hidden" name="linktarget" value="null">
	<input type="hidden" name="linkprotocol" value="null">
</form>

<form NAME="UploadForm" method="POST">
	<input type="hidden" name="LinkTitle">
	<input type="hidden" name="Message">
	<input type="hidden" name="BuildNumber" LANGUAGE=javascript onkeypress="return BuildNumber_onkeypress()" onkeypress="return BuildNumber_onkeypress()" onkeypress="return BuildNumber_onkeypress()" onkeypress="return BuildNumber_onkeypress()">
</form>

<!-- Toolbar Code File. Note: This must always be the last thing on the page -->
<script LANGUAGE="Javascript" SRC="dhtml/Toolbars/toolbarsStatic.asp?lEditorWidth=<%=lEditorWidth%>&lLeftPadding=<%=lLeftPadding%>&lTopPadding=<%=lTopPadding%>&lEditorWidth=<%=lEditorWidth%>&lEditorHeight=<%=lEditorHeight%>">
</script>