<%
dim strTarget
dim strHTMLProperty
dim strFormName

strTarget = request("Targetfld")
if (request("fldType")="textarea") then
	strHTMLProperty = "value"
else
	strHTMLProperty = "innerHTML"
end if

strFormName = request("frmName")
%>
<html>
<head>
<title>HTML Editor</title>
<!-- Styles -->
<link REL="stylesheet" TYPE="text/css" HREF="dhtml/toolbars/toolbars.css">
<style>
	.editorMain			{ width:50%; height:80%; }
</style>

<!-- Script Functions and Event Handlers -->
<script LANGUAGE="JavaScript" SRC="dhtml/dhtmled.js">
</script>

<script language="vbscript">

	Sub MENU_FILE_SAVE_onclick()

		If Len(tbContentElement.DOM.body.innerHTML) = 0 Then
		  filteredHTML = ""
		Else
			filteredHTML = tbContentElement.FilterSourceCode (tbContentElement.DOM.body.innerHTML)
		End If

		if (window.opener is nothing) then
			UploadForm.submit
		else
			window.Opener.document.all.<%=strTarget%>.<%=strHTMLProperty%> = filteredHTML
			<%
			if len(trim(strFormName))>0 then
			%>
			window.Opener.document.all.<%=strFormName%>.fireEvent("onSubmit")
			'window.Opener.document.all.<%=strFormName%>.submit
			<%
			end if
			%>
			window.close
		end if

	End Sub

</script>

<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

	//
	// Constants
	//
	var MENU_SEPARATOR = ""; // Context menu separator

	// Special context menu commands
	BUILD_CONTROL = 1
	BUILD_TRIEDITDLL = 2
	BUILD_SAMPLE = 3
	BUILD_HEADER = 4
	BUILD_SETUP = 5
	BUILD_LAST_COMMAND = 5

	//
	// Globals
	//
	var QueryStatusToolbarButtons = new Array();
	var QueryStatusEditMenu = new Array();
	var QueryStatusFormatMenu = new Array();
	var QueryStatusHTMLMenu = new Array();
	var QueryStatusTableMenu = new Array();
	var QueryStatusZOrderMenu = new Array();
	var ContextMenu = new Array();
	var GeneralContextMenu = new Array();
	var TableContextMenu = new Array();
	var AbsPosContextMenu = new Array();

	var docInit = false; // flag for loading initial document

	//
	// Utility functions
	//

	// Converts double quotes in a string to HTML quot; entities
	function encodeHTMLQuotes(strIn)
	{
	  var strOut;

	  while (strOut != strIn)
	  {
	    strOut = strIn;
	    strIn = strIn.replace('"', 'quote;');
	  }
	  return strOut;
	}

	// Constructor for custom object that represents an item on the context menu
	function ContextMenuItem(string, cmdId) {
	  this.string = string;
	  this.cmdId = cmdId;
	}

	// Constructor for custom object that represents a QueryStatus command and
	// corresponding toolbar element.
	function QueryStatusItem(command, element) {
	  this.command = command;
	  this.element = element;
	}

	//
	// Event handlers
	//
	function window_onload() {
	  var today, year, beginningOfYear, daysSinceBeginningOfYear;

	  // Initialize the build number to today's build
	  today = new Date();
	  year = today.getYear();
	  beginningOfYear = new Date(year, 0, 0);
	  msSinceBeginningOfYear = today.getTime() - beginningOfYear.getTime();
	  daysSinceBeginningOfYear = Math.floor(msSinceBeginningOfYear / (1000 * 60 * 60 * 24));
	  BuildNumber.value = "1000";

	  // Initialze QueryStatus tables. These tables associate a command id with the
	  // corresponding button object. Must be done on window load, 'cause the buttons must exist.
	  QueryStatusToolbarButtons[0] = new QueryStatusItem(DECMD_BOLD, document.body.all["DECMD_BOLD"]);
	  QueryStatusToolbarButtons[1] = new QueryStatusItem(DECMD_COPY, document.body.all["DECMD_COPY"]);
	  QueryStatusToolbarButtons[2] = new QueryStatusItem(DECMD_CUT, document.body.all["DECMD_CUT"]);
	  QueryStatusToolbarButtons[3] = new QueryStatusItem(DECMD_HYPERLINK, document.body.all["DECMD_HYPERLINK"]);
	  QueryStatusToolbarButtons[4] = new QueryStatusItem(DECMD_ITALIC, document.body.all["DECMD_ITALIC"]);
	  QueryStatusToolbarButtons[5] = new QueryStatusItem(DECMD_JUSTIFYLEFT, document.body.all["DECMD_JUSTIFYLEFT"]);
	  QueryStatusToolbarButtons[6] = new QueryStatusItem(DECMD_JUSTIFYCENTER, document.body.all["DECMD_JUSTIFYCENTER"]);
	  QueryStatusToolbarButtons[7] = new QueryStatusItem(DECMD_JUSTIFYRIGHT, document.body.all["DECMD_JUSTIFYRIGHT"]);
	  QueryStatusToolbarButtons[8] = new QueryStatusItem(DECMD_ORDERLIST, document.body.all["DECMD_ORDERLIST"]);
	  QueryStatusToolbarButtons[9] = new QueryStatusItem(DECMD_PASTE, document.body.all["DECMD_PASTE"]);
	  QueryStatusToolbarButtons[10] = new QueryStatusItem(DECMD_REDO, document.body.all["DECMD_REDO"]);
	  QueryStatusToolbarButtons[11] = new QueryStatusItem(DECMD_UNDERLINE, document.body.all["DECMD_UNDERLINE"]);
	  QueryStatusToolbarButtons[12] = new QueryStatusItem(DECMD_UNDO, document.body.all["DECMD_UNDO"]);
	  QueryStatusToolbarButtons[13] = new QueryStatusItem(DECMD_UNORDERLIST, document.body.all["DECMD_UNORDERLIST"]);
	  QueryStatusEditMenu[0] = new QueryStatusItem(DECMD_UNDO, document.body.all["EDIT_UNDO"]);
	  QueryStatusEditMenu[1] = new QueryStatusItem(DECMD_REDO, document.body.all["EDIT_REDO"]);
	  QueryStatusEditMenu[2] = new QueryStatusItem(DECMD_CUT, document.body.all["EDIT_CUT"]);
	  QueryStatusEditMenu[3] = new QueryStatusItem(DECMD_COPY, document.body.all["EDIT_COPY"]);
	  QueryStatusEditMenu[4] = new QueryStatusItem(DECMD_PASTE, document.body.all["EDIT_PASTE"]);
	  QueryStatusEditMenu[5] = new QueryStatusItem(DECMD_DELETE, document.body.all["EDIT_DELETE"]);
	  QueryStatusHTMLMenu[0] = new QueryStatusItem(DECMD_HYPERLINK, document.body.all["HTML_HYPERLINK"]);
	  QueryStatusHTMLMenu[1] = new QueryStatusItem(DECMD_IMAGE, document.body.all["HTML_IMAGE"]);
	  QueryStatusFormatMenu[0] = new QueryStatusItem(DECMD_FONT, document.body.all["FORMAT_FONT"]);
	  QueryStatusFormatMenu[1] = new QueryStatusItem(DECMD_BOLD, document.body.all["FORMAT_BOLD"]);
	  QueryStatusFormatMenu[2] = new QueryStatusItem(DECMD_ITALIC, document.body.all["FORMAT_ITALIC"]);
	  QueryStatusFormatMenu[3] = new QueryStatusItem(DECMD_UNDERLINE, document.body.all["FORMAT_UNDERLINE"]);
	  QueryStatusFormatMenu[4] = new QueryStatusItem(DECMD_JUSTIFYLEFT, document.body.all["FORMAT_JUSTIFYLEFT"]);
	  QueryStatusFormatMenu[5] = new QueryStatusItem(DECMD_JUSTIFYCENTER, document.body.all["FORMAT_JUSTIFYCENTER"]);
	  QueryStatusFormatMenu[6] = new QueryStatusItem(DECMD_JUSTIFYRIGHT, document.body.all["FORMAT_JUSTIFYRIGHT"]);
	  QueryStatusZOrderMenu[0] = new QueryStatusItem(DECMD_SEND_TO_BACK, document.body.all["ZORDER_SENDBACK"]);
	  QueryStatusZOrderMenu[1] = new QueryStatusItem(DECMD_BRING_TO_FRONT, document.body.all["ZORDER_BRINGFRONT"]);
	  QueryStatusZOrderMenu[2] = new QueryStatusItem(DECMD_SEND_BACKWARD, document.body.all["ZORDER_SENDBACKWARD"]);
	  QueryStatusZOrderMenu[3] = new QueryStatusItem(DECMD_BRING_FORWARD, document.body.all["ZORDER_BRINGFORWARD"]);
	  QueryStatusZOrderMenu[4] = new QueryStatusItem(DECMD_SEND_BELOW_TEXT, document.body.all["ZORDER_BELOWTEXT"]);
	  QueryStatusZOrderMenu[5] = new QueryStatusItem(DECMD_BRING_ABOVE_TEXT, document.body.all["ZORDER_ABOVETEXT"]);

	  // Initialize the context menu arrays.
	  GeneralContextMenu[0] = new ContextMenuItem("Cut", DECMD_CUT);
	  GeneralContextMenu[1] = new ContextMenuItem("Copy", DECMD_COPY);
	  GeneralContextMenu[2] = new ContextMenuItem("Paste", DECMD_PASTE);

	  //if this was opened from another window, transfer existing text to editor
	  if (window.opener != null)
	  {
		if (window.opener.document.all.<%=strTarget%>.<%=strHTMLProperty%> != "")
		{
		document.all.tbContentElement.DocumentHTML = window.opener.document.all.<%=strTarget%>.<%=strHTMLProperty%>;
		}else{
		document.all.tbContentElement.DocumentHTML = "<p></p>";
		};
	  };
	  //Populate listbox containing valid formats
	  PopulateFormatList();
	}

	// Execute our custom commands
	function ExecBuildCommand(command) {
	  var selection, customHTML;

	  switch (command) {
	    case BUILD_CONTROL :
	      customHTML = "<STRONG>DHTMLEdit control changes:<BR></STRONG>";
	      break;

	    case BUILD_TRIEDITDLL :
	      customHTML = "<STRONG>TriEdit dll changes:<BR></STRONG>";
	      break;

	    case BUILD_SAMPLE :
	      customHTML = "<STRONG>Sample changes:<BR></STRONG>";
	      break;

	    case BUILD_HEADER :
	      customHTML = "<STRONG>Header changes:<BR></STRONG>";
	      break;

	    case BUILD_SETUP :
	      customHTML = "<STRONG>Setup changes:<BR></STRONG>";
	      break;
	  }

	  selection = tbContentElement.DOM.selection.createRange();
	  selection.pasteHTML(customHTML);
	}

	function tbContentElement_ShowContextMenu() {
	  var menuStrings = new Array();
	  var menuStates = new Array();
	  var state;
	  var i
	  var idx = 0;

	  // Rebuild the context menu.
	  ContextMenu.length = 0;

	  // Always show general menu
	  for (i=0; i<GeneralContextMenu.length; i++) {
	    ContextMenu[idx++] = GeneralContextMenu[i];
	  }

	  // Is the selection on an absolutely positioned element? Add z-index commands if so
	  if (tbContentElement.QueryStatus(DECMD_LOCK_ELEMENT) != DECMDF_DISABLED) {
	    for (i=0; i<AbsPosContextMenu.length; i++) {
	      ContextMenu[idx++] = AbsPosContextMenu[i];
	    }
	  }

	  // Set up the actual arrays that get passed to SetContextMenu
	  for (i=0; i<ContextMenu.length; i++) {
	    menuStrings[i] = ContextMenu[i].string;
	    if (menuStrings[i] != MENU_SEPARATOR) {
	      if (ContextMenu[i].cmdId > BUILD_LAST_COMMAND) {
	        state = tbContentElement.QueryStatus(ContextMenu[i].cmdId);
	      } else {
	        state = DECMDF_ENABLED;
	      }
	    } else {
	      state = DECMDF_ENABLED;
	    }
	    if (state == DECMDF_DISABLED || state == DECMDF_NOTSUPPORTED) {
	      menuStates[i] = OLE_TRISTATE_GRAY;
	    } else if (state == DECMDF_ENABLED || state == DECMDF_NINCHED) {
	      menuStates[i] = OLE_TRISTATE_UNCHECKED;
	    } else { // DECMDF_LATCHED
	      menuStates[i] = OLE_TRISTATE_CHECKED;
	    }
	  }

	  // Set the context menu
	  tbContentElement.SetContextMenu(menuStrings, menuStates);
	}

	function tbContentElement_ContextMenuAction(itemIndex) {

	  if (ContextMenu[itemIndex].cmdId <= BUILD_LAST_COMMAND) {
	    ExecBuildCommand(ContextMenu[itemIndex].cmdId);
	    return;
	  }
	    tbContentElement.ExecCommand(ContextMenu[itemIndex].cmdId, OLECMDEXECOPT_DODEFAULT);
	}

	// DisplayChanged handler. Very time-critical routine; this is called
	// every time a character is typed. QueryStatus those toolbar buttons that need
	// to be in synch with the current state of the document and update.
	function tbContentElement_DisplayChanged() {
	  var i, s;

	  for (i=0; i<QueryStatusToolbarButtons.length; i++) {
	  s = tbContentElement.QueryStatus(QueryStatusToolbarButtons[i].command);
	  if (s == DECMDF_DISABLED || s == DECMDF_NOTSUPPORTED) {
	      TBSetState(QueryStatusToolbarButtons[i].element, "gray");
	    } else if (s == DECMDF_ENABLED  || s == DECMDF_NINCHED) {
	       TBSetState(QueryStatusToolbarButtons[i].element, "unchecked");
	    } else { // DECMDF_LATCHED
	       TBSetState(QueryStatusToolbarButtons[i].element, "checked");
	    }
	  }

	  s = tbContentElement.QueryStatus(DECMD_GETBLOCKFMT);
	  if (s == DECMDF_DISABLED || s == DECMDF_NOTSUPPORTED) {
	    ParagraphStyle.disabled = true;
	  } else {
	    ParagraphStyle.disabled = false;
	    ParagraphStyle.value = tbContentElement.ExecCommand(DECMD_GETBLOCKFMT, OLECMDEXECOPT_DODEFAULT);
	  }

	  // Load the initial release note into the content element
	  // Make sure the document that this document (builded.htm)
	  // has been loaded so that the BuildNumber.value is initialized
	  if ("complete" == document.readyState && false == docInit)
	  {
	    docInit = true;
	  }
	  // don't think this helps
	  tbContentElement.UseDivOnCarriageReturn = true;

	}

	function DECMD_VISIBLEBORDERS_onclick() {
	  tbContentElement.ShowBorders = !tbContentElement.ShowBorders;
	  tbContentElement.focus();
	}

	function DECMD_UNORDERLIST_onclick() {
	  tbContentElement.ExecCommand(DECMD_UNORDERLIST, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_UNDO_onclick() {
	  tbContentElement.ExecCommand(DECMD_UNDO, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_UNDERLINE_onclick() {
	  tbContentElement.ExecCommand(DECMD_UNDERLINE, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_SNAPTOGRID_onclick() {
	  tbContentElement.SnapToGrid = !tbContentElement.SnapToGrid;
	  tbContentElement.focus();
	}

	function DECMD_SHOWDETAILS_onclick() {
	  tbContentElement.ShowDetails = !tbContentElement.ShowDetails;
	  tbContentElement.focus();
	}

	function DECMD_SELECTALL_onclick() {
	  tbContentElement.ExecCommand(DECMD_SELECTALL, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_REDO_onclick() {
	  tbContentElement.ExecCommand(DECMD_REDO, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_PASTE_onclick() {
	  tbContentElement.ExecCommand(DECMD_PASTE, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_ORDERLIST_onclick() {
	  tbContentElement.ExecCommand(DECMD_ORDERLIST, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_MAKE_ABSOLUTE_onclick() {
	  tbContentElement.ExecCommand(DECMD_MAKE_ABSOLUTE, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_LOCK_ELEMENT_onclick() {
	  tbContentElement.ExecCommand(DECMD_LOCK_ELEMENT, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_OUTDENT_onclick() {
	  tbContentElement.ExecCommand(DECMD_OUTDENT, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_INDENT_onclick() {
	  tbContentElement.ExecCommand(DECMD_INDENT, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_JUSTIFYRIGHT_onclick() {
	  tbContentElement.ExecCommand(DECMD_JUSTIFYRIGHT, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_JUSTIFYLEFT_onclick() {
	  tbContentElement.ExecCommand(DECMD_JUSTIFYLEFT, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_JUSTIFYCENTER_onclick() {
	  tbContentElement.ExecCommand(DECMD_JUSTIFYCENTER, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_ITALIC_onclick() {
	  tbContentElement.ExecCommand(DECMD_ITALIC, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_HYPERLINK_onclick() {
		findOldLinks();
		insertHyperlink();
	}

	function findOldLinks(){
		var selectedText = tbContentElement.DOM.selection;
		if("Text" == selectedText.type)
		{
			var selectedRange = selectedText.createRange();
			var selectedParent = selectedRange.parentElement();

			// see if text already contains a link
			if(selectedParent.tagName == "A")
			{
				document.linkform.elements[0].value = selectedParent.href;
				document.linkform.elements[1].value = selectedParent.target;
				document.linkform.elements[2].value = selectedParent.protocol;
			}
		}
	}


	function insertHyperlink() {
		var args = new Array();
		var arr = null;
		var strURL;
		var strTarget;
		var strProtocol;

		// new or existing link? get link info
		if (linkform.elements[0].value != "null")
			{ args["URL"] = linkform.elements[0].value }
		else
			{ args["URL"] = "http://" }

		// get target info
		if (linkform.elements[1].value != "same")
			{ args["Target"] = linkform.elements[1].value }
		else
			{ args["Target"] = "" }

		// get type info
		if (linkform.elements[2].value != "null")
			{ args["Protocol"] = linkform.elements[2].value }
		else
			{ args["Protocol"] = "http:" }

		// set modal window properties
		strFeatures = "dialogWidth=400px;dialogHeight=150px;scrollbars=no;"
			+ "center=yes;border=thin;help=no;status=no"

		// open modal window for user to input link
		arr = window.showModalDialog("dhtml/inslink.htm", args, strFeatures);

		// process result, if something is returned (if not, user pressed Cancel, do nothing)
	  if (arr != null)
	  {

			// read returned hyperlink properties
			  if (arr["URL"] != null)
			  {
			    strURL = arr["URL"];
			  }
			  if (arr["Target"] == null)
			  {
			  	strTarget ="";
			  }
			  else
			  {
				strTarget = arr["Target"];
			  };
			  if (arr["Protocol"] != null)
			  {
			    strProtocol = arr["Protocol"];
			  };

			makeLink(strURL, strTarget);
	  }

	}

	function makeLink(sURL, sTarget) {
		//Setter inn en ny link
		if (sURL != "")
		{
			if (tbContentElement.QueryStatus(DECMD_HYPERLINK) != DECMDF_DISABLED)
			{
				tbContentElement.ExecCommand(DECMD_HYPERLINK, OLECMDEXECOPT_DONTPROMPTUSER, sURL);
			}
			tbContentElement.focus();
			setTarget(sTarget);
		}
		//Fjern gammel link hvis en finnes
		else{
			var selectedText = tbContentElement.DOM.selection;
			if("Text" == selectedText.type){
				var selectedRange = selectedText.createRange();
				var selectedParent = selectedRange.parentElement();

				var selectedTag = selectedParent.tagName;
				var selectednytt = selectedParent.innerHTML;

				if(selectedTag == "A"){
					//fjerner hele 'a' taggen.
					tbContentElement.DOM.selection.createRange().parentElement().outerHTML = selectednytt
				}
			}
		}
	}

	function setTarget(sTarget) {
		var selectedText = tbContentElement.DOM.selection;
		if ("Text" == selectedText.type)
		{
			tbContentElement.DOM.selection.createRange().parentElement().target = sTarget;
		}
	}

	function DECMD_FINDTEXT_onclick() {
	  tbContentElement.ExecCommand(DECMD_FINDTEXT, OLECMDEXECOPT_PROMPTUSER);
	  tbContentElement.focus();
	}

	function DECMD_DELETE_onclick() {
	  tbContentElement.ExecCommand(DECMD_DELETE, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_CUT_onclick() {
	  tbContentElement.ExecCommand(DECMD_CUT, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_COPY_onclick() {
	  tbContentElement.ExecCommand(DECMD_COPY, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_BOLD_onclick() {
	  tbContentElement.ExecCommand(DECMD_BOLD, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function OnMenuShow(QueryStatusArray) {
	  var i, s;

	  for (i=0; i<QueryStatusArray.length; i++) {
	    s = tbContentElement.QueryStatus(QueryStatusArray[i].command);
	    if (s == DECMDF_DISABLED || s == DECMDF_NOTSUPPORTED) {
	      TBSetState(QueryStatusArray[i].element, "gray");
	    } else if (s == DECMDF_ENABLED  || s == DECMDF_NINCHED) {
	       TBSetState(QueryStatusArray[i].element, "unchecked");
	    } else { // DECMDF_LATCHED
	       TBSetState(QueryStatusArray[i].element, "checked");
	    }
	  }
	  tbContentElement.focus();
	}

	function INTRINSICS_onclick(html) {
	  var selection;

	  selection = tbContentElement.DOM.selection.createRange();
	  selection.pasteHTML(html);
	  tbContentElement.focus();
	}

	function FORMAT_FONT_onclick() {
	  tbContentElement.ExecCommand(DECMD_FONT, OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.focus();
	}

	function DECMD_ABSOLUTEMODE_onclick() {
	  tbContentElement.AbsoluteDropMode = !tbContentElement.AbsoluteDropMode
	  tbContentElement.focus();
	}

	function ParagraphStyle_onchange() {
	  tbContentElement.ExecCommand(DECMD_SETBLOCKFMT, OLECMDEXECOPT_DODEFAULT, ParagraphStyle.value);
	  tbContentElement.focus();
	}

	function PopulateFormatList()
	{
		var oOptions;
		var iNofOptions;
		oOptions = document.all.ParagraphStyle.options;
		//alert(oOptions.length);
		if (oOptions.length==0)  {
			var f = new ActiveXObject("DEGetBlockFmtNamesParam.DEGetBlockFmtNamesParam");
			tbContentElement.ExecCommand(DECMD_GETBLOCKFMTNAMES, OLECMDEXECOPT_DODEFAULT, f);
			var strIgnoreFormats = ", Numbered List, Bulleted List, Directory List, Menu List, Definition Term, Definition, Paragraph,";
			strIgnoreFormats += "Nummerert liste, Punktliste, Katalogliste, Menyliste, Definisjonsterm, Definisjon, Avsnitt,"
			var vbarr = new VBArray(f.Names);
			var arr = vbarr.toArray();

			for (var i=0;i<arr.length;i++) {
				if (strIgnoreFormats.indexOf(','+arr[i]+',')==-1){
					iNofOptions=oOptions.length;
					if (iNofOptions==null){
						iNofOptions = 0;
					}
					oOptions[iNofOptions] = new Option;
					oOptions[iNofOptions].text = arr[i];
					oOptions[iNofOptions].value = arr[i];
				}
			}
		}
	}

	function BuildNumber_onkeypress()
	{
	}

//-->
</script>

<script LANGUAGE="javascript" FOR="tbContentElement" EVENT="DisplayChanged">
<!--
	return tbContentElement_DisplayChanged()
//-->
</script>

<script LANGUAGE="javascript" FOR="tbContentElement" EVENT="ShowContextMenu">
<!--
	return tbContentElement_ShowContextMenu()
//-->
</script>

<script LANGUAGE="javascript" FOR="tbContentElement" EVENT="ContextMenuAction(itemIndex)">
<!--
	return tbContentElement_ContextMenuAction(itemIndex)
//-->
</script>

</head>
<body BGCOLOR="#a3a3a3" TEXT="#000000" LINK="#008000" VLINK="#008000" ALINK="#808000" LANGUAGE="javascript" onload="return window_onload()">

<!-- Hidden field for link -->
<form name="linkform">
	<input type="hidden" name="linkname" value="null">
	<input type="hidden" name="linktarget" value="null">
	<input type="hidden" name="linkprotocol" value="null">
</form>

	<!-- Toolbars -->
	<div class="tbToolbar" ID="StandardToolbar">
		<input type="hidden" class="tbGeneral" ID="BuildNumber" TITLE="Enter build number, hit RETURN" style="width:50" LANGUAGE="javascript" onkeypress="return BuildNumber_onkeypress()">

		<div class="tbButton" ID="MENU_FILE_SAVE" TITLE="Lagre artikkel" LANGUAGE="javascript" onclick="return MENU_FILE_SAVE_onclick()">
			<img class="tbIcon" src="dhtml/images/save.gif" WIDTH="23" HEIGHT="22">
		</div>

		<div class="tbSeparator"></div>

		<div class="tbButton" ID="DECMD_CUT" TITLE="Klipp" LANGUAGE="javascript" onclick="return DECMD_CUT_onclick()">
			<img class="tbIcon" src="dhtml/images/cut.gif" WIDTH="23" HEIGHT="22">
		</div>
		<div class="tbButton" ID="DECMD_COPY" TITLE="Kopier" LANGUAGE="javascript" onclick="return DECMD_COPY_onclick()">
			<img class="tbIcon" src="dhtml/images/copy.gif" WIDTH="23" HEIGHT="22">
		</div>

		<div class="tbButton" ID="DECMD_PASTE" TITLE="Lim inn" LANGUAGE="javascript" onclick="return DECMD_PASTE_onclick()">
			<img class="tbIcon" src="dhtml/images/paste.gif" WIDTH="23" HEIGHT="22">
		</div>

		<div class="tbSeparator"></div>

		<div class="tbButton" ID="DECMD_UNDO" TITLE="Angre" LANGUAGE="javascript" onclick="return DECMD_UNDO_onclick()">
			<img class="tbIcon" src="dhtml/images/undo.gif" WIDTH="23" HEIGHT="22">
		</div>
		<div class="tbButton" ID="DECMD_REDO" TITLE="Reverser angre" LANGUAGE="javascript" onclick="return DECMD_REDO_onclick()">
			<img class="tbIcon" src="dhtml/images/redo.gif" WIDTH="23" HEIGHT="22">
		</div>

		<div class="tbSeparator"></div>

		<div class="tbButton" ID="DECMD_FINDTEXT" TITLE="Finn" LANGUAGE="javascript" onclick="return DECMD_FINDTEXT_onclick()">
			<img class="tbIcon" src="dhtml/images/find.gif" WIDTH="23" HEIGHT="22">
		</div>

		<div class="tbSeparator"></div>

		<div class="tbButton" ID="DECMD_OUTDENT" TITLE="Innrykk" LANGUAGE="javascript" onclick="return DECMD_OUTDENT_onclick()">
			<img class="tbIcon" src="dhtml/images/deindent.gif" WIDTH="23" HEIGHT="22">
		</div>
		<div class="tbButton" ID="DECMD_INDENT" TITLE="Fjern innrykk" LANGUAGE="javascript" onclick="return DECMD_INDENT_onclick()">
			<img class="tbIcon" src="dhtml/images/inindent.GIF" WIDTH="23" HEIGHT="22">
		</div>

		<div class="tbButton" ID="DECMD_JUSTIFYLEFT" TITLE="Venstrejuster" TBTYPE="radio" NAME="Justify" LANGUAGE="javascript" onclick="return DECMD_JUSTIFYLEFT_onclick()">
			<img class="tbIcon" src="dhtml/images/left.gif" WIDTH="23" HEIGHT="22">
		</div>
		<div class="tbButton" ID="DECMD_JUSTIFYCENTER" TITLE="Midtstill" TBTYPE="radio" NAME="Justify" LANGUAGE="javascript" onclick="return DECMD_JUSTIFYCENTER_onclick()">
			<img class="tbIcon" src="dhtml/images/center.gif" WIDTH="23" HEIGHT="22">
		</div>
		<div class="tbButton" ID="DECMD_JUSTIFYRIGHT" TITLE="Høyrejuster" TBTYPE="radio" NAME="Justify" LANGUAGE="javascript" onclick="return DECMD_JUSTIFYRIGHT_onclick()">
			<img class="tbIcon" src="dhtml/images/right.gif" WIDTH="23" HEIGHT="22">
		</div>

		<select ID="ParagraphStyle" class="tbGeneral" style="width:90" TITLE="Tekststiler" LANGUAGE="javascript" onchange="return ParagraphStyle_onchange()">
		</select>

		<div class="tbSeparator"></div>

		<div class="tbButton" ID="DECMD_BOLD" TITLE="Uthev" TBTYPE="toggle" LANGUAGE="javascript" onclick="return DECMD_BOLD_onclick()">
			<img class="tbIcon" src="dhtml/images/bold.gif" WIDTH="23" HEIGHT="22">
		</div>

		<div class="tbButton" ID="DECMD_ITALIC" TITLE="Kursiv" TBTYPE="toggle" LANGUAGE="javascript" onclick="return DECMD_ITALIC_onclick()">
			<img class="tbIcon" src="dhtml/images/italic.gif" WIDTH="23" HEIGHT="22">
		</div>
		<div class="tbButton" ID="DECMD_UNDERLINE" TITLE="Understrek" TBTYPE="toggle" LANGUAGE="javascript" onclick="return DECMD_UNDERLINE_onclick()">
			<img class="tbIcon" src="dhtml/images/under.gif" WIDTH="23" HEIGHT="22">
		</div>

		<div class="tbSeparator"></div>

		<div class="tbButton" ID="DECMD_ORDERLIST" TITLE="Nummerert liste" TBTYPE="toggle" LANGUAGE="javascript" onclick="return DECMD_ORDERLIST_onclick()">
			<img class="tbIcon" src="dhtml/images/numlist.gif" WIDTH="23" HEIGHT="22">
		</div>
		<div class="tbButton" ID="DECMD_UNORDERLIST" TITLE="Unummerert liste" TBTYPE="toggle" LANGUAGE="javascript" onclick="return DECMD_UNORDERLIST_onclick()">
			<img class="tbIcon" src="dhtml/images/bullist.gif" WIDTH="23" HEIGHT="22">
		</div>

		<div class="tbSeparator"></div>

		<div class="tbButton" ID="DECMD_HYPERLINK" TITLE="Peker" LANGUAGE="javascript" onclick="return DECMD_HYPERLINK_onclick()">
			<img class="tbIcon" src="dhtml/images/link.gif" WIDTH="23" HEIGHT="22">
		</div>

	</div>
	<!-- DHTML Editing control Object. This will be the body object for the toolbars. -->
	<OBJECT class="tbContentElement" id="tbContentElement" classid="clsid:2D360201-FFF5-11D1-8D03-00A0C959BC0A" class="editorMain">
	<PARAM NAME="ActivateApplets" VALUE="0">
	<PARAM NAME="width" VALUE="1">
	<PARAM NAME="ActivateActiveXControls" VALUE="0">
	<PARAM NAME="ActivateDTCs" VALUE="-1">
	<PARAM NAME="ShowDetails" VALUE="0">
	<PARAM NAME="ShowBorders" VALUE="0">
	<PARAM NAME="Appearance" VALUE="1">
	<PARAM NAME="Scrollbars" VALUE="1">
	<PARAM NAME="ScrollbarAppearance" VALUE="1">
	<PARAM NAME="SourceCodePreservation" VALUE="1">
	<PARAM NAME="AbsoluteDropMode" VALUE="0">
	<PARAM NAME="SnapToGrid" VALUE="0">
	<PARAM NAME="SnapToGridX" VALUE="0">
	<PARAM NAME="SnapToGridY" VALUE="0">
	<PARAM NAME="UseDivOnCarriageReturn" VALUE="0">
	</OBJECT>
	<!-- DEInsertTableParam Object -->

	<object ID="ObjTableInfo" CLASSID="clsid:47B0DFC7-B7A3-11D1-ADC5-006008A5848C" VIEWASTEXT></object>
	<!-- DEGetBlockFmtNamesParam Object -->
	<object ID="ObjBlockFormatInfo" CLASSID="clsid:8D91090E-B955-11D1-ADC5-006008A5848C" VIEWASTEXT></object>
<!-- Form for posting updated files back to server -->
<form NAME="UploadForm" action="minitor2.asp" method="POST">
	<input type="hidden" name="LinkTitle">
	<input type="hidden" name="Message">
	<input type="hidden" name="BuildNumber" LANGUAGE=javascript onkeypress="return BuildNumber_onkeypress()" onkeypress="return BuildNumber_onkeypress()" onkeypress="return BuildNumber_onkeypress()" onkeypress="return BuildNumber_onkeypress()">
</form>

<!-- Toolbar Code File. Note: This must always be the last thing on the page -->
<script LANGUAGE="Javascript" SRC="dhtml/Toolbars/toolbars.js">
</script>

    </div>
</body>
</html>

