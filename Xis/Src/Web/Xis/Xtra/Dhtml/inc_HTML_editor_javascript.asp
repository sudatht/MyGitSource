<%
'HTML Editor: File 1/3.
'The HTML editor consists of 3 files:
' - inc_HTML_editor_javascript  (all the javascripts necessary, events, initializations etc.)
' - inc_HTML_editor_main		(the editor object, menu bars and buttons)
' - inc_HTML_editor_end			(Menubar initialization, a couple of forms temp forms)
'This include file must be included in the <HEAD> tag of the ASP file, as it contains the
'necessary javascript functions for the editor.
'The editor expects the following Vars to be declared and initialized:
'strTarget - The HTML field in the main ASP file to transfer HTML to/from
'strHTMLProperty - the property of the field that contains the HTML
'lLeftPadding - Number of pixels from the left border to place editor
'lTopPadding - Number of pixels from the top to place editor
'lEditorWidth - Width of the HTML-editor
'lEditorHeight - Height of the HTML-editor
%>

<!-- Script Functions and Event Handlers -->
<script LANGUAGE="JavaScript" SRC="dhtml/dhtmled.js">
</script>



<script LANGUAGE="javascript" FOR="tbContentElement" EVENT="onkeydown">
	Key = tbContentElement.DOM.parentWindow.event.keyCode
	cKey = tbContentElement.DOM.parentWindow.event.ctrlKey

	if (cKey)
	{ switch(Key)
		{
		case 86:
			tbContentElement.ExecCommand(DECMD_PASTE,OLECMDEXECOPT_DODEFAULT);

			var editBody = tbContentElement.DOM.body;
			for (var intLoop = 0; intLoop < editBody.all.length; intLoop++) {
				el = editBody.all[intLoop];
				el.removeAttribute("className","",0);
				el.removeAttribute("style","",0);

			}

			var shtml = tbContentElement.DOM.body.innerHTML;
			shtml = shtml.replace(/<o:p>&nbsp;<\/o:p>/g, ""); // Remove all instances of <o:p>&nbsp;</o:p>
			shtml = shtml.replace(/o:/g, ""); // remove all o: prefixes
			shtml = shtml.replace(/<st1:.*?>/g, ""); // remove all SmartTags (from Word XP!)
			shtml = shtml.replace(/<FONT.*?>/g, ""); // remove all start font tags
			shtml = shtml.replace(/<\/FONT.*?>/g, ""); // remove all end font tags
			tbContentElement.DOM.body.innerHTML = shtml;

			//alert(tbContentElement.DOM.body.innerHTML);
			tbContentElement.DOM.parentWindow.event.cancelBubble=true;
			tbContentElement.DOM.parentWindow.event.returnValue=false;
			break
		}
	}
</script>


<script language="vbscript">
	Sub MENU_FILE_SAVE_onclick()

		If Len(tbContentElement.DOM.body.innerHTML) = 0 Then
		  filteredHTML = ""
		Else
			filteredHTML = tbContentElement.FilterSourceCode (tbContentElement.DOM.body.innerHTML)
		End If

		document.all.<%=strTarget%>.<%=strHTMLProperty%> = filteredHTML
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
	  //QueryStatusEditMenu[6] = new QueryStatusItem(DECMD_PASTE_WORD, document.body.all["EDIT_PASTE_WORD"]);
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

	   if (document.all.<%=strTarget%>.<%=strHTMLProperty%> != "")
	   {
			document.all.tbContentElement.DocumentHTML = document.all.<%=strTarget%>.<%=strHTMLProperty%>;
	   }else{
			document.all.tbContentElement.DocumentHTML = "<p></p>";
	   }
	  //Populate listbox containing valid formats
	  PopulateFormatList();
	  BodyTextTmp.document.designMode = "On";

	}

	function DECMD_WORD_onclick()
	{
		var selectedRange;
		var selectedText = tbContentElement.DOM.selection;

		if(selectedText.type == "Text" || selectedText.type == "None" )
		{
			selectedRange = selectedText.createRange();
			if(selectedText.type == "None")
			{
				selectedRange.collapse();
			}
			document.frames("BodyTextTmp").focus();
			document.frames("BodyTextTmp").document.execCommand("SelectAll");
			document.frames("BodyTextTmp").document.execCommand("Paste");
			selectedRange.pasteHTML(this.cleanFromWord());
			selectedRange.select();
		}
	}


	function cleanUpWordHTML( html )
	{
		html = html.replace(/<\/xml:namespace prefix = o ns = "urn:schemas-microsoft-com:office:office"*>/, "");
		html = html.replace(/<\/o:p*>/, "") ;
		html = html.replace(/<\/?o:p*>/, "") ;
		html = html.replace(/<\\?\?xml[^>]*>/, "");
		html = html.replace(/<\/?SPAN[^>]*>/gi, "" );
		html = html.replace(/<(\w[^>]*) class=([^ |>]*)([^>]*)/gi, "<$1$3") ;
		html = html.replace(/<(\w[^>]*) style="([^"]*)"([^>]*)/gi, "<$1$3") ;
		html = html.replace(/<(\w[^>]*) lang=([^ |>]*)([^>]*)/gi, "<$1$3") ;
		html = html.replace(/<\\?\?xml[^>]*>/gi, "") ;
		html = html.replace(/<\/?\w+:[^>]*>/gi, "") ;
		html = html.replace(/&nbsp;/, " " );
		var re = new RegExp("(<P)([^>]*>.*?)(<\/P>)","gi") ;
		html = html.replace( re, "<div$2</div>" ) ;

		return html;
	}

	function cleanFromWord()
	{
		var BodyText = BodyTextTmp;
		var sHTML = BodyText.document.body.innerHTML;
		return cleanUpWordHTML(sHTML);
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
	  tbContentElement.DOM.focus();
	}

	function DECMD_UNORDERLIST_onclick() {
	  tbContentElement.ExecCommand(DECMD_UNORDERLIST,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
	}

	function DECMD_UNDO_onclick() {
	  tbContentElement.ExecCommand(DECMD_UNDO,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
	}

	function DECMD_UNDERLINE_onclick() {
	  tbContentElement.ExecCommand(DECMD_UNDERLINE,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
	}

	function DECMD_SNAPTOGRID_onclick() {
	  tbContentElement.SnapToGrid = !tbContentElement.SnapToGrid;
	  tbContentElement.DOM.focus();
	}

	function DECMD_SHOWDETAILS_onclick() {
	  tbContentElement.ShowDetails = !tbContentElement.ShowDetails;
	  tbContentElement.DOM.focus();
	}

	function DECMD_SELECTALL_onclick() {
	  tbContentElement.ExecCommand(DECMD_SELECTALL,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
	}

	function DECMD_REDO_onclick() {
	  tbContentElement.ExecCommand(DECMD_REDO,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
	}

	function DECMD_PASTE_onclick() {
	  tbContentElement.ExecCommand(DECMD_PASTE,OLECMDEXECOPT_DODEFAULT);
	  StripHTML();
	  tbContentElement.DOM.focus();
	}

	function StripHTML() {
		var editBody = tbContentElement.DOM.body;
		for (var intLoop = 0; intLoop < editBody.all.length; intLoop++) {
			el = editBody.all[intLoop];
			el.removeAttribute("className","",0);
			el.removeAttribute("style","",0);
		}

		var shtml = tbContentElement.DOM.body.innerHTML;
		shtml = shtml.replace(/<o:p>&nbsp;<\/o:p>/g, ""); // Remove all instances of <o:p>&nbsp;</o:p>
		shtml = shtml.replace(/o:/g, ""); // remove all o: prefixes
		shtml = shtml.replace(/<st1:.*?>/g, ""); // remove all SmartTags (from Word XP!)
		tbContentElement.DOM.body.innerHTML = shtml;
	}

	function DECMD_ORDERLIST_onclick() {
	  tbContentElement.ExecCommand(DECMD_ORDERLIST,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
	}

	function DECMD_MAKE_ABSOLUTE_onclick() {
	  tbContentElement.ExecCommand(DECMD_MAKE_ABSOLUTE,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
	}

	function DECMD_LOCK_ELEMENT_onclick() {
	  tbContentElement.ExecCommand(DECMD_LOCK_ELEMENT,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
	}

	function DECMD_OUTDENT_onclick() {
	  tbContentElement.ExecCommand(DECMD_OUTDENT,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
	}

	function DECMD_INDENT_onclick() {
	  tbContentElement.ExecCommand(DECMD_INDENT,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
	}

	function DECMD_JUSTIFYRIGHT_onclick() {
	  tbContentElement.ExecCommand(DECMD_JUSTIFYRIGHT,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
	}

	function DECMD_JUSTIFYLEFT_onclick() {
	  tbContentElement.ExecCommand(DECMD_JUSTIFYLEFT,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
	}

	function DECMD_JUSTIFYCENTER_onclick() {
	  tbContentElement.ExecCommand(DECMD_JUSTIFYCENTER,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
	}

	function DECMD_ITALIC_onclick() {
	  tbContentElement.ExecCommand(DECMD_ITALIC,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
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
	  tbContentElement.ExecCommand(DECMD_FINDTEXT,OLECMDEXECOPT_PROMPTUSER);
	  tbContentElement.focus();
	}

	function DECMD_DELETE_onclick() {
	  tbContentElement.ExecCommand(DECMD_DELETE,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
	}

	function DECMD_CUT_onclick() {
	  tbContentElement.ExecCommand(DECMD_CUT,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
	}

	function DECMD_COPY_onclick() {
	  tbContentElement.ExecCommand(DECMD_COPY,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
	}

	function DECMD_BOLD_onclick() {
	  tbContentElement.ExecCommand(DECMD_BOLD,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
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
	  tbContentElement.DOM.focus();
	}

	function INTRINSICS_onclick(html) {
	  var selection;

	  selection = tbContentElement.DOM.selection.createRange();
	  selection.pasteHTML(html);
	  tbContentElement.DOM.focus();
	}

	function FORMAT_FONT_onclick() {
	  tbContentElement.ExecCommand(DECMD_FONT,OLECMDEXECOPT_DODEFAULT);
	  tbContentElement.DOM.focus();
	}

	function DECMD_ABSOLUTEMODE_onclick() {
	  tbContentElement.AbsoluteDropMode = !tbContentElement.AbsoluteDropMode
	  tbContentElement.DOM.focus();
	}

	function ParagraphStyle_onchange() {
	  tbContentElement.ExecCommand(DECMD_SETBLOCKFMT, OLECMDEXECOPT_DODEFAULT, ParagraphStyle.value);
	  tbContentElement.DOM.focus();
	}

	function PopulateFormatList()
	{
		var oOptions;
		var iNofOptions;
		oOptions = document.all.ParagraphStyle.options;
		if (oOptions.length==0)  {
			var f = new ActiveXObject("DEGetBlockFmtNamesParam.DEGetBlockFmtNamesParam");
			tbContentElement.ExecCommand(DECMD_GETBLOCKFMTNAMES,OLECMDEXECOPT_DODEFAULT,f);
			var strIgnoreFormats = ",Numbered List,Bulleted List,Directory List,Menu List,Definition Term,Definition,Paragraph,";
			strIgnoreFormats += "Nummerert liste,Punktliste,Katalogliste,Menyliste,Definisjonsterm,Definisjon,Avsnitt,"
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
