
<!-- Hidden field for link -->
	<form name="linkform">
		<input type="hidden" name="linkname" value="null">
		<input type="hidden" name="linktarget" value="null">
		<input type="hidden" name="linkprotocol" value="null">
	</form>

	<!-- Toolbars -->
	<div class="tbToolbar" ID="StandardToolbar">
		<input type="hidden" class="tbGeneral" ID="BuildNumber" TITLE="Enter build number, hit RETURN" style="width:50" LANGUAGE="javascript" onkeypress="return BuildNumber_onkeypress()">

		<div class="tbButton" ID="DECMD_CUT" TITLE="Klipp" LANGUAGE="javascript" onclick="return DECMD_CUT_onclick()">
			<img class="tbIcon" src="dhtml/images/cut.gif" WIDTH="23" HEIGHT="22">
		</div>
		<div class="tbButton" ID="DECMD_COPY" TITLE="Kopier" LANGUAGE="javascript" onclick="return DECMD_COPY_onclick()">
			<img class="tbIcon" src="dhtml/images/copy.gif" WIDTH="23" HEIGHT="22">
		</div>

		<div class="tbButton" ID="DECMD_PASTE" TITLE="Lim inn" LANGUAGE="javascript" onclick="return DECMD_PASTE_onclick()">
			<img class="tbIcon" src="dhtml/images/paste.gif" WIDTH="23" HEIGHT="22">
		</div>
		<div class="tbButton" ID="DECMD_WORD" TITLE="Lim inn fra word" LANGUAGE="javascript" onclick="return DECMD_WORD_onclick()">
			<img class="tbIcon" src="dhtml/images/pfw.gif" WIDTH="23" HEIGHT="22">
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

		<select ID="ParagraphStyle" class="tbGeneral" style="width:90px" TITLE="Tekststiler" LANGUAGE="javascript" onChange="return ParagraphStyle_onchange()">
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
	<IFRAME id="BodyTextTmp" name="BodyTextTmp" width="1" height="1"></IFRAME>

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

