<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<fieldset>
    <legend><%= GetString("SystemFont") %></legend>
	<select id="sel_font" style="width:240">
		<option value=""><%= GetString("NotSet") %></option>
		<option value="caption"><%= GetString("Caption") %></option>
		<option value="icon"><%= GetString("Icon") %></option>
		<option value="menu"><%= GetString("Menu") %></option>
		<option value="messagebox"><%= GetString("messagebox") %></option>
		<option value="smallcaption"><%= GetString("smallcaption") %></option>
		<option value="statusbar"><%= GetString("statusbar") %></option>
	</select>
</fieldset>
<div id="div_font_detail">
	<fieldset><legend><%= GetString("FontFamily") %></legend>		
		<select id="sel_fontfamily" style="width:240" NAME="sel_fontfamily">
			<option value=""><%= GetString("NotSet") %></option>
			<option value="Arial">Arial</option>
			<option value="Verdana">Verdana</option>
			<option value="Comic Sans MS">Comic Sans MS</option>
			<option value="Courier">Courier</option>
			<option value="Georgia">Georgia</option>
			<option value="Tahoma">Tahoma</option>
			<option value="Times New Roman">Times New Roman</option>
			<option value="Wingdings">Wingdings</option>
		</select>
	</fieldset>
	<fieldset><legend><%= GetString("Decoration") %></legend>
		<input type="checkbox" id="cb_decoration_under" /><label for="cb_decoration_under"><%= GetString("Underline") %></label>
		<input type="checkbox" id="cb_decoration_over" /><label for="cb_decoration_over"><%= GetString("Overline") %></label>
		<input type="checkbox" id="cb_decoration_through" /><label for="cb_decoration_through"><%= GetString("Strikethrough") %></label>
	</fieldset>
	<fieldset><legend><%= GetString("Style") %></legend>
		<input type="checkbox" id="cb_style_bold" /><label for="cb_style_bold"><%= GetString("Bold") %></label>
		<input type="checkbox" id="cb_style_italic" /><label for="cb_style_italic"><%= GetString("Italic") %></label>
		&nbsp;&nbsp;<%= GetString("Capitalization") %>:
		<select id="sel_fontTransform">
			<option value=""><%= GetString("NotSet") %></option>
			<option value="uppercase"><%= GetString("uppercase") %></option>
			<option value="lowercase"><%= GetString("lowercase") %></option>
			<option value="capitalize"><%= GetString("InitialCap") %></option>
		</select>
	</fieldset>
	<fieldset><legend><%= GetString("Size") %></legend>
		<select id="sel_fontsize" style="width:80px">
			<option value=""><%= GetString("NotSet") %></option>
			<option value="xx-large">xx-large</option>
			<option value=" x-large">x-large</option>
			<option value="large">large</option>
			<option value="medium">medium</option>
			<option value="small">small</option>
			<option value="x-small">x-small</option>
			<option value="xx-small">xx-small</option>
			<option value="larger">larger</option>
			<option value="smaller">Smaller</option>
		</select>
		<%= GetString("OR") %> <input type="text" id="inp_fontsize" style="width:70px" />
		<select id="sel_fontsize_unit">
			<option value="px">px</option>
			<option value="pt">pt</option>
			<option value="pc">pc</option>
			<option value="em">em</option>
			<option value="cm">cm</option>
			<option value="mm">mm</option>
			<option value="in">in</option>
		</select>
	</fieldset>
	<fieldset><legend><%= GetString("Color") %></legend>	
				<input autocomplete="off" size="7" type="text" id="inp_color" style="width:57px"/>
				<img alt="" id="inp_color_Preview" src="../Images/colorpicker.gif" style="vertical-align:top" />			
	</fieldset>
</div>

<div id="outer" style="height:100px; margin-bottom:10px; padding:5px;"><div id="div_demo"><%= GetString("DemoText") %></div></div><br />
<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_Style_Font.js"></script>