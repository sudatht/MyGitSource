<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<fieldset><legend><%= GetString("Alignment") %></legend>
	<table border="0" cellspacing="0" cellpadding="2" class="normal">
		<tr>
			<td><%= GetString("Horizontal") %>:</td>
			<td>
				<select id="sel_align">
					<option value=""><%= GetString("NotSet") %></option>
					<option value="left"><%= GetString("Left") %></option>
					<option value="center"><%= GetString("Center") %></option>
					<option value="right"><%= GetString("Right") %></option>
					<option value="justify"><%= GetString("Justify") %></option>
				</select>
			</td>
			<td style="white-space:nowrap;width:10" ></td>
			<td><%= GetString("Vertical") %>:</td>
			<td>
				<select id="sel_valign" style="width:90">
					<option value=""><%= GetString("NotSet") %></option>
					<option value="sub"><%= GetString("Subscript") %></option>
					<option value="super"><%= GetString("Superscript") %></option>
					<option value="baseline"><%= GetString("Normal") %></option>
				</select>
			</td>
		</tr>
		<tr>
			<td><%= GetString("Justification") %>:</td>
			<td colspan="4">
				<select id="sel_justify">
					<option value=""><%= GetString("NotSet") %></option>
					<option value="auto">Auto</option>
					<option value="newspaper">newspaper</option>
					<option value="distribute">distribute</option>
					<option value="distribute-all-lines">distribute-all-lines</option>
					<option value="inter-word">inter-word</option>
				</select>
			</td>
		</tr>
	</table>
</fieldset>
<fieldset>
	<legend>
		<%= GetString("Spacing") %></legend>
	<table border="0" cellpadding="2" cellspacing="0" class="normal">
		<tr>
			<td><%= GetString("Letters") %></td>
			<td><select style="width:80px" id="sel_letter">
					<option value=""><%= GetString("NotSet") %></option>
					<option value="normal"><%= GetString("Normal") %></option>
				</select>
				<%= GetString("OR") %> <input type="text" id="tb_letter" style="width:60px" />
				<select id="sel_letter_unit">
					<option value="px">px</option>
					<option value="pt">pt</option>
					<option value="pc">pc</option>
					<option value="em">em</option>
					<option value="cm">cm</option>
					<option value="mm">mm</option>
					<option value="in">in</option>
				</select>
			</td>
		</tr>
		<tr>
			<td><%= GetString("Height") %></td>
			<td><select style="width:80px" id="sel_line">
					<option value=""><%= GetString("NotSet") %></option>
					<option value="normal"><%= GetString("Normal") %></option>
				</select>
				<%= GetString("OR") %> <input type="text" id="tb_line" style="width:60px" />
				<select id="sel_line_unit">
					<option value="px">px</option>
					<option value="%">%</option>
					<option value="pt">pt</option>
					<option value="pc">pc</option>
					<option value="em">em</option>
					<option value="cm">cm</option>
					<option value="mm">mm</option>
					<option value="in">in</option>
				</select>
			</td>
		</tr>
	</table>
</fieldset>
<fieldset><legend><%= GetString("TextFlow") %></legend>
	<table border="0" cellspacing="0" cellpadding="2" class="normal">
		<tr>
			<td style="width:80"><%= GetString("Indentation") %>:
			</td>
			<td><input type="text" id="tb_indent" style="width:60px" />
				<select id="sel_indent_unit">
					<option value="px">px</option>
					<option value="pt">pt</option>
					<option value="pc">pc</option>
					<option value="em">em</option>
					<option value="cm">cm</option>
					<option value="mm">mm</option>
					<option value="in">in</option>
				</select>
			</td>
		</tr>
		<tr>
			<td><%= GetString("TextDirection") %>:</td>
			<td><select id="sel_direction">
					<option value=""><%= GetString("NotSet") %></option>
					<option value="ltr"><%= GetString("LTR") %></option>
					<option value="rtl"><%= GetString("RTL") %></option>
				</select>
			</td>
		</tr>
		<tr>
			<td><%= GetString("WritingMode") %>:</td>
			<td>
				<select id="sel_writingmode">
					<option value=""><%= GetString("NotSet") %></option>
					<option value="lr-tb"><%= GetString("lr-tb") %></option>
					<option value="tb-rl"><%= GetString("tb-rl") %></option>
				</select>
			</td>
		</tr>
	</table>
</fieldset>


<div id="outer" style="height:100px; margin-bottom:10px; padding:5px;"><div id="div_demo"><%= GetString("DemoText") %></div></div>
<br />

<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_Style_Text.js"></script>

