<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<fieldset><legend><%= GetString("Hyperlink_Information") %></legend>
	<table class="normal">
		<tr>
			<td style="width:60px"><%= GetString("Url") %>:</td>
			<td colspan="3"><input type="text" id="inp_src" style="width:220px" />&nbsp;&nbsp;
			<button id="btnbrowse" class="formbutton"><%= GetString("Browse") %></button>
			</td>
		</tr>
		<tr>
			<td style="width:60px"><%= GetString("Type") %>:</td>
			<td>
				<select id="sel_protocol" onchange="sel_protocol_change()">
					<option value="http://">http://</option>
					<option value="https://">https://</option>
					<option value="ftp://">ftp://</option>
					<option value="news://">news://</option>
					<option value="mailto:">mailto:</option>
					<!-- last one : if move this to front , change the script too -->
					<option value="others"><%= GetString("Other") %></option>
				</select>
			</td>
			<td>
				<%= GetString("Target") %>: 
			</td>
			<td>
				<select id="inp_target" name="inp_target">
					<option value="" selected="selected"><%= GetString("NotSet") %></option>
					<option value="_blank"><%= GetString("Newwindow") %></option>
					<option value="_self"><%= GetString("Samewindow") %></option>
					<option value="_top"><%= GetString("Topmostwindow") %></option>
					<option value="_parent"><%= GetString("Parentwindow") %></option>
				</select>
			</td>
		</tr>
		<tr>
			<td valign="middle" style="white-space:nowrap" ><%= GetString("ID") %>:</td>
			<td>
				<input type="text" id="inp_id" size="14" NAME="inp_id"/>
			</td>
			<td>Nofollow:</td>
			<td>
				<input type="checkbox" name="Nofollow" id="Nofollow" value="Nofollow" />
			</td>
		</tr>
		<tr>
			<td valign="middle" style="white-space:nowrap" ><%= GetString("CssClass") %>:</td>
			<td>
				<input type="text" id="inp_class" size="14" NAME="inp_class"/>
			</td>
			<td></td>
			<td>
			</td>
		</tr>
		<tr>
			<td valign="middle" style="white-space:nowrap" ><%= GetString("AccessKey") %>:</td>
			<td colspan="3">
				<input type="text" id="inp_access" size="2" maxlength="1" />
				&nbsp;
				<%= GetString("TabIndex") %>:&nbsp;
				<input type="text" id="inp_index" size="5" maxlength="5" onkeypress="return CancelEventIfNotDigit()" />
				&nbsp;
				<%= GetString("Color") %>:&nbsp;
				<input autocomplete="off" type="text" id="inp_color" size="7" style="width:57px" />
				<img alt="" id="inp_color_Preview" src="../Images/colorpicker.gif" style="vertical-align:top"/>
			</td>
		</tr>
		<tr>
			<td><%= GetString("Title") %>:</td>
			<td colspan="3">
				<textarea id="inp_title" rows="2" cols="20" style="width:260px"></textarea>
			</td>
		</tr>
		<tr>
			<td align="right"></td>
			<td colspan="3">
				<input type="checkbox" id="inp_checked" onclick="ToggleAnchors();" /><%= GetString("SelectAnchor") %>
				<br />
				<select size="5" name="allanchors" style="width: 255" id="allanchors" onchange="selectAnchor(this.value);"></select>
			</td>
		</tr>
	</table>									
</fieldset>
<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_A.js"></script>