<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<fieldset><legend><%= GetString("Input") %></legend>
	<table class="normal">
		<tr>
			<td><%= GetString("Type") %>:</td>
			<td colspan="3"><input type="text" id="inp_type" readonly="readonly" disabled="disabled" style="width:100px" /></td>
		</tr>
		<tr>
			<td style="width:60"><%= GetString("Name") %>:</td>
			<td colspan="3"><input type="text" id="inp_name" style="width:100px" /></td>
		</tr>
		<tr>
			<td><%= GetString("Value") %>:</td>
			<td colspan="3"><input type="text" id="inp_value" style="width:250px" /></td>
		</tr>
		<tr id="row_txt1">
			<td><%= GetString("Size") %>:</td>
			<td colspan="3"><input type="text" id="inp_Size" style="width:100px" onkeypress="return CancelEventIfNotDigit()" /></td>
		</tr>
		<tr id="row_txt2">
			<td><%= GetString("MaxLength") %>:</td>
			<td colspan="3"><input type="text" id="inp_MaxLength" style="width:100px" maxlength="9" onkeypress="return CancelEventIfNotDigit()" /></td>
		</tr>
		<tr id="row_img">
			<td><%= GetString("Src") %>:</td>
			<td colspan="3">
			    <input type="text" id="inp_src" style="width:250px" />&nbsp; 
			    <input id="btnbrowse" value="<%= GetString("Browse") %>" type="button" />
			</td>
		</tr>
		<tr id="row_img2">
			<td><%= GetString("Alignment") %>:</td>
			<td>
				<select name="inp_Align" style="WIDTH : 80px" id="sel_Align">
					<option id="optNotSet" value=""><%= GetString("NotSet") %></option>
					<option id="optLeft" value="left"><%= GetString("Left") %></option>
					<option id="optRight" value="right"><%= GetString("Right") %></option>
					<option id="optTexttop" value="textTop"><%= GetString("Texttop") %></option>
					<option id="optAbsMiddle" value="absMiddle"><%= GetString("Absmiddle") %></option>
					<option id="optBaseline" value="baseline" selected="selected"><%= GetString("Baseline") %></option>
					<option id="optAbsBottom" value="absBottom"><%= GetString("Absbottom") %></option>
					<option id="optBottom" value="bottom"><%= GetString("Bottom") %></option>
					<option id="optMiddle" value="middle"><%= GetString("Middle") %></option>
					<option id="optTop" value="top"><%= GetString("Top") %></option>
				</select>
			</td>
			<td><%= GetString("Bordersize") %>:</td>
			<td>
				<input type="text" size="2" name="inp_Border" onkeypress="return CancelEventIfNotDigit()"
					style="WIDTH : 80px" id="inp_Border" />
			</td>
		</tr>
		<tr id="row_img3">
			<td><%= GetString("Width") %>:</td>
			<td>
				<input type="text" onkeypress="return CancelEventIfNotDigit()" style="width:80px" size="2"
					id="inp_width" />
			</td>
			<td><%= GetString("Height") %>:</td>
			<td>
				<input type="text" onkeypress="return CancelEventIfNotDigit()" style="width:80px" size="2"
					id="inp_height" />
			</td>
		</tr>
		<tr id="row_img4">
			<td><%= GetString("Horizontal") %>:</td>
			<td>
				<input type="text" onkeypress="return CancelEventIfNotDigit()" style="width:80px" size="2"
					id="inp_HSpace" />
			</td>
			<td><%= GetString("Vertical") %>:</td>
			<td>
				<input type="text" onkeypress="return CancelEventIfNotDigit()" style="width:80px" size="2"
					id="inp_VSpace" />
			</td>
		</tr>
		<tr id="row_img5">
			<td valign="middle" style="white-space:nowrap" ><%= GetString("Alternate") %>:</td>
			<td colspan="3"><input type="text" id="AlternateText" size="24" name="AlternateText" style="width:250px" /></td>
		</tr>
		<tr>
			<td><%= GetString("ID") %>:</td>
			<td colspan="3"><input type="text" id="inp_id" style="width:100px" /></td>
		</tr>
		<tr id="row_txt3">
			<td><%= GetString("AccessKey") %>:</td>
			<td colspan="3">
				<input type="text" id="inp_access" size="1" maxlength="1" />
			</td>
		</tr>
		<tr id="row_txt4">
			<td>
				<%= GetString("TabIndex") %>:
			</td>
			<td colspan="3">
				<input type="text" id="inp_index" size="5" value="" maxlength="5" onkeypress="return CancelEventIfNotDigit()" />&nbsp;
			</td>
		</tr>
		<tr id="row_chk">
			<td></td>
			<td><input type="checkbox" id="inp_checked" /><label for="inp_checked"><%= GetString("Checked") %></label></td>
		</tr>
		<tr id="row_txt5">
			<td>
			</td>
			<td colspan="3"><input type="checkbox" id="inp_Disabled" /><label for="inp_Disabled"><%= GetString("Disabled") %></label>
			</td>
		</tr>
		<tr id="row_txt6">
			<td>
			</td>
			<td colspan="3"><input type="checkbox" id="inp_Readonly" /><label for="inp_Readonly"><%= GetString("Readonly") %></label>
			</td>
		</tr>
	</table>
</fieldset>
<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_Input.js"></script>