<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<fieldset>
    <legend>Textarea</legend>
	<table class="normal">
		<tr>
			<td style="width:60"><%= GetString("Name") %>:</td>
			<td><input type="text" id="inp_name" style="width:100px" /></td>
		</tr>
		<tr>
			<td><%= GetString("Columns") %>:</td>
			<td><input type="text" id="inp_cols" style="width:100px" onkeypress="return CancelEventIfNotDigit()" /></td>
		</tr>
		<tr>
			<td><%= GetString("Rows") %>:</td>
			<td><input type="text" id="inp_rows" style="width:100px" onkeypress="return CancelEventIfNotDigit()" /></td>
		</tr>
		<tr>
			<td><%= GetString("Value") %>:</td>
			<td><textarea rows="2" cols="30" style="width:258px" id="inp_value"></textarea></td>
		</tr>
		<tr>
			<td><%= GetString("Wrap") %>:</td>
			<td>
				<select id="sel_Wrap" name="sel_Wrap">
					<option value="">Default</option>
					<option value="off">off</option>
					<option value="virtual">virtual</option>
					<option value="physical">physical</option>
				</select>
			</td>		
		</tr>
		<tr>
			<td style="width:60"><%= GetString("ID") %>:</td>
			<td><input type="text" id="inp_id" style="width:100px" /></td>
		</tr>
		<tr>
			
			<td><%= GetString("AccessKey") %>:</td>
			<td>
				<input type="text" id="inp_access" size="1" maxlength="1" />
			</td>
		</tr>
		<tr>
			<td><%= GetString("TabIndex") %>:</td>
			<td><input type="text" id="inp_index" size="5" value="" maxlength="5" onkeypress="return CancelEventIfNotDigit()" />		
			</td>
		</tr>
		<tr>
			<td>
			</td>
			<td><input type="checkbox" id="inp_Disabled" /><label for="inp_Disabled"><%= GetString("Disabled") %></label>
			</td>
		</tr>
		<tr>
			<td>
			</td>
			<td><input type="checkbox" id="inp_Readonly" /><label for="inp_Readonly"><%= GetString("Readonly") %></label>
			</td>
		</tr>
	</table>
</fieldset>
<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_Textarea.js"></script>
