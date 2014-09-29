<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<style type="text/css">
	.codebutton
	{
		width:110px; 
	}
</style>
<fieldset><legend><%= GetString("Input") %></legend>
	<table class="normal">
		<tr>
			<td style="width:60"><%= GetString("Name") %>:</td>
			<td><input type="text" id="inp_name" style="width:100px" /></td>
			<td>&nbsp;&nbsp;&nbsp;</td>
			<td><%= GetString("AccessKey") %>:</td>
			<td>
				<input type="text" id="inp_access" size="1" maxlength="1" />
			</td>
		</tr>
		<tr>
			<td><%= GetString("ID") %>:</td>
			<td><input type="text" id="inp_id" style="width:100px" /></td>
			<td>&nbsp;&nbsp;</td>
			<td>
				<%= GetString("TabIndex") %>:
			</td>
			<td>
				<input type="text" id="inp_index" size="5" value="" maxlength="5" onkeypress="return CancelEventIfNotDigit()" />&nbsp;
			</td>
		</tr>
		<tr>
			<td><%= GetString("Size") %>:</td>
			<td colspan="4"><input type="text" id="inp_size" style="width:100px" /></td>
		</tr>
		<tr>
			<td>
			</td>
			<td colspan="4"><input type="checkbox" id="inp_Multiple" /><label for="inp_Multiple"><%= GetString("AllowMultipleSelections") %></label>
			</td>
		</tr>
		<tr>
			<td>
			</td>
			<td colspan="4"><input type="checkbox" id="inp_Disabled" /><label for="inp_Disabled"><%= GetString("Disabled") %></label>
			</td>
		</tr>
	</table>
	<%= GetString("Items") %>:
	<br />
	<table class="normal">
		<tr>
			<td><%= GetString("Text") %>:
				<br />
				<input type="text" id="inp_item_text" style="width:130px" />
			</td>
			<td><%= GetString("Value") %>:
				<br />
				<input type="text" id="inp_item_value" style="width:130px" />
			</td>
			<td rowspan="3" valign="top">
				<table>
					<tr>
						<td colspan="2"><button class="codebutton" onclick="Insert();" id="btnInsertItem"><%= GetString("Insert") %></button>
						</td>
					</tr>
					<tr>
						<td colspan="2"><button class="codebutton" onclick="Update();" id="btnUpdateItem"><%= GetString("Update") %></button>
						</td>
					</tr>
					<tr>
						<td colspan="2"><button class="codebutton" onclick="Delete();" id="btnDeleteItem"><%= GetString("Delete") %></button>
						</td>
					</tr>
					<tr>
						<td colspan="2"><button class="codebutton" onclick="Move(1);" id="btnMoveUpItem"><%= GetString("MoveUp") %></button>
						</td>
					</tr>
					<tr>
						<td colspan="2"><button class="codebutton" onclick="Move(-1);" id="btnMoveDownItem"><%= GetString("MoveDown") %></button>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td><select size="6" id="list_options" style="width:130px" onchange="document.getElementById('list_options2').selectedIndex = this.selectedIndex;Select(this);FireUIChanged();"></select></td>
			<td><select size="6" id="list_options2" style="width:130px" onchange="document.getElementById('list_options').selectedIndex = this.selectedIndex;Select(this);FireUIChanged();"></select></td>
		</tr>
		<tr>
			<td><%= GetString("Color") %>:&nbsp;<input autocomplete="off" size="7" type="text" id="inp_item_forecolor" />
			<img alt="" id="inp_item_forecolor_Preview" src="../Images/colorpicker.gif" style="vertical-align:top"/>
			</td>
			<td><%= GetString("BackColor") %>:&nbsp;<input autocomplete="off" size="7" type="text" id="inp_item_backcolor" />
			<img alt="" id="inp_item_backcolor_Preview" src="../Images/colorpicker.gif" style="vertical-align:top"/>
			</td>
		</tr>
	</table>
</fieldset>
<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_Select.js"></script>