<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<fieldset>
    <legend><%= GetString("Table") %></legend>
	<table class="normal">
		<tr>
			<td><%= GetString("CellSpacing") %>:</td>
			<td><input type="text" id="inp_cellspacing"  size="14" onkeypress="return CancelEventIfNotDigit()" /></td>
			<td><%= GetString("CellPadding") %>:</td>
			<td><input type="text" id="inp_cellpadding"  size="14" onkeypress="return CancelEventIfNotDigit()" /></td>
		</tr>
		<tr>
			<td><%= GetString("ID") %>:</td>
			<td><input type="text" id="inp_id" size="14" />&nbsp;&nbsp;</td>
			<td><%= GetString("Border") %>:</td>
			<td><input type="text" id="inp_border"  size="14" onkeypress="return CancelEventIfNotDigit()" /></td>
		</tr>
		<tr>
			<td><%= GetString("Backgroundcolor") %>:
			</td>
			<td><input autocomplete="off" type="text" id="inp_bgcolor"  size="14"/>
			</td>
			<td><%= GetString("BorderColor") %>:
			</td>
			<td><input autocomplete="off" type="text" id="inp_bordercolor" size="14"/>
			</td>
		</tr>
		<tr>
			<td valign="middle" style="white-space:nowrap" ><%= GetString("Rules") %>:</td>
			<td><select id="sel_rules">
					<option value=""><%= GetString("NotSet") %></option>
					<option value="all">all</option>
					<option value="rows">rows</option>
					<option value="cols">cols</option>
					<option value="none">none</option>
				</select>
			</td>
			<td colspan="2">
				<input type="checkbox" id="inp_collapse" />
				<label for="inp_collapse"><%= GetString("CollapseBorder") %></label>&nbsp;				
			</td>
		</tr>
	</table>
	<table class="normal">
		<tr>
			<td style='width:60px'><%= GetString("Summary") %> :</td>
			<td>
				<textarea id="inp_summary" rows="3" cols="20" style="width:320px"></textarea>
			</td>
		</tr>
	</table>
	<table class="normal" id="CaptionTable">
		<tr>
			<td style='width:60px'><%= GetString("Caption") %> :</td>
			<td>
			     <button id="btn_editcaption"><%= GetString("Insert") %></button>
			     <button id="btn_delcaption"><%= GetString("Delete") %></button>
			</td>
			<td>&nbsp;</td>
			<td><%= GetString("THEAD") %>:</td>
			<td>
			     <button id="btn_insthead"><%= GetString("Insert") %></button>
			</td>
			
			<td>&nbsp;</td>
			<td><%= GetString("TFOOT") %>:</td>
			<td>
			     <button id="btn_instfoot"><%= GetString("Insert") %></button>		
			</td>
			<td style="width:5"></td>
			<td><img src="../Images/Accessibility.gif" title="Accessibility" /></td>
		</tr>
	</table>
</fieldset>
<fieldset><legend><%= GetString("Common") %></legend>
	<table class="normal">
		<tr>
			<td style='width:60px'><%= GetString("CssClass") %>:</td>
			<td><input type="text" id="inp_class" style="width:80px" /></td>
			<td><%= GetString("Width") %>:</td>
			<td style="white-space:nowrap">
				<input type="text" id="inp_width" style="width:42px" />
				<select id="sel_width_unit">
					<option value="px">px</option>
					<option value="%">%</option>
				</select>
			</td>
			<td><%= GetString("Height") %>:</td>
			<td style="white-space:nowrap">
				<input type="text" id="inp_height" style="width:42px" />
				<select id="sel_height_unit">
					<option value="px">px</option>
					<option value="%">%</option>
				</select>
			</td>
		</tr>
	</table>
	<table class="normal">
		<tr>
			<td style='width:60px'><%= GetString("Alignment") %>:</td>
			<td><select id="sel_align">
					<option value=""><%= GetString("NotSet") %></option>
					<option value="left"><%= GetString("Left") %></option>
					<option value="center"><%= GetString("Center") %></option>
					<option value="right"><%= GetString("Right") %></option>
				</select></td>
			<td>
				<%= GetString("Text-Align") %> :</td>
			<td><select id="sel_textalign">
					<option value=""><%= GetString("NotSet") %></option>
					<option value="left"><%= GetString("Left") %></option>
					<option value="center"><%= GetString("Center") %></option>
					<option value="right"><%= GetString("Right") %></option>
					<option value="justify"><%= GetString("Justify") %></option>
				</select></td>
			<td><%= GetString("Float") %>:
			</td>
			<td><select id="sel_float">
					<option value=""><%= GetString("NotSet") %></option>
					<option value="left"><%= GetString("Left") %></option>
					<option value="right"><%= GetString("Right") %></option>
				</select></td>
		</tr>
	</table>
	<table class="normal">
		<tr>
			<td style='width:60px'><%= GetString("Title") %> :</td>
			<td>
				<textarea id="inp_tooltip" rows="3" cols="20" style="width:320px"></textarea>
			</td>
		</tr>
	</table>
</fieldset>
<script type="text/javascript" >
	    var Caption = "<%= GetString("Caption")%>";
	    var Delete = "<%= GetString("Delete")%>";
	    var Insert = "<%= GetString("Insert")%>";
	    var Edit = "<%= GetString("Edit")%>";
	    var ValidID = "<%= GetString("ValidID")%>";
</script>
<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_Table.js"></script>
