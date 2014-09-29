<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<fieldset>
	<legend><%= GetString("Attributes") %></legend>
	<div align="left" style="padding-left:12px">
		<table class="normal">
			<tr>		
				<td><%= GetString("CssClass") %>:</td>
				<td><input type="text" id="inp_class" style="width:100px" /></td>	
			</tr>
			<tr>
			<td><%= GetString("Width") %> :</td>
				<td><input type="text" id="inp_width" style="width:100px" onkeypress="return CancelEventIfNotDigit()" /></td>
				
			</tr>
			<tr>
				<td><%= GetString("Height") %> :</td>
				<td><input type="text" id="inp_height" style="width:100px" onkeypress="return CancelEventIfNotDigit()" /></td>
			</tr>
			<tr>		
				<td><%= GetString("Alignment") %>:</td>
				<td><select id="sel_align" style="width:100px">
						<option value=""><%= GetString("NotSet") %></option>
						<option value="left"><%= GetString("Left") %></option>
						<option value="center"><%= GetString("Center") %></option>
						<option value="right"><%= GetString("Right") %></option>
					</select>
				</td>
			</tr>
			<tr>		
				<td><%= GetString("Text-Align") %>:</td>
				<td><select id="sel_textalign" style="width:100px">
						<option value=""><%= GetString("NotSet") %></option>
						<option value="left"><%= GetString("Left") %></option>
						<option value="center"><%= GetString("Center") %></option>
						<option value="right"><%= GetString("Right") %></option>
						<option value="justify"><%= GetString("Justify") %></option>
					</select>
				</td>
			</tr>
			<tr>		
				<td><%= GetString("Float") %>:</td>
				<td><select id="sel_float" style="width:100px">
						<option value=""><%= GetString("NotSet") %></option>
						<option value="left"><%= GetString("Left") %></option>
						<option value="right"><%= GetString("Right") %></option>
					</select>
				</td>
			</tr>
			<tr>
				<td><%= GetString("Color") %></td>
				<td>
<input autocomplete="off" type="text" id="inp_forecolor" name="inp_forecolor" size="7" style="WIDTH:57px;" />
<img alt="" src="../images/colorpicker.gif" id="img_forecolor" style="vertical-align:top;" />
				</td>
			</tr>
			<tr>
				<td><%= GetString("BackColor") %></td>
				<td>
<input autocomplete="off" type="text" id="inp_backcolor" name="inp_forecolor" size="7" style="WIDTH:57px;" />
<img alt="" src="../images/colorpicker.gif" id="img_backcolor" style="vertical-align:top;" />
				</td>
			</tr>
			<tr>		
				<td style='width:100px'><%= GetString("Title") %>:</td>
				<td>
					<textarea id="inp_tooltip" rows="5" cols="20" style="width:200px"></textarea>
				</td>
			</tr>
		</table>
	</div>
</fieldset>
<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_Common.js"></script>