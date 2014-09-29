<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<fieldset><legend><%= GetString("Layout") %></legend>
	<table border="0" cellspacing="0" cellpadding="2" class="normal">
		<tr>
			<td style="width:50"><%= GetString("Position") %>:
			</td>
			<td><select style="width:80" id="sel_position">
					<option value=""><%= GetString("NotSet") %></option>
					<option value="absolute"><%= GetString("Absolute") %></option>
					<option value="relative"><%= GetString("Relative") %></option>
				</select></td>
			<td style="width:50"><%= GetString("Display") %>:
			</td>
			<td><select style="width:80" id="sel_display">
					<option value=""><%= GetString("NotSet") %></option>
					<option value="block">block</option>
					<option value="inline">inline</option>
					<option value="inline-block">inline-block</option>
				</select></td>
		</tr>
		<tr>
			<td style="width:50"><%= GetString("Float") %>:
			</td>
			<td><select style="width:80" id="sel_float">
					<option value=""><%= GetString("NotSet") %></option>
					<option value="left"><%= GetString("FloatLeft") %></option>
					<option value="right"><%= GetString("FloatRight") %></option>
					<option value="none"><%= GetString("FloatNone") %></option>
				</select></td>
			<td style="width:50"><%= GetString("Clear") %>:
			</td>
			<td><select style="width:80" id="sel_clear">
					<option value=""><%= GetString("NotSet") %></option>
					<option value="left"><%= GetString("ClearLeft") %></option>
					<option value="right"><%= GetString("ClearRight") %></option>
					<option value="both"><%= GetString("ClearBoth") %></option>
					<option value="none"><%= GetString("ClearNone") %></option>
				</select></td>
		</tr>
	</table>
</fieldset>
<fieldset><legend><%= GetString("Size") %></legend>
	<table border="0" cellspacing="0" cellpadding="2" class="normal">
		<tr>
			<td style="width:50"><%= GetString("Top") %></td>
			<td><input type="text" id="tb_top" style="width:50px" />
				<select id="sel_top_unit">
					<option value="px">px</option>
					<option value="pt">pt</option>
					<option value="pc">pc</option>
					<option value="em">em</option>
					<option value="cm">cm</option>
					<option value="mm">mm</option>
					<option value="in">in</option>
				</select></td>
			<td style="width:50"><%= GetString("Height") %></td>
			<td><input type="text" id="tb_height" style="width:50px" />
				<select id="sel_height_unit">
					<option value="px">px</option>
					<option value="%">%</option>
					<option value="pt">pt</option>
					<option value="pc">pc</option>
					<option value="em">em</option>
					<option value="cm">cm</option>
					<option value="mm">mm</option>
					<option value="in">in</option>
				</select></td>
		</tr>
		<tr>
			<td style="width:50"><%= GetString("Left") %></td>
			<td><input type="text" id="tb_left" style="width:50px" />
				<select id="sel_left_unit">
					<option value="px">px</option>
					<option value="pt">pt</option>
					<option value="pc">pc</option>
					<option value="em">em</option>
					<option value="cm">cm</option>
					<option value="mm">mm</option>
					<option value="in">in</option>
				</select></td>
			<td style="width:50"><%= GetString("Width") %></td>
			<td><input type="text" id="tb_width" style="width:50px" />
				<select id="sel_width_unit">
					<option value="px">px</option>
					<option value="%">%</option>
					<option value="pt">pt</option>
					<option value="pc">pc</option>
					<option value="em">em</option>
					<option value="cm">cm</option>
					<option value="mm">mm</option>
					<option value="in">in</option>
				</select></td>
		</tr>
	</table>
</fieldset>
<fieldset><legend><%= GetString("Clipping") %></legend>
	<table border="0" cellspacing="0" cellpadding="2" class="normal">
		<tr>
			<td style="width:50"><%= GetString("Top") %></td>
			<td><input type="text" id="tb_cliptop" style="width:50px" />
				<select id="sel_cliptop_unit">
					<option value="px">px</option>
					<option value="%">%</option>
					<option value="pt">pt</option>
					<option value="pc">pc</option>
					<option value="em">em</option>
					<option value="cm">cm</option>
					<option value="mm">mm</option>
					<option value="in">in</option>
				</select></td>
			<td style="width:50"><%= GetString("Bottom") %></td>
			<td><input type="text" id="tb_clipbottom" style="width:50px" />
				<select id="sel_clipbottom_unit">
					<option value="px">px</option>
					<option value="%">%</option>
					<option value="pt">pt</option>
					<option value="pc">pc</option>
					<option value="em">em</option>
					<option value="cm">cm</option>
					<option value="mm">mm</option>
					<option value="in">in</option>
				</select></td>
		</tr>
		<tr>
			<td style="width:50"><%= GetString("Left") %></td>
			<td><input type="text" id="tb_clipleft" style="width:50px" />
				<select id="sel_clipleft_unit">
					<option value="px">px</option>
					<option value="%">%</option>
					<option value="pt">pt</option>
					<option value="pc">pc</option>
					<option value="em">em</option>
					<option value="cm">cm</option>
					<option value="mm">mm</option>
					<option value="in">in</option>
				</select></td>
			<td style="width:50"><%= GetString("Right") %></td>
			<td><input type="text" id="tb_clipright" style="width:50px" />
				<select id="sel_clipright_unit">
					<option value="px">px</option>
					<option value="%">%</option>
					<option value="pt">pt</option>
					<option value="pc">pc</option>
					<option value="em">em</option>
					<option value="cm">cm</option>
					<option value="mm">mm</option>
					<option value="in">in</option>
				</select></td>
		</tr>
	</table>
</fieldset>
<fieldset>
    <legend><%= GetString("Misc") %></legend>
	<div><%= GetString("Overflow") %>:
		<select id="sel_overflow">
			<option value=""><%= GetString("NotSet") %></option>
			<option value="auto"><%= GetString("OverflowAuto") %></option>
			<option value="scroll"><%= GetString("OverflowScroll") %></option>
			<option value="visible"><%= GetString("OverflowVisible") %></option>
			<option value="hidden"><%= GetString("OverflowHidden") %></option>
		</select>
		z-index: <input type="text" style="width:60px" id="tb_zindex" />
	</div>
	<table border="0" cellspacing="0" cellpadding="2" class="normal">
		<tr>
			<td style="width:120"><%= GetString("PrintingBefore") %>:</td>
			<td><select id="sel_pagebreakbefore"><option value=""><%= GetString("NotSet") %></option>
					<option value="auto"><%= GetString("Auto") %></option>
					<option value="always"><%= GetString("Always") %></option>
				</select>
			</td>		    
		</tr>
		<tr>
			<td><%= GetString("PrintingAfter") %>:
			</td>
			<td><select id="sel_pagebreakafter"><option value=""><%= GetString("NotSet") %></option>
					<option value="auto"><%= GetString("Auto") %></option>
					<option value="always"><%= GetString("Always") %></option>
				</select>
			</td>
		</tr>
	</table>
</fieldset>
<div id="outer" style="height:80px; margin-bottom:10px; padding:5px;"><div id="div_demo"><%= GetString("DemoText") %></div></div>
<br />
<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_Style_Layout.js"></script>