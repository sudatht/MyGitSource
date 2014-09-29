<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<table border="0" cellspacing="2" cellpadding="5" width="100%">
	<tr>
		<td>
			<div>
				<fieldset style="padding:2px"><legend><%= GetString("InsertHorizontalRule") %></legend>
					<table border="0" cellpadding="5" cellspacing="0">
						<tr>
							<td style='vertical-align:middle' class="normal"><%= GetString("Width") %>:</td>
							<td style='vertical-align:middle'>
								<input type="text" id="inp_width" style="width:45px" onkeypress="return CancelEventIfNotDigit()" />
								<select id="eenheid">
									<option selected="selected" value="%">%</option>
									<option value="">px</option>
								</select>
							</td>
							<td style="width:3">&nbsp;
							</td>
							<td style='vertical-align:middle' class="normal"><%= GetString("Alignment") %>:</td>
							<td style='vertical-align:middle'>
								<select style="WIDTH: 68px" id="alignment">
									<option value="left">Left</option>
									<option value="center">Center</option>
									<option value="right">Right</option>
								</select>
							</td>
						</tr>
						<tr>
							<td class="normal"><%= GetString("Color") %>:</td>
							<td>
<input autocomplete="off" type="text" id="hrcolor" name="hrcolor" size="7" style="WIDTH:57px;" />
<img alt="" src="../images/colorpicker.gif" id="hrcolorpreview" style="vertical-align:top;" />
							</td>
							<td style="width:3">&nbsp;
							</td>
							<td style='vertical-align:middle' class="normal"><%= GetString("Shade") %>:</td>
							<td style='vertical-align:middle'>
								<select id="shade">
									<option value="shade" selected="selected">Shade</option>
									<option value="noshade">No shade</option>
								</select>
							</td>
						</tr>
						<tr>
							<td style='vertical-align:middle' class="normal"><%= GetString("Size") %>:</td>
							<td style='vertical-align:middle' colspan="4">
								<select id="sel_size" style="width:67px">
									<option selected="selected" value="1">1 px</option>
									<option value="2">2 px</option>
									<option value="3">3 px</option>
									<option value="4">4 px</option>
									<option value="5">5 px</option>
									<option value="6">6 px</option>
									<option value="7">7 px</option>
									<option value="8">8 px</option>
									<option value="9">9 px</option>
									<option value="10">10 px</option>
								</select>
							</td>
						</tr>
					</table>
				</fieldset>
			</div>
		</td>
	</tr>
</table>
<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_Hr.js"></script>
