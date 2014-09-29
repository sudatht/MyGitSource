<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<fieldset><legend><%= GetString("Insert") %></legend>
	<table class="normal">
		<tr>
			<td style="width:80"><%= GetString("Url") %>:</td>
			<td colspan="3"><input type="text" id="inp_src" size="45" /></td>
			<td>
			    <input type="button" id="btnbrowse" value="<%= GetString("Browse") %>" class="formbutton" />
			 </td>
		</tr>
		<tr>
			<td valign="middle" style="white-space:nowrap" ><%= GetString("Alternate") %>:</td>
			<td><input type="text" id="AlternateText" size="24" name="AlternateText" /></td>
			<td valign="middle" style="white-space:nowrap" ><%= GetString("ID") %>:</td>
			<td><input type="text" id="inp_id" size="12" /></td>
			<td></td>
		</tr>
		<tr>
			<td valign="middle" style="white-space:nowrap" ><%= GetString("longDesc") %>:</td>
			<td colspan="3"><input type="text" id="longDesc" size="45" name="longDesc" />
			</td>
			<td><img alt="" src="../Images/Accessibility.gif" /></td>
		</tr>
	</table>
</fieldset>
<table class="normal" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top">
			<fieldset>
				<legend>
					<%= GetString("Layout") %></legend>
				<table border="0" cellpadding="4" cellspacing="0" class="normal" width="100%">
					<tr>
						<td>
							<table border="0" cellpadding="2" cellspacing="0" class="normal" width="100%">
								<tr>
									<td style="width:60"><%= GetString("Alignment") %>:</td>
									<td>
										<select name="ImgAlign" style="WIDTH : 80px" id="Align">
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
									<td></td>
								</tr>
								<tr>
									<td><%= GetString("Bordersize") %>:</td>
									<td>
										<input type="text" size="2" name="Border" onkeypress="return CancelEventIfNotDigit()" style="WIDTH : 80px"
											id="Border" />
									</td>
									<td></td>
								</tr>
								<tr>
									<td><%= GetString("BorderColor") %>:</td>
									<td style="white-space:nowrap;">
<input autocomplete="off" type="text" id="bordercolor" name="bordercolor" size="7" style="WIDTH:57px;" />
<img alt="" src="../images/colorpicker.gif" id="bordercolor_Preview" style="vertical-align:top;" />
									</td>
									<td></td>
								</tr>
								<tr>
									<td style="white-space:nowrap; width:60"><%= GetString("Width") %>:</td>
									<td>
										<input type="text" size="2" id="inp_width" onkeyup="checkConstrains('width');" rem-skipautofirechanged="1"
											onkeypress="return CancelEventIfNotDigit()" style="WIDTH : 80px" />
									</td>
									<td rowspan="2" align="right" valign="middle"><img src="../images/locked.gif" id="imgLock" width="25" height="32" alt="<%= GetString("ConstrainProportions") %>" /></td>
								</tr>
								<tr>
									<td><%= GetString("Height") %>:</td>
									<td>
										<input type="text" size="2" id="inp_height" onkeyup="checkConstrains('width');" rem-skipautofirechanged="1"
											onkeypress="return CancelEventIfNotDigit()" style="WIDTH : 80px" />
									</td>
								</tr>
								<tr>
									<td colspan="2">
										<input type="checkbox" id="constrain_prop" checked="checked" onclick="javascript:toggleConstrains();" />
										<%= GetString("ConstrainProportions") %></td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</fieldset>
			<fieldset>
				<legend>
					<%= GetString("Spacing") %></legend>
				<table border="0" cellpadding="4" cellspacing="0" class="normal">
					<tr>
						<td>
							<table border="0" cellpadding="2" cellspacing="0" class="normal">
								<tr>
									<td valign="middle" style="width:60"><%= GetString("Horizontal") %>:</td>
									<td>
									    <input type="text" size="2" value="5" onkeypress="return CancelEventIfNotDigit()" style="WIDTH:80px" id="HSpace" />
								    </td>
								</tr>
								<tr>
									<td valign="middle"><%= GetString("Vertical") %>:</td>
									<td>
									    <input type="text" size="2" name="VSpace" onkeypress="return CancelEventIfNotDigit()" style="WIDTH:80px" id="VSpace" />
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</fieldset>
		</td>
		<td style="white-space:nowrap; width:5" >&nbsp;</td>
		<td valign="top">
			<div id="outer" style="width:230px; height:251px">
			    <img alt="" src="../Images/1x1.gif" id="img_demo" />
			</div>
		</td>
	</tr>
</table>
<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_Image.js"></script>
