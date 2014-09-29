<!-- #include file = "Include_GetString.asp" -->
<%
dim GetDialogQueryString
Theme="Office2007"

GetDialogQueryString = "Theme=Office2007"
if Request.QueryString("Dialog") = "Standard" then
    GetDialogQueryString=GetDialogQueryString & "&Dialog=Standard"
End If
if Request.QueryString("setting") <> "" then
    GetDialogQueryString=GetDialogQueryString & "&setting=" & Request.QueryString("setting")
end if 
%>
<html xmlns="http://www.w3.org/1999/xhtml">
	<head runat="server">
		<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
		<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>
		<script type="text/javascript" src="../Scripts/Dialog/Dialog_ColorPicker.js"></script>
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<style type="text/css">
			.colorcell
			{
				width:16px;
				height:17px;
				cursor:hand;
			}
			.colordiv,.customdiv
			{
				border:solid 1px #808080;
				width:16px;
				height:17px;
				font-size:1px;
			}
		</style>
		<title><%= GetString("NamedColors") %></title>
		<script>
								
		var OxO176f=["Green","#008000","Lime","#00FF00","Teal","#008080","Aqua","#00FFFF","Navy","#000080","Blue","#0000FF","Purple","#800080","Fuchsia","#FF00FF","Maroon","#800000","Red","#FF0000","Olive","#808000","Yellow","#FFFF00","White","#FFFFFF","Silver","#C0C0C0","Gray","#808080","Black","#000000","DarkOliveGreen","#556B2F","DarkGreen","#006400","DarkSlateGray","#2F4F4F","SlateGray","#708090","DarkBlue","#00008B","MidnightBlue","#191970","Indigo","#4B0082","DarkMagenta","#8B008B","Brown","#A52A2A","DarkRed","#8B0000","Sienna","#A0522D","SaddleBrown","#8B4513","DarkGoldenrod","#B8860B","Beige","#F5F5DC","HoneyDew","#F0FFF0","DimGray","#696969","OliveDrab","#6B8E23","ForestGreen","#228B22","DarkCyan","#008B8B","LightSlateGray","#778899","MediumBlue","#0000CD","DarkSlateBlue","#483D8B","DarkViolet","#9400D3","MediumVioletRed","#C71585","IndianRed","#CD5C5C","Firebrick","#B22222","Chocolate","#D2691E","Peru","#CD853F","Goldenrod","#DAA520","LightGoldenrodYellow","#FAFAD2","MintCream","#F5FFFA","DarkGray","#A9A9A9","YellowGreen","#9ACD32","SeaGreen","#2E8B57","CadetBlue","#5F9EA0","SteelBlue","#4682B4","RoyalBlue","#4169E1","BlueViolet","#8A2BE2","DarkOrchid","#9932CC","DeepPink","#FF1493","RosyBrown","#BC8F8F","Crimson","#DC143C","DarkOrange","#FF8C00","BurlyWood","#DEB887","DarkKhaki","#BDB76B","LightYellow","#FFFFE0","Azure","#F0FFFF","LightGray","#D3D3D3","LawnGreen","#7CFC00","MediumSeaGreen","#3CB371","LightSeaGreen","#20B2AA","DeepSkyBlue","#00BFFF","DodgerBlue","#1E90FF","SlateBlue","#6A5ACD","MediumOrchid","#BA55D3","PaleVioletRed","#DB7093","Salmon","#FA8072","OrangeRed","#FF4500","SandyBrown","#F4A460","Tan","#D2B48C","Gold","#FFD700","Ivory","#FFFFF0","GhostWhite","#F8F8FF","Gainsboro","#DCDCDC","Chartreuse","#7FFF00","LimeGreen","#32CD32","MediumAquamarine","#66CDAA","DarkTurquoise","#00CED1","CornflowerBlue","#6495ED","MediumSlateBlue","#7B68EE","Orchid","#DA70D6","HotPink","#FF69B4","LightCoral","#F08080","Tomato","#FF6347","Orange","#FFA500","Bisque","#FFE4C4","Khaki","#F0E68C","Cornsilk","#FFF8DC","Linen","#FAF0E6","WhiteSmoke","#F5F5F5","GreenYellow","#ADFF2F","DarkSeaGreen","#8FBC8B","Turquoise","#40E0D0","MediumTurquoise","#48D1CC","SkyBlue","#87CEEB","MediumPurple","#9370DB","Violet","#EE82EE","LightPink","#FFB6C1","DarkSalmon","#E9967A","Coral","#FF7F50","NavajoWhite","#FFDEAD","BlanchedAlmond","#FFEBCD","PaleGoldenrod","#EEE8AA","Oldlace","#FDF5E6","Seashell","#FFF5EE","PaleGreen","#98FB98","SpringGreen","#00FF7F","Aquamarine","#7FFFD4","PowderBlue","#B0E0E6","LightSkyBlue","#87CEFA","LightSteelBlue","#B0C4DE","Plum","#DDA0DD","Pink","#FFC0CB","LightSalmon","#FFA07A","Wheat","#F5DEB3","Moccasin","#FFE4B5","AntiqueWhite","#FAEBD7","LemonChiffon","#FFFACD","FloralWhite","#FFFAF0","Snow","#FFFAFA","AliceBlue","#F0F8FF","LightGreen","#90EE90","MediumSpringGreen","#00FA9A","PaleTurquoise","#AFEEEE","LightCyan","#E0FFFF","LightBlue","#ADD8E6","Lavender","#E6E6FA","Thistle","#D8BFD8","MistyRose","#FFE4E1","Peachpuff","#FFDAB9","PapayaWhip","#FFEFD5"];var colorlist=[{n:OxO176f[0],h:OxO176f[1]},{n:OxO176f[2],h:OxO176f[3]},{n:OxO176f[4],h:OxO176f[5]},{n:OxO176f[6],h:OxO176f[7]},{n:OxO176f[8],h:OxO176f[9]},{n:OxO176f[10],h:OxO176f[11]},{n:OxO176f[12],h:OxO176f[13]},{n:OxO176f[14],h:OxO176f[15]},{n:OxO176f[16],h:OxO176f[17]},{n:OxO176f[18],h:OxO176f[19]},{n:OxO176f[20],h:OxO176f[21]},{n:OxO176f[22],h:OxO176f[23]},{n:OxO176f[24],h:OxO176f[25]},{n:OxO176f[26],h:OxO176f[27]},{n:OxO176f[28],h:OxO176f[29]},{n:OxO176f[30],h:OxO176f[31]}];var colormore=[{n:OxO176f[32],h:OxO176f[33]},{n:OxO176f[34],h:OxO176f[35]},{n:OxO176f[36],h:OxO176f[37]},{n:OxO176f[38],h:OxO176f[39]},{n:OxO176f[40],h:OxO176f[41]},{n:OxO176f[42],h:OxO176f[43]},{n:OxO176f[44],h:OxO176f[45]},{n:OxO176f[46],h:OxO176f[47]},{n:OxO176f[48],h:OxO176f[49]},{n:OxO176f[50],h:OxO176f[51]},{n:OxO176f[52],h:OxO176f[53]},{n:OxO176f[54],h:OxO176f[55]},{n:OxO176f[56],h:OxO176f[57]},{n:OxO176f[58],h:OxO176f[59]},{n:OxO176f[60],h:OxO176f[61]},{n:OxO176f[62],h:OxO176f[63]},{n:OxO176f[64],h:OxO176f[65]},{n:OxO176f[66],h:OxO176f[67]},{n:OxO176f[68],h:OxO176f[69]},{n:OxO176f[70],h:OxO176f[71]},{n:OxO176f[72],h:OxO176f[73]},{n:OxO176f[74],h:OxO176f[75]},{n:OxO176f[76],h:OxO176f[77]},{n:OxO176f[78],h:OxO176f[79]},{n:OxO176f[80],h:OxO176f[81]},{n:OxO176f[82],h:OxO176f[83]},{n:OxO176f[84],h:OxO176f[85]},{n:OxO176f[86],h:OxO176f[87]},{n:OxO176f[88],h:OxO176f[89]},{n:OxO176f[90],h:OxO176f[91]},{n:OxO176f[92],h:OxO176f[93]},{n:OxO176f[94],h:OxO176f[95]},{n:OxO176f[96],h:OxO176f[97]},{n:OxO176f[98],h:OxO176f[99]},{n:OxO176f[100],h:OxO176f[101]},{n:OxO176f[102],h:OxO176f[103]},{n:OxO176f[104],h:OxO176f[105]},{n:OxO176f[106],h:OxO176f[107]},{n:OxO176f[108],h:OxO176f[109]},{n:OxO176f[110],h:OxO176f[111]},{n:OxO176f[112],h:OxO176f[113]},{n:OxO176f[114],h:OxO176f[115]},{n:OxO176f[116],h:OxO176f[117]},{n:OxO176f[118],h:OxO176f[119]},{n:OxO176f[120],h:OxO176f[121]},{n:OxO176f[122],h:OxO176f[123]},{n:OxO176f[124],h:OxO176f[125]},{n:OxO176f[126],h:OxO176f[127]},{n:OxO176f[128],h:OxO176f[129]},{n:OxO176f[130],h:OxO176f[131]},{n:OxO176f[132],h:OxO176f[133]},{n:OxO176f[134],h:OxO176f[135]},{n:OxO176f[136],h:OxO176f[137]},{n:OxO176f[138],h:OxO176f[139]},{n:OxO176f[140],h:OxO176f[141]},{n:OxO176f[142],h:OxO176f[143]},{n:OxO176f[144],h:OxO176f[145]},{n:OxO176f[146],h:OxO176f[147]},{n:OxO176f[148],h:OxO176f[149]},{n:OxO176f[150],h:OxO176f[151]},{n:OxO176f[152],h:OxO176f[153]},{n:OxO176f[154],h:OxO176f[155]},{n:OxO176f[156],h:OxO176f[157]},{n:OxO176f[158],h:OxO176f[159]},{n:OxO176f[160],h:OxO176f[161]},{n:OxO176f[162],h:OxO176f[163]},{n:OxO176f[164],h:OxO176f[165]},{n:OxO176f[166],h:OxO176f[167]},{n:OxO176f[168],h:OxO176f[169]},{n:OxO176f[170],h:OxO176f[171]},{n:OxO176f[172],h:OxO176f[173]},{n:OxO176f[174],h:OxO176f[175]},{n:OxO176f[176],h:OxO176f[177]},{n:OxO176f[178],h:OxO176f[179]},{n:OxO176f[180],h:OxO176f[181]},{n:OxO176f[182],h:OxO176f[183]},{n:OxO176f[184],h:OxO176f[185]},{n:OxO176f[186],h:OxO176f[187]},{n:OxO176f[188],h:OxO176f[189]},{n:OxO176f[190],h:OxO176f[191]},{n:OxO176f[192],h:OxO176f[193]},{n:OxO176f[194],h:OxO176f[195]},{n:OxO176f[196],h:OxO176f[197]},{n:OxO176f[198],h:OxO176f[199]},{n:OxO176f[200],h:OxO176f[201]},{n:OxO176f[202],h:OxO176f[203]},{n:OxO176f[204],h:OxO176f[205]},{n:OxO176f[206],h:OxO176f[207]},{n:OxO176f[208],h:OxO176f[209]},{n:OxO176f[210],h:OxO176f[211]},{n:OxO176f[212],h:OxO176f[213]},{n:OxO176f[214],h:OxO176f[215]},{n:OxO176f[216],h:OxO176f[217]},{n:OxO176f[218],h:OxO176f[219]},{n:OxO176f[220],h:OxO176f[221]},{n:OxO176f[156],h:OxO176f[157]},{n:OxO176f[222],h:OxO176f[223]},{n:OxO176f[224],h:OxO176f[225]},{n:OxO176f[226],h:OxO176f[227]},{n:OxO176f[228],h:OxO176f[229]},{n:OxO176f[230],h:OxO176f[231]},{n:OxO176f[232],h:OxO176f[233]},{n:OxO176f[234],h:OxO176f[235]},{n:OxO176f[236],h:OxO176f[237]},{n:OxO176f[238],h:OxO176f[239]},{n:OxO176f[240],h:OxO176f[241]},{n:OxO176f[242],h:OxO176f[243]},{n:OxO176f[244],h:OxO176f[245]},{n:OxO176f[246],h:OxO176f[247]},{n:OxO176f[248],h:OxO176f[249]},{n:OxO176f[250],h:OxO176f[251]},{n:OxO176f[252],h:OxO176f[253]},{n:OxO176f[254],h:OxO176f[255]},{n:OxO176f[256],h:OxO176f[257]},{n:OxO176f[258],h:OxO176f[259]},{n:OxO176f[260],h:OxO176f[261]},{n:OxO176f[262],h:OxO176f[263]},{n:OxO176f[264],h:OxO176f[265]},{n:OxO176f[266],h:OxO176f[267]},{n:OxO176f[268],h:OxO176f[269]},{n:OxO176f[270],h:OxO176f[271]},{n:OxO176f[272],h:OxO176f[273]}];
		
		</script>
	</head>
	<body>
		<div id="container">
			<div class="tab-pane-control tab-pane" id="tabPane1">
				<div class="tab-row">
					<h2 class="tab">
						<a tabindex="-1" href='colorpicker.asp?<%=GetDialogQueryString%>'>
							<span style="white-space:nowrap;">
								<%= GetString("WebPalette") %>
							</span>
						</a>
					</h2>
					<h2 class="tab selected">
							<a tabindex="-1" href='colorpicker_basic.asp?<%=GetDialogQueryString%>'>
								<span style="white-space:nowrap;">
									<%= GetString("NamedColors") %>
								</span>
							</a>
					</h2>
					<h2 class="tab">
							<a tabindex="-1" href='colorpicker_more.asp?<%=GetDialogQueryString%>'>
								<span style="white-space:nowrap;">
									<%= GetString("CustomColor") %>
								</span>
							</a>
					</h2>
				</div>
				<div class="tab-page">			
					<table class="colortable" align="center">
						<tr>
							<td colspan="16" height="16"><p align="left">Basic:
								</p>
							</td>
						</tr>
						<tr>
							<script>
								var OxOd5f0=["length","\x3Ctd class=\x27colorcell\x27\x3E\x3Cdiv class=\x27colordiv\x27 style=\x27background-color:","\x27 title=\x27"," ","\x27 cname=\x27","\x27 cvalue=\x27","\x27\x3E\x3C/div\x3E\x3C/td\x3E",""];var arr=[];for(var i=0;i<colorlist[OxOd5f0[0]];i++){arr.push(OxOd5f0[1]);arr.push(colorlist[i].n);arr.push(OxOd5f0[2]);arr.push(colorlist[i].n);arr.push(OxOd5f0[3]);arr.push(colorlist[i].h);arr.push(OxOd5f0[4]);arr.push(colorlist[i].n);arr.push(OxOd5f0[5]);arr.push(colorlist[i].h);arr.push(OxOd5f0[6]);} ;document.write(arr.join(OxOd5f0[7]));
							</script>
						</tr>
						<tr>
							<td colspan="16" height="12"><p align="left"></p>
							</td>
						</tr>
						<tr>
							<td colspan="16"><p align="left">Additional:
								</p>
							</td>
						</tr>
						<script>
							var OxOeff6=["length","\x3Ctr\x3E","\x3Ctd class=\x27colorcell\x27\x3E\x3Cdiv class=\x27colordiv\x27 style=\x27background-color:","\x27 title=\x27"," ","\x27 cname=\x27","\x27 cvalue=\x27","\x27\x3E\x3C/div\x3E\x3C/td\x3E","\x3C/tr\x3E",""];var arr=[];for(var i=0;i<colormore[OxOeff6[0]];i++){if(i%16==0){arr.push(OxOeff6[1]);} ;arr.push(OxOeff6[2]);arr.push(colormore[i].n);arr.push(OxOeff6[3]);arr.push(colormore[i].n);arr.push(OxOeff6[4]);arr.push(colormore[i].h);arr.push(OxOeff6[5]);arr.push(colormore[i].n);arr.push(OxOeff6[6]);arr.push(colormore[i].h);arr.push(OxOeff6[7]);if(i%16==15){arr.push(OxOeff6[8]);} ;} ;if(colormore%16>0){arr.push(OxOeff6[8]);} ;document.write(arr.join(OxOeff6[9]));
						</script>
						<tr>
							<td colspan="16" height="8">
							</td>
						</tr>
						<tr>
							<td colspan="16" height="12">
								<input checked id="CheckboxColorNames" style="width: 16px; height: 20px" type="checkbox">
								<span style="width: 118px;">Use color names</span>
							</td>
						</tr>
						<tr>
							<td colspan="16" height="12">
							</td>
						</tr>
						<tr>
							<td colspan="16" valign="middle" height="24">
							<span style="height:24px;width:50px;vertical-align:middle;">Color : </span>&nbsp;
							<input type="text" id="divpreview" size="7" maxlength="7" style="width:180px;height:24px;border:#a0a0a0 1px solid; Padding:4;"/>
					
							</td>
						</tr>
				</table>
			</div>
		</div>
		<div id="container-bottom">
			<input type="button" id="buttonok" value="<%= GetString("OK") %>" class="formbutton" style="width:70px"	onclick="do_insert();" /> 
			&nbsp;&nbsp;&nbsp;&nbsp; 
			<input type="button" id="buttoncancel" value="<%= GetString("Cancel") %>" class="formbutton" style="width:70px"	onclick="do_Close();" />	
		</div>
	</div>
	</body>
</html>

