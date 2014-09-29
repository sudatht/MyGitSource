<!-- #include file = "Include_GetString.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title><%= GetString("InsertChars") %> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
			&nbsp; </title>
		
		<meta name="content-type" content="text/html ;charset=Unicode" />
		<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
		<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<!--[if IE]>
			<link href="../Style/IE.css" type="text/css" rel="stylesheet" />
		<![endif]-->
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>		
		<script type="text/javascript">
		
			var OxOcd27=["","\x3Ctr\x3E","\x3Ctd style=\x27height: 20; font-size: 12px; \x27 bgcolor=white width=\x2718\x27 onClick=\x27getchar(this)\x27 onmouseover=\x27spcOver(this)\x27 onmouseout=\x27spcOut(this)\x27 title=\x27","\x27 \x3E","\x26#",";","\x3C/td\x3E","\x3C/tr\x3E","background","style","#0A246A","color","white","black","Verdana","innerHTML","Unicode","\x3CFONT CLASS=\x27UNICODE\x27\x3E","\x3Cspan style=\x27font-family:","\x27\x3E","\x3C/span\x3E"];var editor=Window_GetDialogArguments(window);function cancel(){Window_CloseDialog(window);} ;var tds=22;function writeChars(){var Ox78=OxOcd27[0];for(var i=33;i<256;){document.write(OxOcd27[1]);for(var j=0;j<=tds;j++){document.write(OxOcd27[2]+i+OxOcd27[3]);document.write(OxOcd27[4]+i+OxOcd27[5]);document.write(OxOcd27[6]);i++;} ;document.write(OxOcd27[7]);} ;} ;function spcOver(Ox7b){Ox7b[OxOcd27[9]][OxOcd27[8]]=OxOcd27[10];Ox7b[OxOcd27[9]][OxOcd27[11]]=OxOcd27[12];} ;function spcOut(Ox7b){Ox7b[OxOcd27[9]][OxOcd27[8]]=OxOcd27[12];Ox7b[OxOcd27[9]][OxOcd27[11]]=OxOcd27[13];} ;function getchar(obj){var Ox7e;var Ox7f=getFontValue()||OxOcd27[14];if(!obj[OxOcd27[15]]){return ;} ;Ox7e=obj[OxOcd27[15]];if(Ox7f==OxOcd27[16]){Ox7e=OxOcd27[17]+obj[OxOcd27[15]]+OxOcd27[0];} else {if(Ox7f!=OxOcd27[14]){Ox7e=OxOcd27[18]+Ox7f+OxOcd27[19]+obj[OxOcd27[15]]+OxOcd27[20];} ;} ;editor.PasteHTML(Ox7e);Window_CloseDialog(window);} ;
		</script>
	</head>
	<body>
		<div id="container">
			<table border="0" cellspacing="2" cellpadding="2" width="99%" id="Table1">
				<tr style="display:none">
					<td class="normal">
						<%= GetString("FontName") %>: 
						<input type="radio" onpropertychange="sel_font_change()" id="selfont1" name="selfont" value="" checked="checked" />
						<label for="selfont1"><%= GetString("Default") %></label> 
						<input type="radio" onpropertychange="sel_font_change()" id="selfont2" name="selfont" value="webdings" />
						<label for="selfont2">Webdings</label>
						<input type="radio" onpropertychange="sel_font_change()" id="selfont3" name="selfont" value="wingdings" />
						<label for="selfont3">Wingdings</label>
						<input type="radio" onpropertychange="sel_font_change()" id="selfont4" name="selfont" value="symbol" />
						<label for="selfont4">Symbol</label>
						<input type="radio" onpropertychange="sel_font_change()" id="selfont5" name="selfont" value="Unicode" />
						<label for="selfont5">Unicode</label>
						<script type="text/javascript">
						var OxOd3fa=["selfont","length","checked","value","Verdana","display","style","charstable1","Unicode","block","none","charstable2","fontFamily"];function getFontValue(){var Ox82=document.getElementsByName(OxOd3fa[0]);for(var i=0;i<Ox82[OxOd3fa[1]];i++){if(Ox82[i][OxOd3fa[2]]){return Ox82[i][OxOd3fa[3]];} ;} ;} ;function sel_font_change(){var Ox84=getFontValue()||OxOd3fa[4];document.getElementById(OxOd3fa[7])[OxOd3fa[6]][OxOd3fa[5]]=(Ox84!=OxOd3fa[8]?OxOd3fa[9]:OxOd3fa[10]);document.getElementById(OxOd3fa[11])[OxOd3fa[6]][OxOd3fa[5]]=(Ox84==OxOd3fa[8]?OxOd3fa[9]:OxOd3fa[10]);document.getElementById(OxOd3fa[7])[OxOd3fa[6]][OxOd3fa[12]]=Ox84;if(Ox84==OxOd3fa[8]){} ;} ;
						</script>
					</td>
				</tr>
				<tr>
					<td align="center">
						<fieldset>
							<legend>
								<%= GetString("InsertChars") %>
							</legend>
							<br />
							<table id="charstable1" width="95%" cellspacing="1" cellpadding="1" border="0"
								 style="FONT-FAMILY: Verdana; background-color:#696969; border-color:#696969; height:222">
								<script type="text/javascript">
								var OxO24d8=[];writeChars();
								</script>
							</table>
							<table id="charstable2" width="95%" cellspacing="1" cellpadding="1" border="0"
								style="FONT-FAMILY: Verdana; background-color:#696969; border-color:#696969;display:none; height:222">
								<script type="text/javascript">
								var OxO8d27=["\x26#402;","\x26#913;","\x26#914;","\x26#915;","\x26#916;","\x26#917;","\x26#918;","\x26#919;","\x26#920;","\x26#921;","\x26#922;","\x26#923;","\x26#924;","\x26#925;","\x26#926;","\x26#927;","\x26#928;","\x26#929;","\x26#931;","\x26#932;","\x26#933;","\x26#934;","\x26#935;","\x26#936;","\x26#937;","\x26#945;","\x26#946;","\x26#947;","\x26#948;","\x26#949;","\x26#950;","\x26#951;","\x26#952;","\x26#953;","\x26#954;","\x26#955;","\x26#956;","\x26#957;","\x26#958;","\x26#959;","\x26#960;","\x26#961;","\x26#962;","\x26#963;","\x26#964;","\x26#965;","\x26#966;","\x26#967;","\x26#968;","\x26#969;","\x26#977;","\x26#978;","\x26#982;","\x26#8226;","\x26#8230;","\x26#8242;","\x26#8243;","\x26#8254;","\x26#8260;","\x26#8472;","\x26#8465;","\x26#8476;","\x26#8482;","\x26#8501;","\x26#8592;","\x26#8593;","\x26#8594;","\x26#8595;","\x26#8596;","\x26#8629;","\x26#8656;","\x26#8657;","\x26#8658;","\x26#8659;","\x26#8660;","\x26#8704;","\x26#8706;","\x26#8707;","\x26#8709;","\x26#8711;","\x26#8712;","\x26#8713;","\x26#8715;","\x26#8719;","\x26#8722;","\x26#8727;","\x26#8730;","\x26#8733;","\x26#8734;","\x26#8736;","\x26#8869;","\x26#8870;","\x26#8745;","\x26#8746;","\x26#8747;","\x26#8756;","\x26#8764;","\x26#8773;","\x26#8800;","\x26#8801;","\x26#8804;","\x26#8805;","\x26#8834;","\x26#8835;","\x26#8836;","\x26#8838;","\x26#8839;","\x26#8853;","\x26#8855;","\x26#8901;","\x26#8968;","\x26#8969;","\x26#8970;","\x26#8971;","\x26#9001;","\x26#9002;","\x26#9674;","\x26#9824;","\x26#9827;","\x26#9829;","\x26#9830;","length","\x3Ctr\x3E","\x3Ctd style=\x27height: 20; font-size: 12px; \x27 bgcolor=white width=\x2718\x27 onClick=\x27getchar(this)\x27 onmouseover=\x27spcOver(this)\x27 onmouseout=\x27spcOut(this)\x27 title=\x27"," - ","\x26","\x26amp;","\x27 \x3E","\x3C/td\x3E","\x3C/tr\x3E"];var arr=[OxO8d27[0],OxO8d27[1],OxO8d27[2],OxO8d27[3],OxO8d27[4],OxO8d27[5],OxO8d27[6],OxO8d27[7],OxO8d27[8],OxO8d27[9],OxO8d27[10],OxO8d27[11],OxO8d27[12],OxO8d27[13],OxO8d27[14],OxO8d27[15],OxO8d27[16],OxO8d27[17],OxO8d27[18],OxO8d27[19],OxO8d27[20],OxO8d27[21],OxO8d27[22],OxO8d27[23],OxO8d27[24],OxO8d27[25],OxO8d27[26],OxO8d27[27],OxO8d27[28],OxO8d27[29],OxO8d27[30],OxO8d27[31],OxO8d27[32],OxO8d27[33],OxO8d27[34],OxO8d27[35],OxO8d27[36],OxO8d27[37],OxO8d27[38],OxO8d27[39],OxO8d27[40],OxO8d27[41],OxO8d27[42],OxO8d27[43],OxO8d27[44],OxO8d27[45],OxO8d27[46],OxO8d27[47],OxO8d27[48],OxO8d27[49],OxO8d27[50],OxO8d27[51],OxO8d27[52],OxO8d27[53],OxO8d27[54],OxO8d27[55],OxO8d27[56],OxO8d27[57],OxO8d27[58],OxO8d27[59],OxO8d27[60],OxO8d27[61],OxO8d27[62],OxO8d27[63],OxO8d27[64],OxO8d27[65],OxO8d27[66],OxO8d27[67],OxO8d27[68],OxO8d27[69],OxO8d27[70],OxO8d27[71],OxO8d27[72],OxO8d27[73],OxO8d27[74],OxO8d27[75],OxO8d27[76],OxO8d27[77],OxO8d27[78],OxO8d27[79],OxO8d27[80],OxO8d27[81],OxO8d27[82],OxO8d27[83],OxO8d27[84],OxO8d27[84],OxO8d27[85],OxO8d27[86],OxO8d27[87],OxO8d27[88],OxO8d27[89],OxO8d27[90],OxO8d27[91],OxO8d27[92],OxO8d27[93],OxO8d27[94],OxO8d27[95],OxO8d27[96],OxO8d27[97],OxO8d27[97],OxO8d27[98],OxO8d27[99],OxO8d27[100],OxO8d27[101],OxO8d27[102],OxO8d27[103],OxO8d27[104],OxO8d27[105],OxO8d27[106],OxO8d27[107],OxO8d27[108],OxO8d27[90],OxO8d27[109],OxO8d27[110],OxO8d27[111],OxO8d27[112],OxO8d27[113],OxO8d27[114],OxO8d27[115],OxO8d27[116],OxO8d27[117],OxO8d27[118],OxO8d27[119],OxO8d27[120]];for(var i=0;i<arr[OxO8d27[121]];i+=tds){document.write(OxO8d27[122]);for(var j=i;j<i+tds&&j<arr[OxO8d27[121]];j++){var n=arr[j];document.write(OxO8d27[123]+n+OxO8d27[124]+n.replace(OxO8d27[125],OxO8d27[126])+OxO8d27[127]);document.write(n);document.write(OxO8d27[128]);} ;document.write(OxO8d27[129]);} ;
								</script>
							</table>
							<br />	
						</fieldset>
					</td>
				</tr>
				<tr>
					<td align="right">
					    <input type="button" value="<%= GetString("Cancel") %>" onclick="cancel()" />
					</td>
				</tr>
			</table>
	</body>
</html>
