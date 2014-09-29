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
	<head ID="Head1">
		<title><%= GetString("WebPalette") %></title>
		<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
		<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>
		<script type="text/javascript" src="../Scripts/Dialog/Dialog_ColorPicker.js"></script>
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<style type="text/css">
			.colorcell
			{
				width:22px;
				height:11px;
				cursor:hand;
			}
			.colordiv
			{
				border:solid 1px #808080;
				width:22px;
				height:11px;
				font-size:1px;
			}
		</style>
		<script>
var OxO9671=["0","#","length","\x3Ctr\x3E","\x3Ctd class=\x27colorcell\x27\x3E\x3Cdiv class=\x27colordiv\x27 style=\x27background-color:","\x27 cvalue=\x27","\x27 title=\x27","\x27\x3E\x26nbsp;\x3C/div\x3E\x3C/td\x3E","\x3C/tr\x3E"];function DoubleHex(Ox57){if(Ox57<16){return OxO9671[0]+Ox57.toString(16);} ;return Ox57.toString(16);} ;function ToHexString(Ox59,Ox5a,b){return (OxO9671[1]+DoubleHex(Ox59*51)+DoubleHex(Ox5a*51)+DoubleHex(b*51)).toUpperCase();} ;function MakeHex(z,x,y){var Ox60=z%2;var Ox16=(z-Ox60)/2;z=Ox60*3+Ox16;if(z<3){x=5-x;} ;if(z==1||z==4){y=5-y;} ;return ToHexString(5-y,5-x,5-z);} ;var colors= new Array(216);for(var z=0;z<6;z++){for(var x=0;x<6;x++){for(var y=0;y<6;y++){var hex=MakeHex(z,x,y);var xx=(z%2)*6+x;var yy=Math.floor(z/2)*6+y;colors[yy*12+xx]=hex;} ;} ;} ;var arr=[];for(var i=0;i<colors[OxO9671[2]];i++){if(i%12==0){arr.push(OxO9671[3]);} ;arr.push(OxO9671[4]);arr.push(colors[i]);arr.push(OxO9671[5]);arr.push(colors[i]);arr.push(OxO9671[6]);arr.push(colors[i]);arr.push(OxO9671[7]);if(i%12==11){arr.push(OxO9671[8]);} ;} ;
		</script>
	</head>
	<body>
	<div id="container">
			<div class="tab-pane-control tab-pane" id="tabPane1">
				<div class="tab-row">
					<h2 class="tab selected">
						<a tabindex="-1" href='colorpicker.asp?<%=GetDialogQueryString%>'>
							<span style="white-space:nowrap;">
								<%= GetString("WebPalette") %>
							</span>
						</a>
					</h2>
					<h2 class="tab">
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
					<table cellSpacing='2' cellPadding="1" align="center">
						<script>
							var OxOe82b=[""];document.write(arr.join(OxOe82b[0]));
						</script>
						<tr>
							<td colspan="12" height="12"><p align="left"></p>
							</td>
						</tr>
						<tr>
							<td colspan="12" valign="middle" height="24">								
					<span style="height:24px;width:50px;vertical-align:middle;"><%= GetString("Color") %>: </span>&nbsp;
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