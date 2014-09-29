<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<% 

	' Connect to database
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
	Conn.CommandTimeOut = Session("xtra_CommandTimeout")
	Conn.Open Session("xtra_ConnectionString"),Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

	strBrukerID	= Session("BrukerID")
	
	'response.write("strBrukerID"  & strBrukerID)
	strSQL = "select fagreementhandler from bruker where id = " & strBrukerID 
	Set rsUser = Conn.Execute(strSQL)
	
	handler = rsUser("fagreementhandler")	
	'response.write("handler" & handler)
	
	rsUser.close
	set rsUser = nothing
	Conn.Close()
	set Conn = nothing

If (HasUserRight(ACCESS_ADMIN, RIGHT_ADMIN) = false) Then 
	call Response.Redirect("/xtra/IngenTilgang.asp")	
End If 
%>
<html>
	<head>
		<title>Administrasjon</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		
		<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>
			//popup window - about
			var popupWindow=null;
			function popup(mypage,myname,w,h,pos,infocus){
			
			if (pos == 'random')
			{LeftPosition=(screen.width)?Math.floor(Math.random()*(screen.width-w)):100;TopPosition=(screen.height)?Math.floor(Math.random()*((screen.height-h)-75)):100;}
			else
			{LeftPosition=(screen.width)?(screen.width-w)/2:100;TopPosition=(screen.height)?(screen.height-h)/2:100;}
			settings='width='+ w + ',height='+ h + ',top=' + TopPosition + ',left=' + LeftPosition + ',menubar=no,toolbar=no';popupWindow=window.open(mypage,myname,settings);
			//if(infocus=='front'){popupWindow.focus();popupWindow.location=mypage;}
			//if(infocus=='back'){popupWindow.blur();popupWindow.location=mypage;popupWindow.blur();}
			
			}
		</SCRIPT>

	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				
				<!-- about button -->
				    <table style="width:99%;" cellpadding="0" cellspacing="0">
				        <tr>
				            <td style="width:97%; text-align:left;"><h1>Administrasjon</h1></td>
				            <td style="width:3%; text-align:right;">
				                <div onclick="window.open('/xtra/About.asp', 'About', 'width=450,height=300,toolbar=no,menubar=no,location=no')" style="width:20px; height:26px;"></div>
				                <!--div onclick="javascript:popup('../About.asp','pagename','400','300','center','front');"></div-->
				            </td>
				        </tr>
				    </table>
			</div>
			<div class="content content2">
				<ul>
					<%
					If (HasUserRight(ACCESS_ADMIN, RIGHT_SUPER) = true) Then
						%>
						<li><a href="../Admin/RettighetSub.asp">Brukeradministrasjon</a></li>
						<li><a href="../Admin/Debug.asp">Debug</a></li>
						<%
					end if
					If (handler) Then
						%>
						<li><a href="../WebUI/SearchFa.aspx">Rammeavtaler</a></li>
						<%
					end if
					If (HasUserRight(ACCESS_ADMIN, RIGHT_SUPER) = true) Then
					%>
						<li><a href="../WebUI/DataDeletion/AdminDeleteSummary.aspx">Slett vikarinfo</a></li>
					<%
					end if
					If (HasUserRight(ACCESS_ADMIN, RIGHT_SUPER) = true) Then
					%>
						<li><a href="../WebUI/Admin/Administration/EmailTemplate/SearchTemplate.aspx">E-postmal</a></li>
					<%
					end if
					%>					
					
				</ul>
				<!--<p><a href="Rutiner.html" target="_new">Rutinebeskrivelse</a></p>-->
				<br>
				<!--<p><a href="../utvikling/BrowserPopper.asp" >Browser popper test</a></p>-->
			</div>
		</div>
	</body>
</html>