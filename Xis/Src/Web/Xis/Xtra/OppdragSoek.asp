<%@ LANGUAGE="VBSCRIPT" %>
<%Option explicit%>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.Settings.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim StrSQL
	dim Conn
	dim rsStatus
	dim rsTjenesteOmrade 
	dim rsKompetanse
	dim strLoad

	'Hotlist menu items variables
	dim strAddToHotlistLink
	dim strHotlistType
	dim blnShowHotList

	blnShowHotList = false
	strAddToHotlistLink = ""
	strHotlistType = ""		
	
	' Get a database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<title>Søk oppdrag</title>
		<script type="text/javascript" src="Js/javascript.js"></script>
		<script type="text/javascript" src="js/contentMenu.js"></script>
		<script language='javascript' src='js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='js/navigation.js' id='navigationScripts'></script>	
		<script language='javascript' src='js/ecajax.js'></script>
		<base TARGET="_self">
		<script language="javaScript" type="text/javascript">
			var handleCallback = function (result) 
			{
				if(result) 
				{
					var str = "<select NAME='dbxMedarbeider' ID='dbxMedarbeider' onkeydown='typeAhead()'><option VALUE='0'></option>";
					str = str.concat(result,"</SELECT>")
					document.getElementById('divResponsible').innerHTML = str;
				}
			};
		
			function ShowAll(ansID)
			{
				var url;
				if(document.all.chkShowAllRes.checked)			
					url = "Callback.asp?FuncName=GetAllCoWorkersAsOptionList&SelectedID=";			
				else			
					url = "Callback.asp?FuncName=GetActiveCoWorkersAsOptionList&SelectedID=";							
				url = url.concat(ansID);
				
				asynchronousCall(url,handleCallback);
			
			}
			function shortKey(e) 
			{	
				var keyChar = String.fromCharCode(event.keyCode);
				var modKey  = event.ctrlKey;
				var modKey2 = event.shiftKey;
				
				if (event.keyCode == 13)
				{
					document.all.frmJobSearch.submit();
				}				
				if (modKey && modKey2 && keyChar=="S")
				{	
					parent.frames[funcFrameIndex].location=("/xtra/OppdragSoek.asp");
				}
			}
			//her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;			
		</script>	
	</head>
	<% 
	if (request("hdnPosted")="1") then 
		strLoad = "onLoad='javascript:document.all.Result.scrollIntoView(true);'" 
	else
		strLoad = "onLoad='javascript:document.all.lnk0.focus();'"
	end if
	%>
	<body <%=strLoad%>>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>S&oslash;k etter oppdrag</h1>
				<div class="contentMenu">
					<table cellpadding="0" cellspacing="0" width="96%" ID="Table1">
						<tr>
							<td>				
								<table cellpadding="0" cellspacing="2" ID="Table2">
									<tr>
										<td class="menu" id="menu1" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
											<img src="/xtra/images/icon_search.gif" width="18" height="15" alt="" align="absmiddle">
											<a onClick="javascript:document.all.frmJobSearch.submit();" href="#" title="Søke etter oppdrag">Utfør søk</a>
										</td>
										<td class="menu" id="menu2" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
											<a onClick="javascript:window.location='oppdragsoek.asp';" href="#" title="Blanker ut alle feltene">Blank ut</a>
										</td>
									</tr>
								</table>
							</td>
							<td class="right">
								<!--#include file="Includes/contentToolsMenu.asp"-->
							</td>							
						</tr>
					</table>
				</div>
			</div>
			<div class="content">
				<form name="frmJobSearch" ACTION="OppdragListe.asp" METHOD="POST">
					<table>
						<tr>
							<td>Vikar etternavn:</td>
							<td><input ID="lnk0"  NAME="tbxEtternavn" TYPE="TEXT" SIZE="15" MAXLENGTH="49"> </td>
							<td>Ansattnummer:</td>
							<td><input ID="lnk1"  NAME="tbxAnsattnr" TYPE="TEXT" SIZE="5" MAXLENGTH="9"></td>
							<td>Kontakt:</td>
							<td><input ID="lnk2"  NAME="tbxFirma" TYPE="TEXT" SIZE="15" MAXLENGTH="49"> </td>
						</tr>
						<tr>
							<td>Kontaktnr:</td>
							<td><input ID="lnk3"  NAME="tbxFirmaID" TYPE="TEXT" SIZE="15" MAXLENGTH="9"></td>
							<td>Oppdragnr:</td>
							<td><input ID="lnk4"  NAME="tbxOppdragID" TYPE="TEXT" SIZE="5" MAXLENGTH="9"></td>
							<td>Beskrivelse:</td>
							<td><input ID="lnk15"  NAME="tbxDescription" TYPE="TEXT" SIZE="15" MAXLENGTH="100"></td>
						</tr>
						<tr>
							<td>F.o.m dato:</td>
							<td><input ID="lnk5"  NAME="tbxFradato" TYPE="TEXT" SIZE="10" MAXLENGTH="10" ONBLUR="dateCheck(this.form, this.name)"> </td>
							<td>T.o.m dato:</td>
							<td><input ID="lnk6"  NAME="tbxTilDato" TYPE="TEXT" SIZE="10" MAXLENGTH="10" ONBLUR="dateCheck(this.form, this.name)"> </td>
						</tr>
						<tr>
							<td>Status:</td>
							<td>
								<select NAME="dbxStatus" ID="lnk7" >
									<option VALUE="0"></option>
									<%
									' Get Status
									StrSQL = "Select OppdragsStatusID, OppdragsStatus from h_oppdrag_status"
									set rsStatus = GetFirehoseRS(StrSQL, Conn)

									Do Until rsStatus.EOF 
										%>
										<option VALUE="<% =rsStatus("OppdragsStatusID") %>"><% =rsStatus("OppdragsStatus") %></option>
										<%   
										rsStatus.MoveNext
									Loop

									' Close and release recordset
									rsStatus.Close
									Set rsStatus = Nothing
									%>
								</select>
							</td>
							<td>Avdelingskontor:</td>
							<td>
								<select NAME="dbxAvdeling" ID="lnk8" >
									<option value="0"></option>
									<%
									Dim rsAvdeling 'as adodb.recordset
									'Fred 31.08.2000, endring fra 'avdeling' til 'avdelingskontor'.
									' Get avdelingskontor
									StrSQL = "select id, navn from avdelingskontor where show_hide = 1 order by id"
									set rsAvdeling = GetFirehoseRS(StrSQL, Conn)

									Do Until rsAvdeling.EOF
										Response.Write "<option VALUE='" & rsAvdeling("ID") & "' >" & rsAvdeling("navn") & "</option>" & chr(13)
										rsAvdeling.MoveNext
									Loop
									rsAvdeling.Close
									Set rsAvdeling = Nothing
									%>
								</select>
							</td>
							<td>Tjenesteomr&aring;de:</td>
							<td>
								<select NAME="dbxtjenesteomrade" ID="lnk9" >
									<option VALUE="0"></option>
									<%
									StrSQL = "SELECT tomid, navn FROM tjenesteomrade ORDER BY tomid"
									set rsTjenesteOmrade = GetFirehoseRS(StrSQL, Conn)
									Do Until rsTjenesteOmrade.EOF
										Response.Write "<option VALUE='" & rsTjenesteOmrade("tomID") & "' >" & rsTjenesteOmrade("navn") & "</option>"
										rsTjenesteOmrade.MoveNext
									Loop
									' Close and release recordset
									rsTjenesteOmrade.Close
									Set rsTjenesteOmrade = Nothing
									%>
								</select>
							</td>
						</tr>
						<tr>
							<td>Ansvarlig:</td>
							<td>
								<div id="divResponsible">
								<select NAME="dbxMedarbeider" ID="dbxMedarbeider" onkeydown="typeAhead()">
									<option VALUE="0"></option>
									<%
     								response.write GetCoWorkersAsOptionList(-2)
									%>
								</select>
								</div>
							</td>
							<td>
							<input id='chkShowAllRes' name='chkShowAllRes' type='checkbox' class='checkbox' Value='1' onClick="ShowAll(-2);">Vis Alle
							</td>
						</tr>
					</table>
				</form>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>
