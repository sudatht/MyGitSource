<%@ LANGUAGE="VBSCRIPT" %>
<%Option explicit%>
<%
	dim strSQL
	dim oRSCV
	dim oCon
	dim strSearchPhrase
	dim AWordArray
	dim strContainsClause
	dim nArraySize
	dim nCurrentIndex
	dim strSokeType
	dim strSelected1
	dim strSelected2
	dim strSelected3
	dim strHeading : strHeading = "Fil søk"

	'Hotlist menu items variables
	dim strAddToHotlistLink
	dim strHotlistType
	dim blnShowHotList

	blnShowHotList = false
	strAddToHotlistLink = ""
	strHotlistType = ""		


if (Request("hdnPosted")="1") then
	strSearchPhrase = trim(Request("txtPhrase"))
	If (len(strSearchPhrase)>0) then
		strSokeType = lcase(Request("cboSokeType"))
		if (strSokeType="1") then
			AWordArray = split(strSearchPhrase," ")
			nArraySize = ubound(AWordArray)
			nCurrentIndex = 0
			while (nCurrentIndex<=nArraySize)
				strContainsClause = strContainsClause & " """ & AWordArray(nCurrentIndex) & """ OR "
				nCurrentIndex = nCurrentIndex + 1
			wend
			strContainsClause = mid(strContainsClause,1, len(strContainsClause)-4)			
		elseif (strSokeType="2") then
			AWordArray = split(strSearchPhrase," ")
			nArraySize = ubound(AWordArray)
			nCurrentIndex = 0
			while (nCurrentIndex<=nArraySize)
				strContainsClause = strContainsClause & " """ & AWordArray(nCurrentIndex) & """ AND "				
				nCurrentIndex = nCurrentIndex + 1
			wend
			strContainsClause = mid(strContainsClause,1, len(strContainsClause)-5)				
		elseif (strSokeType="3") then
			strContainsClause = """" & strSearchPhrase & """"
		end if
		'Build search string
		strSQL = "SELECT filename, path, Size, write, rank, Directory  FROM SCOPE() WHERE CONTAINS('" & strContainsClause & "') ORDER BY write desc"
	
		'Response.Write "strSQL:" & strSQL & "<br>"
		'Response.Write "strSearchPhrase:" & strSearchPhrase & "<br>"		
		'Response.Write "cboSokeType:" & Request("cboSokeType") & "<br>"		
		'Response.End
		
		'Prepare for dataacess..
		On Error Resume Next
		Set oCon = Server.CreateObject("ADODB.Connection")
		oCon.Open("provider=MSIDXS;Data Source=DocArchieve;")
		set oRSCV = oCon.Execute(strSQL)	
	end if
else
	strSearchPhrase = ""
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"
    "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<title><%=strHeading%></title>
	</head>
	<script language="javascript">
		function ChangeSearch()
		{
			window.location = "VikarSoek.asp"
		}
	</script>
	<script type="text/javascript" src="/xtra/js/CVFunctions.js"></script>
	<script type="text/javascript" src="/xtra/js/contentMenu.js"></script>
	<script type="text/javascript" src="/xtra/js/fontSizer.js"></script>
	<BODY>
		<Form name="frmSearch" ID="Form1">
			<input type="hidden" name="hdnPosted" value="1" ID="Hidden1">
			<div class="pageContainer" id="pageContainer">
				<div class="contentHead1">
					<h1><%=strHeading%></h1>
					<div class="contentMenu">
						<table cellpadding="0" cellspacing="0" width="96%" ID="Table2">
							<tr>
								<td>				
									<table cellpadding="0" cellspacing="2" ID="Table3">
										<tr>
											<td class="menu" id="menu6" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
												Søk i
												<select ID="Select2" NAME="cboSearchIn" onchange="javascript:ChangeSearch()">
													<option>Database</option>
													<option selected>Dokumentarkiv</option>
												</select>
											</td>
											<td class="menu" id="menu7" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
												<img src="/xtra/images/icon_search.gif" width="18" height="15" alt="" align="absmiddle">
												<a onClick="javascript:document.all.frmSearch.submit();" href="#" title="Søke i filer">Utfør søk</a>
											</td>
											<td class="menu" id="menu8" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
												<a onClick="javascript:window.location='HTMLSearch.asp';" href="#" title="Blanker ut alle feltene">Blank ut</a>
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
					<strong>S&oslash;k etter:</strong>
					<input type="text" size="50" value="<%=strSearchPhrase%>" name="txtPhrase" maxlength="100" ID="Text1">
					<select name="cboSokeType" ID="Select1">
						<%
						if strSokeType="1" then
							strSelected1 = "selected"
							strSelected2 = ""
							strSelected3 = ""
						elseif strSokeType="2" then
							strSelected1 = ""
							strSelected2 = "selected"
							strSelected3 = ""
						elseif strSokeType="3" then
							strSelected1 = ""
							strSelected2 = ""
							strSelected3 = "selected"
						end if
						%>			
						<option value="1" <%=strSelected1%>>Et av ordene</option>
						<option value="2" <%=strSelected2%>>Alle ordene</option>
						<option value="3" <%=strSelected3%>>Frase</option>			
					</select>
					<!--<input type="submit" value="Søk" ID="Submit1" NAME="Submit1">--><br>
					<DIV>(Tast inn ordet du &oslash;nsker &aring; s&oslash;ke p&aring;. 
					Dersom du vil s&oslash;ke p&aring; ord som innholder s&oslash;kefrasen, bruk asterisk (*) p&aring; slutten.)
					</DIV>
				</DIV>
				<%
				if (IsObject(oRSCV)) then
					if (not oRSCV.EOF) then
				%>
				<div class="contentHead1">
					<h2>Søke resultat</h2>
				</div>			
				<div class="content">
							<div class="listing">
								<TABLE width="96%">
									<col width="70%">
									<col width="15%">
									<col width="15%">								
									<TR>
										<TH>Fil</TH>
										<TH>St&oslash;rrelse</TH>
										<TH>Endret</TH>
									</TR>
									<%
									Dim fsobj
									set fsobj = Server.CreateObject("Scripting.FileSystemObject")
									
									while (not oRSCV.EOF)
									%>
									<TR>
										<TD class="left"><a href="<%=replace(oRSCV.fields("path").value,fsobj.GetParentFolderName(oRSCV.fields("Directory").value)&"\",Application("ConsultantFileRoot"))%>" target="_new"><%=oRSCV.fields("filename").value%></a></TD>
										<TD class="right"><%=ToKB(oRSCV.fields("Size").value)%></TD>
										<TD class="right"><%=oRSCV.fields("write").value%></TD>						
									</TR>
									<%
									oRSCV.movenext
									wend
									set fsobj = nothing
									%>
								</TABLE>
							</div>
							<%
						end if
						oRSCV.close
						set oRSCV = nothing
						oCon.Close
						set oCon = nothing
					end if
					%>
				</div>
			</div>
		</Form>
	</BODY>
</HTML>
<%

function ToKB(nSize)
	dim nLarge
	nLarge = FormatNumber(clng(nSize)/1000,2)
	ToKB = nLarge & " KB" 
end function
%>