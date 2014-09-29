<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.Settings.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Renderfunctions.inc"-->
<%

	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	
	brukerID	= Session("BrukerID")

	dim strFornavn
	dim strEtternavn
	dim strFodtselsdato
	dim strAdresse
	dim strPostAdresse
	dim oConn
	dim oCommand
	dim rsDuplicates
	dim strHeading
	dim rsVikar
	dim blnValidated

	if (request("hdnPosted") <> "") then

		blnValidated = false

		strFornavn		= trim(Request("txtFornavn"))
		strEtternavn		= trim(Request("txtEtternavn"))
		strFodtselsdato		= trim(Request("dtFodselsDato"))

		if ((len(strFornavn) > 0) AND (len(strEtternavn) > 0))  then
			strHeading = "Mulig duplikat oppføring"
			blnValidated = true
			
			' Open connection
			Set oConn = GetConnection(GetConnectionstring(XIS, ""))	

			'Retrieve all consultants that are possible duplicates
			set rsDuplicates = GetFirehoseRS("exec sp_intra_getPossibleDuplicates '"  & strFornavn & "', '" & strEtternavn & "', " & dbdate(strFodtselsdato), oConn)

			if (HasRows(rsDuplicates) = false) then
				CloseConnection(oConn)
				set oConn = nothing
				'There are no possible duplicates, redirect to tranfer suspect page
				Response.Redirect "vikarNy.asp?txtForNavn=" & strFornavn & "&txtEtterNavn=" & strEtternavn & "&dtFodselsDato=" & strFodtselsdato
			end if
		else
			strHeading		= "Ny vikar"
		end if
	else
		blnPosted			= false
		strHeading			= "Ny vikar"
		strFornavn			= ""
		strEtternavn		= ""
		strFodtselsdato		= ""
	end if
%>
<html>
	<head>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<title><%=strHeading%></title>
		<script language="javascript" src="Js/javascript.js" type="text/javascript"></script>
		<script language='javascript' src='js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='js/navigation.js' id='navigationScripts'></script>			
		<script language="javaScript">
			function shortKey(e) 
			{
				var keyChar = String.fromCharCode(event.keyCode);
				var modKey  = event.ctrlKey;
				var modKey2 = event.shiftKey;

				if (event.keyCode == 13)
				{					
					parent.frames[funcFrameIndex].location=("/xtra/VikarDuplikatSoek.asp");					
				}
				//linker i submeny
				<% 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>		
					if (modKey && modKey2 && keyChar=="S")
					{
						parent.frames[funcFrameIndex].location=("/xtra/vikarSoek.asp");
					}
					<% 
				End If 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
					%>
					if (modKey && modKey2 && keyChar=="Y")
					{
						parent.frames[funcFrameIndex].location=("/xtra/VikarDuplikatSoek.asp");
					}
					<% 
				End If 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>
					if (modKey && modKey2 && keyChar=="W")
					{
						parent.frames[funcFrameIndex].location=("/xtra/jobb/SuspectList.asp");
					}
					<% 
				End If 
				%>
			}
			// her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;
		</script>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
		<div class="contentHead1">
			<H1><%=strHeading%></H1>
		</div>
		<div class="content">
			<%
			if (blnValidated = true) then
				%>
				<form Name="suspekt" action="VikarNy.asp" METHOD="post" ID="Form1">
					<input type='hidden' value="<%=strFornavn%>" name="txtFornavn" ID="Hidden1">
					<input type='hidden' value="<%=strEtternavn%>" name="txtEtternavn" ID="Hidden2">
					<input type='hidden' value="<%=strFodtselsdato%>" name="dtFodselsDato" ID="Hidden3">
					<div class="listing">
						<h2>Du pr&oslash;ver &aring; opprette f&oslash;lgende vikar</h2>
						<table cellspacing='1' cellpadding='0' width="70%" ID="Table1">
							<col width="20%">
							<col width="80%">
							<tr>
								<TH>F&oslash;dselsdato</TH>
								<TH>Navn</TH>
							</tr>
							<tr>
								<TD><%=strFodtselsdato%>&nbsp;</TD>
								<TD><%=strFornavn%>&nbsp;<%=strEtternavn%></TD>
							</tr>
						</table>
						<h2>F&oslash;lgende mulige duplikater finnes</h2>
						<table cellspacing='1' cellpadding='0' width="70%" ID="Table2">				
							<tr>
								<th>Ansattnr</th>
								<th>f&oslash;dselsdato</th>
								<th>vikar</th>
								<th>Adresse</th>
								<th>Postadresse</th>
								<th>Type</th>
								<th>Opprettet</th>
							</tr>
							<%
							while (not rsDuplicates.EOF)
								%>
								<tr>
									<td><%=rsDuplicates("AnsattNummer").value%>&nbsp;</td>
									<td><%=rsDuplicates("foedselsDato").value%>&nbsp;</td>
									<td><a href='vikarny.asp?vikarid=<%=rsDuplicates("vikarid").value%>'><%=rsDuplicates("etternavn").value & ", " & rsDuplicates("fornavn").value%></a></td>
									<td><%=rsDuplicates("Adresse").value%>&nbsp;</td>
									<td><%=rsDuplicates("Postnr").value & " " & rsDuplicates("poststed").value%>&nbsp;</td>
									<td><%=rsDuplicates("vikarType").value%>&nbsp;</td>
									<td><%=rsDuplicates("regdato").value%>&nbsp;</td>
								</tr>
								<%
								rsDuplicates.movenext
							wend
							rsDuplicates.close
							set rsDuplicates = nothing
							%>
						</table>
					</div>
					<input type='submit' value="Opprett som ny" name="submit" ID="Submit1">
					<br>&nbsp;
				</form>
				<%
			else
				%>
				<form Name="frmDuplicateCheck" action="VikarDuplikatSoek.asp" METHOD="post" ID="Form2">
					<input type="hidden" id="hdnPosted" name="hdnPosted" value="true">
					<table ID="Table3">
					<tr>
						<td>Fornavn:&nbsp;</td>
						<td><input class="mandatory" type='text' value="<%=strFornavn%>" id="txtFornavn" name="txtFornavn" SIZE="20" MAXLENGTH="50"></td>
						<td>Etternavn:&nbsp;</td>
						<td><input class="mandatory" type='text' value="<%=strEtternavn%>" id="txtEtternavn" name="txtEtternavn" SIZE="25" MAXLENGTH="50"></td>
						<td>F&oslash;dselsdato:&nbsp;</td>
						<td><input type='text' value="<%=strFodtselsdato%>" id="dtFodselsDato" name="dtFodselsDato"  SIZE="8" MAXLENGTH="10" ONBLUR="dateCheck(this.form, this.name)"></TH>
					</tr>
				</table>
				<input type='submit' value="til duplikatsjekk" name="submit" ID="Submit2">
				<br>&nbsp;
				<%
			end if
			%>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(oConn)
set oConn = nothing
%>