<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Renderfunctions.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	strAnsMed = Session("medarbID")

	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

	if  request("soek") = "ja" Then

		strAnsMed = request.form("dbxMedarbeider")

		if request("dbxMedarbeider") = "0"   then

			strSql = "select distinct V.suspectId, V.regdato, V.Etternavn, V.Fornavn, Adresse, Postnr, Poststed,V.AnsMedID, V.Telefon, V.MobilTlf, V.EPost "&_
				" from V_SUSPECT V, V_SUSPECT_ADRESSE A" &_
				" where V.suspectID = A.AdresseRelID "&_
				" and A.AdresseType = 2 "&_				
				" and V.CVFerdig = 1 "&_
				" and V.slettes  = 0 "&_
				" and v.overfort= 0 "&_
				" order by V.suspectID desc"

		elseif request("dbxMedarbeider") = "1"   then

			strSql = "select distinct V.suspectId, V.regdato, V.Etternavn, V.Fornavn, Adresse, Postnr, Poststed,V.AnsMedID, V.Telefon, V.MobilTlf, V.EPost "&_
				" from V_SUSPECT V, V_SUSPECT_ADRESSE A " &_
				" where V.suspectID = A.AdresseRelID "&_
				" and A.AdresseType = 2 "&_
				" and V.CVFerdig = 1 "&_
				" and V.slettes  = 0 "&_
				" and v.overfort= 0 "&_
				" and V.AnsMedID is Null "&_
				" order by V.suspectID desc"
		else
			
			strSql = "select distinct V.suspectId, V.regdato, V.Etternavn, V.Fornavn, Adresse, Postnr, Poststed,V.AnsMedID, V.Telefon, V.MobilTlf, V.EPost "&_
				" from V_SUSPECT V, V_SUSPECT_ADRESSE A " &_
				" where V.suspectID = A.AdresseRelID "&_
				" and A.AdresseType = 2 "&_
				" and V.AnsMedID =" & request("dbxMedarbeider") &_
				" and V.CVFerdig = 1 "&_
				" and V.slettes  = 0 "&_
				" and v.overfort = 0 "&_
				" order by V.suspectID desc"
		end If

		set rsVikar = GetFirehoseRS(strSql, Conn)
	else
		if strAnsMed > "1"   then
			strSql = "select distinct V.suspectId, V.regdato, V.Etternavn, V.Fornavn, Adresse, Postnr, Poststed,V.AnsMedID, V.Telefon, V.MobilTlf, V.EPost " &_
				" from V_SUSPECT V, V_SUSPECT_ADRESSE A " &_
				" where V.suspectID = A.AdresseRelID " &_
				" and A.AdresseType = 2 " &_
				" and V.AnsMedID =" & strAnsMed & _
				" and V.slettes  = 0 " &_
				" and v.overfort = 0 " &_
				" order by V.suspectID desc"

			set rsVikar = GetFirehoseRS(strSql, Conn)
		end if
	end if

' Visning av data
%>
<html>
	<head>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<title>Resultat</title>
		<script language='javascript' src='/xtra/js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='/xtra/js/navigation.js' id='navigationScripts'></script>
		<script language="javaScript" type="text/javascript">
			function shortKey(e) 
			{
				var keyChar = String.fromCharCode(event.keyCode);
				var modKey  = event.ctrlKey;
				var modKey2 = event.shiftKey;

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
			//her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;
		</script>		
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<a id="Top"></a>
			<div class="contentHead1">
				<h1>Jobbsøkere fra web</h1>
			</div>
			<div class="content">
				<form action="SuspectList.asp?soek=ja" METHOD="post">
					<table>
						<tr>
							<td>
								<select NAME="dbxMedarbeider" onChange="submit()" onkeydown="typeAhead()">
									<option <%if clng(strAnsMed) = 0 then response.write "selected" %>  VALUE="0"> ALLE </option>
									<option <%if clng(strAnsMed) = 1 then response.write "selected" %> VALUE="1"> NYE, uten ansvarlig</option>
									<%
									' Get ansvarlig medarbeider
									response.write GetCoWorkersAsOptionList(strAnsMed)
									%>
								</select>
							</td>	
						</tr>
					</table>
				</form>
				<div class="listing">
					<table cellpadding="0" cellspacing="1">
						<tr>
							<TH>Nr</TH>
							<TH>Dato</TH>
							<TH>Navn</TH>
							<TH>Hjemmeadresse</TH>
							<TH>Telefon</TH>
							<TH>Mobil</TH>
							<TH>EPost</TH>
						</tr>
						<%
						i = 0
						Do Until (rsVikar.EOF)
							nummer = nummer + 1
							strFullName	  = rsVikar("Etternavn") & " " & rsVikar("Fornavn")
							strFullAdress = rsVikar("Adresse") & " " & rsVikar("Postnr") & " " & rsVikar("Poststed")
							%>
							<TR>
								<TD><%=rsVikar("suspectId")%></TD>
								<TD><%=rsVikar("regDato") %></td>
								<TD><a Href="SuspectVis.asp?suspectId=<%=rsVikar("suspectId") %>"><font ID=vikar<%=nummer%>><%= strFullName %></font></a></TD>
								<TD><%=strFullAdress %>&nbsp;</td>
								<TD><%=rsVikar("Telefon") %>&nbsp;</td>
								<TD><%=rsVikar("MobilTlf") %>&nbsp;</td>
								<TD><A Href="mailto:<%=rsVikar("EPost") %>"><font ID=epost<%=nummer%>><%=rsVikar("EPost") %></font></a></td>
							</TR>
							<%
							i = i + 1
							rsVikar.MoveNext
						Loop
						rsVikar.Close
						set rsVikar = Nothing
						%>
						</tr>
					</table>
					<a href="#top"><img src="../Images/icon_GoToTop.gif" alt="Til toppen">Til toppen<a>
				</div>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(objCon)
set objCon = nothing
%>