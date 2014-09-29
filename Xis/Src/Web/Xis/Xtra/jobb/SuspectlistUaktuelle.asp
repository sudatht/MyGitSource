<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<!--#INCLUDE FILE="../includes/SuperOffice.Integration.Contact.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<%
	call Response.Redirect("/xtra/IngenTilgang.asp")

	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim Conn
	dim strSQL
	dim rsVikar

	if  request("soek") = "ja" Then

		if request("dbxMedarbeider")="0"   then
		
			strSQL = "SELECT DISTINCT V.suspectId, V.regdato, V.Etternavn, V.Fornavn, Adresse, Postnr, Poststed,V.AnsMedID, V.Telefon, V.MobilTlf, V.EPost "&_
				" FROM  V_SUSPECT V, V_SUSPECT_ADRESSE A, V_SUSPECT_AVDKONTOR K " &_
				" WHERE V.suspectID *= A.AdresseRelID "&_		
				" AND (A.AdresseType = 2 OR A.AdresseType IS NULL) "&_	
				" AND V.Uaktuell = 1 "&_
				" AND V.slettes = 0 "&_
				" AND V.suspectID = K.suspectID  "&_
				" AND K.overfoert = 0 "&_			
				" ORDER BY  V.suspectID DESC"
				
		elseif request("dbxMedarbeider")="1"   then
		
			strSQL = "SELECT DISTINCT V.suspectId, V.regdato, V.Etternavn, V.Fornavn, Adresse, Postnr, Poststed,V.AnsMedID, V.Telefon, V.MobilTlf, V.EPost "&_
				" FROM  V_SUSPECT V, V_SUSPECT_ADRESSE A, V_SUSPECT_AVDKONTOR K " &_
				" WHERE V.suspectID *= A.AdresseRelID "&_		
				" AND (A.AdresseType = 2 OR A.AdresseType IS NULL) "&_	
				" AND V.Uaktuell = 1 "&_
				" AND V.slettes = 0 "&_
				" AND V.suspectID = K.suspectID  "&_
				" AND K.overfoert = 0 "&_
				" AND V.AnsMedID IS NULL "&_			
				" ORDER BY  V.suspectID DESC"
				
		else
			strSQL = "SELECT DISTINCT V.suspectId, V.regdato, V.Etternavn, V.Fornavn, Adresse, Postnr, Poststed,V.AnsMedID, V.Telefon, V.MobilTlf, V.EPost "&_
				" FROM  V_SUSPECT V, V_SUSPECT_ADRESSE A,  V_SUSPECT_AVDKONTOR K " &_
				" WHERE V.suspectID *= A.AdresseRelID "&_		
				" AND (A.AdresseType = 2 OR A.AdresseType IS NULL) "&_	
				" AND V.AnsMedID =" & request("dbxMedarbeider") &_	
				" AND V.Uaktuell = 1 "&_
				" AND V.slettes = 0 "&_
				" AND V.suspectID = K.suspectID  "&_
				" AND K.overfoert = 0 "&_			
				" ORDER BY  V.suspectID DESC"
		end If
		Set Conn = GetConnection(GetConnectionstring(XIS, ""))		
		Set rsVikar = GetFirehoseRS( strSql, Conn )		
	end if

' Visning av data
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"  "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<title>Jobbsøknader fra web - uaktuelle</title>
		<script language="javaScript" type="text/javascript">

			//globale variabler
			var i=0;

			function focused(f)
			{
				i=f.substring(3,6);
				i=parseInt(i);
			}
			//###############kode for å sette fokus ved onLoad()################
			function fokus()
			{
				if(document.all('lnk0')){
					document.all('lnk0').focus();
				}
			}
		</script>		
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Jobbsøknader fra web - uaktuelle </h1>
			</div>
			<div class="content">
				<form action="SuspectListUaktuelle.asp?soek=ja" METHOD="post">
					<table>
						<tr>
							<td><INPUT TYPE="submit" VALUE="Sorter på ansvarlig medarbeider" ID="Submit1" NAME="Submit1"></td>
							<td>
								<SELECT NAME="dbxMedarbeider" ID="Select1">     
									<option VALUE="0">(Alle)</option>
									<option VALUE="1">(Uten ansvarlig)</option>
									<%
									' Get ansvarlig medarbeider
									strSQL = "SELECT MedId, Fornavn, Etternavn FROM medarbeider ORDER BY etternavn"
									Set rsMedarbeider = GetFirehoseRS(strSQL, Conn)
								 
									Do Until rsMedarbeider.EOF 
										If rsMedarbeider("MedID") = strAnsMed Then
											sel = " SELECTED"
										Else
											sel = ""
										End If 
										strName = rsMedarbeider("Etternavn") & "  " & rsMedarbeider("Fornavn") %>
										<option VALUE="<% =rsMedarbeider("MedID")%>" <% =sel %>><% =strName %></option>
										<%   
										rsMedarbeider.MoveNext
									Loop
									rsMedarbeider.Close
									Set rsMedarbeider = Nothing
									%>
								</select>
	 						</td>
						</tr>
					</table>
				</form>
				<%
				if  request("soek") = "ja" Then
					%>
					<div class="listing">
						<table>
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
							i=0  
							if(HasRows(rsVikar)) then
								Do Until rsVikar.EOF
									nummer = nummer + 1
									strFullName   = rsVikar("Etternavn") & " " & rsVikar("Fornavn")   
									strFullAdress = rsVikar("Adresse") & " " & rsVikar("Postnr") & " " & rsVikar("Poststed") 
									%>
									<tr>
										<td><%=rsVikar("suspectId") %></td>
										<td><%=rsVikar("regDato") %></td>
										<td><A ID=lnk<%=i%> onFocus="focused(this.id)" Href="SuspectVisUaktuelle.asp?suspectId=<%=rsVikar("suspectId") %>"><font ID=vikar<%=nummer%>><%= strFullName %></font></A></td>
										<td><%=strFullAdress %></td>
										<td><%=rsVikar("Telefon") %></td>
										<td><%=rsVikar("MobilTlf") %></td>
										<td><A Href="mailto:<%=rsVikar("EPost") %>"><font ID=epost<%=nummer%>><%=rsVikar("EPost") %></font></a></td>
										<%
										i=i+1
										%>
									</tr>
									<%   
									rsVikar.MoveNext
								Loop
								rsVikar.Close
							end if
							set rsVikar = Nothing
							%>
					</table>
				</div>
				<%
			end If 'hvis soek = ja 
			%>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>