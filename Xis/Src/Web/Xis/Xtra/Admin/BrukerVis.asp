<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<%  
If  HasUserRight(ACCESS_ADMIN, RIGHT_ADMIN) Then 

	' Connect to database
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
	Conn.CommandTimeOut = Session("xtra_CommandTimeout")
	Conn.Open Session("xtra_ConnectionString"),Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

	' CHECK PARAMETERS
	kode = Request.QueryString("kode")
	strID = Request.QueryString("ID")
	strMedarbID = Request.form("medarbID")
	'response.write strMedarbID

	If Request.QueryString("BrukerID") = "" Then
		strBrukerID = Request.Form("BrukerID")
		strNavn2 = Request.Form("Navn")
	Else
		strBrukerID = Request.QueryString("BrukerID")
		strNavn2 = Request.QueryString("Navn")
	End If

	Endre = Request.QueryString("Endre")
	Ny = Request.QueryString("Ny")
	Slett = Request.QueryString("Slett")

	' SQL to UPDATE
	If Endre = "Ja" Then
		strSQL = "Update BRUKER set BrukerID = '" & strBrukerID & "', Navn = '" & strNavn2 &"', medarbID= '"& strMedarbID & "' where ID = " & strID
		conn.Execute(strSQL)
	End If
	' SQL to INSERT
	If Ny = "Ja" Then
		strSQL = "Insert into BRUKER (BrukerID,Navn,Profil,medarbID) values('" & strBrukerID & "', '" & strNavn2 & "','000000','"& strMedarbID & "')"
		conn.Execute(strSQL)
	End If
	' SQL to DELETE
	If Slett = "Ja" Then
		strSQL = "Delete from BRUKER where ID = " & strID
		'Response.write strSQL & "<br>"
		conn.Execute(strSQL)
	End If
	' SQL to find USERS
	strSQL = "select * from BRUKER ORDER BY BrukerID"
	Set rsNavn = Conn.Execute(strSQL)
	' PAGE HEADING
	%>
	<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
	<html>
		<head>
			<title>Rettigheter</title>
			<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
			<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
			<script language="javascript">
				function velgNavn()
				{			
					var ind;
					ind = document.all('dbxMedarbeider').selectedIndex;
					var navn;
					navn = document.forms[0].elements('dbxMedarbeider').options[ind].text;
					medarbID=document.forms[0].elements('dbxMedarbeider').options[ind].value;
					document.forms[0].elements('Navn').value=navn;
					document.forms[0].elements('medarbID').value=medarbID;			
				}
			</script>
		</head>
		<body>
			<div class="pageContainer" id="pageContainer">
				<div class="content">	
					<h1>Brukeradministrasjon</h1>
					<%
					'--------------------------------------------------------------------------------------------------
					' DISPLAYING USERS
					'--------------------------------------------------------------------------------------------------
					'kode = ""  		Viser liste med linker til
					'kode = 1   		Viser liste med endring og ny-knapp
					'kode = 2	   	Du endrer, viser felter for brukernavn og fulltnavn med bruker og lagreknapp
					'kode = 3	   	Ny, viser felter for brukernavn (tomme) og lagreknapp
					'kode = 4		Slett
					'Endre = Ja		Update
					'Ny = Ja		Insert
					If (kode > 1) Then 
						If (Kode = 2) Then
							action = "BrukerVis.asp?Endre=Ja&ID=" & strID
							brukerId = strBrukerID
							navn = strNavn2
						ElseIf (Kode = 3) Then 
							action = "BrukerVis.asp?Ny=Ja"
							brukerID = ""
							navn = ""
						End If	
						%>
						<div class='listing'>
							<form action="<% =action %>" method="post" >
								<table cellspacing='1' cellpadding='0' width="96%">
									<tr>
										<th>BrukerID:</th>
										<td><input type="text" name="BrukerID" value="<% =brukerID %>"></td>
									</tr>
									<tr>
										<th>Fullt navn:</th>
										<td><input type=text size=20 name=Navn value=<% =navn %> ></td>
									</tr>
									<tr>
										<th>Velg fra medarbeidere:</th>
										<td>
											<SELECT name="dbxMedarbeider" onChange="velgNavn()">
											<% 
												' Get ansvarlig medarbeider
												Set rsMedarbeider = Conn.Execute("Select MedId, Fornavn, Etternavn from medarbeider")

												Do Until rsMedarbeider.EOF
													If rsMedarbeider("MedID") = lAnsMedID Then
														strValueSelected = rsMedarbeider("MedID") & " SELECTED"
													Else
														strValueSelected = rsMedarbeider("MedID")
													End If 
													strName = rsMedarbeider("Fornavn") & " " & rsMedarbeider("Etternavn")  %>
       														<OPTION value="<%=strValueSelected %>"><%=strName %></option>
													<%   
													rsMedarbeider.MoveNext
												Loop 
												rsMedarbeider.Close
											%>
   											</SELECT>
										</td>
									</tr>
								</table>
								<input type="hidden" name="medarbID" value="">
								<input type="submit" value="Lagre">
							</form>	
						</div>
						<% 
					End If 
						%>
						<div class="listing">
							<table cellspacing='1' cellpadding='0' width="96%">
								<tr>
								<% 
								If Kode = 1 Then 
									%>
									<th></th>
									<th>ID</th>
									<th>Navn</th>
									<th></th>
									<%
								Else 
									%>
									<th>ID</th>
									<th>Navn</th>
									<% 
								End If
								%>
								</tr>
								<%
								DO WHILE NOT rsNavn.EOF 
									strNavn = rsNavn("Navn") 		
									If kode = 1 Then   	'vise endrings utgangspunkt (fjerne linker til tilgangsbildet, vise link til endring) 
										%>
										<tr>
											<td><a href="BrukerVis.asp?kode=2&ID=<% =rsNavn("ID") %>&BrukerID=<% =rsNavn("BrukerID") %>&Navn='<% =strNavn %>'" >End</a></td>
											<td><% =rsNavn("BrukerID") %></td>
											<td><% =strNavn %></td>
											<td><a href="BrukerVis.asp?Slett=Ja&ID=<% =rsNavn("ID") %>" >Slett</a></td>
										</tr>
										<% 
									ElseIf kode > 1 Then 
										%>
										<tr>
											<td><% =rsNavn("BrukerID") %></td>
											<td><% =strNavn %></td>
										</tr>
										<% 
									Else 'vanlig visning (linker til tilgangsbildet) 
										%>
										<tr>
											<td><a href="rettigheter.asp?ID=<% =rsNavn("ID") %>&BrukerID=<% =rsNavn("BrukerID") %>&Navn=<% =strNavn %>" TARGET="RIGHT_WINDOW" ><% =rsNavn("BrukerID") %></a></td>
											<td><a href="rettigheter.asp?ID=<% =rsNavn("ID") %>&BrukerID=<% =rsNavn("BrukerID") %>&Navn=<% =strNavn %>" TARGET="RIGHT_WINDOW" ><% =rsNavn("Navn") %></a></td>
										</tr>
										<% 
									End If 
									rsNavn.MoveNext
								loop 
								rsNavn.Close 
								set rsNavn = nothing
								%>
							</table>
						</div>
					<% 
					If kode = "" Then    'vis endreknappen 
					%>
					<FORM ACTION="BrukerVis.asp?kode=1" method="post" >
						<input type="submit" value="Endre">
					</form>
					<% 
					End If
					If kode = 1 Then 		'vise ny knapp 
					%>
					<FORM ACTION="BrukerVis.asp?kode=3" method="post" >
						<input type="submit" value="Ny">
					</form>
					<% 
					End If 
					%>
					<br>
				</div>
			</div>
		</body>
	</html>
	<% 
End If 
%>