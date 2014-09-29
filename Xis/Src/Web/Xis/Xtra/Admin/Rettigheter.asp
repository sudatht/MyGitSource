<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<% 

Dim strBrukerProfil		'As	String
Dim Conn				'As ADODB.Connection
Dim bSkattekort			'As Boolean(1/0)
'Dim bKundegrupper		'As Boolean(1/0)
Dim bFAgreementHanler		'As Boolean (1/0)
Dim bDeletePerson		'As Boolean (1/0)
Dim bQuestback			'As Boolean (1/0)
Dim iBrukerId			'As Integer
Dim strBrukerNavn		'As String
Dim strSQL				'As String
Dim rsProfil			'As Recordset
Dim strFulltNavn		'As String
Dim iAntallRettigheter	'As Integer
Dim iTeller				'As Integer
Dim strSelectNavn		'As Integer
Dim iSelectIndex		'As Integer

If  HasUserRight(ACCESS_ADMIN, RIGHT_SUPER) Then 
	%>
	<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
	<html>
		<head>
			<title>Tilgang</title>
			<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
			<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
			<style type="text/css">
				SELECT {width:auto;}
			</style>
		</head>
		<body>
			<div class="pageContainer" id="pageContainer">
			<%
			'--------------------------------------------------------------------------------------------------
			' Connect to database
			'--------------------------------------------------------------------------------------------------

			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
			Conn.CommandTimeOut = Session("xtra_CommandTimeout")
			Conn.Open Session("xtra_ConnectionString"),Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

			'--------------------------------------------------------------------------------------------------
			' check parameters
			'--------------------------------------------------------------------------------------------------

			iBrukerId = Request.Querystring("ID")
			strBrukerNavn = Request.Querystring("BrukerID")

			'--------------------------------------------------------------------------------------------------
			' sql for profile
			'--------------------------------------------------------------------------------------------------
			strSQL = "SELECT " & _
				"[profil], " & _
				"[Navn], " & _
				"[SkattekortEndringer], " & _
				"[KundeGruppeEndringer], " & _
				"[FAgreementHandler], " & _
				"[DeletePersonInfo], " & _
				"[Questback] " & _
				"FROM [Bruker] " & _
				"WHERE [ID] = '" & iBrukerId & "' "

			Set rsProfil = conn.Execute(strSQL)
			
			If Not rsProfil.EOF Then

				strBrukerProfil = rsProfil("Profil").Value
				strFulltNavn = rsProfil("Navn").Value
				bSkattekort = rsProfil("skattekortendringer").Value
				'bKundegrupper = rsProfil("KundeGruppeEndringer").Value
				bFAgreementHanler = rsProfil("FAgreementHandler").Value
				bDeletePerson = rsProfil("DeletePersonInfo").Value
				bQuestback = rsProfil("Questback").Value
				rsProfil.Close
				set rsProfil = nothing

				'--------------------------------------------------------------------------------------------------
				' display rights
				'--------------------------------------------------------------------------------------------------

				iAntallRettigheter = Len(strBrukerProfil)

				%>
				<div class="contentHead1">
					<h1>Rettigheter for <%=strFulltNavn%></h1>
				</div>
				<div class="content">

					<form action="RettighetDB.asp?ID=<%=iBrukerId%>&BrukerID=<%=strBrukerNavn%>&Navn=<%=strFulltNavn%>" method="post">
						<input type="hidden" name="HowMany" value="<%=iAntallRettigheter%>">
						
						<div class="listing">
						<table style="width:380px;">
							<tr>
								<th>Hotlist</th>
								<th>Kunde</th>
								<th>Vikar</th>
								<th>Rapport</th>
								<th>Oppdrag</th>
								<th>Admin</th>
							</tr>
							<tr>
							<%
							For iTeller = 1 To iAntallRettigheter 
								strSelectNavn = "HovedMeny" & iTeller 
								iSelectIndex = Mid(strBrukerProfil, iTeller, 1)
								%>
								<td>
									<select name="<%=strSelectNavn%>">
										<OPTION VALUE="0" <%If iSelectIndex = 0 Then Response.write "selected"%>>Ingen</option>
										<OPTION VALUE="1" <%If iSelectIndex = 1 Then Response.write "selected"%>>Lese</option>
										<OPTION VALUE="2" <%If iSelectIndex = 2 Then Response.write "selected"%>>Skrive</option>
										<OPTION VALUE="3" <%If iSelectIndex = 3 Then Response.write "selected"%>>Admin</option>
										<%If (iTeller = 6) Then%>
											<OPTION VALUE="4" <%If iSelectIndex = 4 Then Response.write "selected"%>>Super</option>
										<%End If%>
									</select>
								</td>
							<%Next%>
							</tr>
							</table>
							</div>
							<div style="width:380px; padding-bottom:5px; padding-top:10px; text-align:right;">
								Bruker kan redigere skattekort: <input class="checkbox" type="checkbox" id="skattekort" name="skattekort" value="1" <%If (bSkattekort) Then Response.Write "checked"%>>
							</div>
							<div style="width:380px; padding-bottom:5px; text-align:right;">
								Bruker kan vedlikeholde rammeavtaler: <input class="checkbox" type="checkbox" id="fagreement" name="fagreement" value="1" <%If (bFAgreementHanler) Then Response.Write "checked"%>>			
							</div>
							<div style="width:380px; padding-bottom:5px; text-align:right;">
								Tillat bruker å slette vikarinfo på enkeltvikarer: <input class="checkbox" type="checkbox" id="deleteperson" name="deleteperson" value="1" <%If (bDeletePerson) Then Response.Write "checked"%>>			
							</div>
							<div style="width:380px; padding-bottom:5px; text-align:right;">
								Tilgang til Questback-rapporter: <input class="checkbox" type="checkbox" id="questback" name="questback" value="1" <%If (bQuestback) Then Response.Write "checked"%>>								
							</div>
						<input type="submit" value="LAGRE" id="lagrerettighet" name="lagrerettighet">
						<br>&nbsp;
					</form>
				</div>
				<% 
				Else 'ingen linjer 
					rsProfil.Close
					set rsProfil = nothing
				End If
				%>
			</div>
		</body>
	</html>
<%
End If
%>