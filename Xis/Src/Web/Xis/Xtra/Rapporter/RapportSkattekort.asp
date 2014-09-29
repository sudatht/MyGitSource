<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<%
	If (HasUserRight(ACCESS_REPORT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	'Declaring varibles
	dim strSQL	
	Dim Conn				'As ADODB.Connection
	Dim strSkattekortSQL	'As String
	Dim rsSkattekort		'As Recordset
	Dim iTeller				'As Integer
	Dim iNummer				'As Integer

	'Initialzing variables
	iTeller = 0

	' Open database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

	strSkattekortSQL = "SELECT DISTINCT " & _
		"VIKAR.VikarID, " & _
		"VIKAR.Fornavn AS Vikarfornavn, " & _
		"VIKAR.Etternavn AS Vikaretternavn, " & _
		"VIKAR.Epost, " & _
		"VIKAR.MottattSkattekort, " & _
		"VIKAR_ANSATTNUMMER.ansattnummer, " & _
		"MEDARBEIDER.Fornavn AS medarbeiderfornavn, " & _
		"MEDARBEIDER.Etternavn AS medarbeideretternavn " & _
		"FROM VIKAR " & _
		"LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " & _
		"LEFT OUTER JOIN MEDARBEIDER ON VIKAR.AnsMedId = MEDARBEIDER.Medid " & _
		"INNER JOIN VIKAR_UKELISTE ON VIKAR.Vikarid = VIKAR_UKELISTE.Vikarid " & _
		"WHERE VIKAR.Statusid = '3' " & _
		"AND VIKAR.typeID = 1" & _
		"AND ( " & _
		"( VIKAR.MottattSkattekort IS NULL ) " & _
		"OR ( VIKAR.MottattSkattekort = '-1' ) " & _
		"OR ( VIKAR.VikarId NOT IN (  " & _
			"SELECT " & _
			"VIKAR.Vikarid " & _
			"FROM VIKAR " & _
			"WHERE VIKAR.MottattSkattekort = YEAR(GETDATE()) " & _
		"))) " & _
		"ORDER BY VIKAR.Etternavn ASC "

	Set rsSkattekort = GetFireHoseRS(strSkattekortSQL, Conn)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<title>Manglende skattekort</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Manglende skattekort</h1>
			</div>
			<div class="content">
				<p>Rapporten viser kun vikarer registrert som lønnsmotakere.</p>
				<div class="listing">
					<table>
						<tr>
							<th>Ansattnr</th>
							<th>Vikar</th>
							<th>E-post</th>
							<th>Ansvarlig</th>
						</tr>
						<%
						Do While Not rsSkattekort.EOF
							iNummer = iNummer + 1
							%>
							<tr>
								<td class="right">
									<%
									If (rsSkattekort("ansattnummer").Value <> "" ) Then
										Response.Write rsSkattekort("ansattnummer").Value
									Else
										Response.Write "---"
									End If
									%>
								</td>
								<td>
									<%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "vikarvis.asp?VikarID=" & rsSkattekort("vikarid").Value, rsSkattekort("vikaretternavn").Value & ", " & rsSkattekort("vikarfornavn").Value, "Vis vikar " & rsSkattekort("vikaretternavn").Value & ", " & rsSkattekort("vikarfornavn").Value )%>
								</td>
								<td>
									<a href="mailto:<%=rsSkattekort("Epost").Value%>">
										<font class="groenn"><%=rsSkattekort("Epost").Value%>
									</a>
								</td>
								<td>
									<%
									If (rsSkattekort("medarbeideretternavn").Value <> "" And rsSkattekort("medarbeiderfornavn").Value <> "") Then
										Response.Write rsSkattekort("medarbeideretternavn").Value & ", " & rsSkattekort("medarbeiderfornavn").Value
									Else
										Response.Write "Ikke tildelt"
									End If
									%>
								</td>
							</tr>
							<%
							iTeller = iTeller + 1
							rsSkattekort.MoveNext
						Loop
						rsSkattekort.Close()
						set rsSkattekort = nothing
						%>
					</table>
				</div>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>