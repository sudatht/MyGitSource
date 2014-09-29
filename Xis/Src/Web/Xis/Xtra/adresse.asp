<%@ LANGUAGE="VBSCRIPT"%>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<% 
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim strSQL
	dim conn
	dim strID
	dim strName
	dim lAdrRelasjon
	dim strHeading

	'Open database connection
	Set conn = GetConnection(GetConnectionstring(XIS, ""))	

	lAdrRelasjon = Request("Relasjon")
	If lAdrRelasjon = 2 Then '2 = Vikar, 1 = firma (defunct)

		strSQL = "SELECT VikarID, Etternavn, Fornavn FROM Vikar WHERE VikarID = " & Request("ID")
		Set rsVikar = GetFirehoseRS(strSQL, Conn)

		strID       = rsVikar("VikarID")
		strName     = rsVikar("Fornavn") & " " & rsVikar("Etternavn")

		' Close and release recordset
		rsVikar.Close
		Set rsVikar = Nothing
	Else
		AddErrorMessage("Ugyldig addresserelasjon spesifisert.")
		call RenderErrorMessage()
	End If

	' Create page heding 
	strHeading = "Adresse for " & strName

	' Existing adress ?
	If Request.QueryString("AdrID") <> "" Then
		strSQL = "SELECT A.Adresse, A.Postnr, A.Poststed, T.AdrTypeID, T.AdresseType FROM ADRESSE A, H_ADRESSE_TYPE T " & _
                " WHERE A.AdrID = " & Request.QueryString("AdrID") & _
                " AND A.Adressetype = T.AdrtypeID "

		Set rsAdresse = GetFirehoseRS(strSQL, Conn)

		lAdrtype    = rsAdresse("AdrTypeID")
		strAdress   = rsAdresse("Adresse")
		strPostNr   = rsAdresse("Postnr")
		strPoststed = rsAdresse("Poststed")
			    
			' Close and release recordset
		rsAdresse.Close
		Set rsAdresse = Nothing
	End If

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
		<title>Adresse</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
	<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1><%=strHeading%></h1>
			</div>
			<div class="content">
				<form ACTION="adresseDB.asp" METHOD="POST">
					<input TYPE="HIDDEN" NAME="tbxRelasjon" VALUE="<%=Request.QueryString("Relasjon") %>">
					<input TYPE="HIDDEN" NAME="tbxID" VALUE="<%=Request.QueryString("ID") %>">
					<input TYPE="HIDDEN" NAME="tbxAdrID" VALUE="<%=Request.QueryString("AdrID")%>">
					<table>
						<tr>
							<td>Type:</td>
							<td colspan="2">
								<select NAME="dbxAdrType">
									<% 
									strSQL = "SELECT AdrtypeID, Adressetype FROM H_ADRESSE_TYPE WHERE AdrTypeID > 1 ORDER BY AdrtypeID"
									Set rsAdrType = GetFirehoseRS(strSQL, Conn)

									Do Until rsAdrtype.EOF 

										If rsAdrtype("adrtypeID")  = lAdrtype  Then
											strSelected = " SELECTED"
										Else
											strSelected = ""
										End If
										%>
    									<option VALUE="<%=rsAdrtype("AdrtypeID")%>" <%=strSelected%>> <%=rsAdrtype("Adressetype")%>
										<% 
										rsAdrtype.MoveNext
									Loop
									' Close and release recordset
									rsAdrtype.Close
									Set rsAdrtype = Nothing
									%>
								</select>
							</td>
						</tr>
						<tr>
							<td>Adresse:</td>
							<td colspan="3"><input NAME="tbxAdress" TYPE="TEXT" SIZE="40" MAXLENGTH="50" Value="<%=strAdress%>"></td>
						</tr>
						<tr>
							<td>Postnr:</td>
							<td><input NAME="tbxPostnr" TYPE="TEXT" SIZE="5" MAXLENGTH="5" Value="<%=strPostnr%>"></td>
							<td>Poststed:</td>
							<td><input NAME="tbxPoststed" TYPE="TEXT" SIZE="20" MAXLENGTH="50" Value="<%=strPoststed%>"></td>
						</tr>
					</table>
					<table>
						<tr>
							<% 
							If lenb(Request("AdrID")) = 0 Then
								Response.write "<td><INPUT NAME=pbnDataAction TYPE=SUBMIT  VALUE=Lagre>"
							Else
								Response.write "<td><INPUT NAME=pbnDataAction TYPE=SUBMIT  VALUE=Lagre>"
								Response.write "<td><INPUT NAME=pbnDataAction TYPE=SUBMIT  VALUE=Slette>"
							End If
							%>
						</tr>
					</table>
				</form>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(conn)
set conn = nothing
%>