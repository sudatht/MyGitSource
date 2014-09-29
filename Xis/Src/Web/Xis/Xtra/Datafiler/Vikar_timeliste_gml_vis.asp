<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<%
	If (HasUserRight(ACCESS_ADMIN, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	
	dim Conn
	dim strSQL	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<!--#INCLUDE FILE="../includes/Library.inc"-->
		<title></title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
		<%
		' parameter og variabler
		' Connect to database
		Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

		Fradato2 = session("Fradato2")
		Tildato2 = session("Tildato2")

		strSQL = "select distinct T.VikarID, Navn=(V.Etternavn + ', ' + V.Fornavn), T.OppdragID, T.FirmaID, F.Firma, VIKAR_ANSATTNUMMER.Ansattnummer " &_
			"from DAGSLISTE_VIKAR T, VIKAR V LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON V.Vikarid = VIKAR_ANSATTNUMMER.Vikarid, FIRMA F " &_
			"where T.Dato < " & dbDate(Tildato2) &_
			" and T.Dato >= " & dbDate(Fradato2) &_
			" and T.VikarID = V.VikarID" &_
			" and T.FirmaID = F.FirmaID" &_
			" and T.TimelisteVikarStatus = 6 "&_
			" order by V.Navn "

		Set rsT = GetFirehoseRS(strSQL, Conn)
		If hasRows(rsT) = false Then
			AddErrorMessage("Ingen gamle timelister funnet.")
			call RenderErrorMessage()
		End If

							' Vis data
							%>
							<div class="contentHead1">
								<h1>Gamle timelister</h1>
							</div>
							<div class="content">
								<div class="listing">
									<table cellpadding='0' cellspacing='0'>
										<tr>
											<th>Ansattnummer</th>
											<th>Navn</th>
											<th>OppdragID</th>
											<th>kontaktID</th>
											<th>Kontakt</th>
										</tr>
									<% 
									do while not rsT.EOF
										%>
										<tr>
											<td><% =rsT("ansattnummer") %>
											<% 
											href = "Vikar_timeliste_vis_gml.asp" &_
											"?OppdragID=" & rsT("OppdragID") &_
											"&Ansattnr=" & rsT("ansattnummer") &_
											"&FirmaID=" & rsT("FirmaID") &_
											"&fra=liste" &_
											"&Fradato2=" & session("fradato2") &_
											"&Tildato2=" & session("tildato2") 
											%>
											</td>
											<td>
												<%if (rsT("ansattnummer") <> "") then%>
												<a href="<% =href %>" ><% =rsT("Navn") %></a>
												<%
												else
												%>
												<% =rsT("Navn") %>
												<%end if%>
											</td>
											<td><% =rsT("OppdragID") %></td>
											<td><% =rsT("FirmaID") %></td>
											<td><% =rsT("Firma") %></td>
										</tr>
										<% 
									rsT.MoveNext
								loop 
							rsT.Close
							Set rsT = Nothing 
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