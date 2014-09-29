<%@ LANGUAGE="VBSCRIPT" %>
<% 
option explicit 
Response.Expires = 0
%>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<!--#INCLUDE FILE="../includes/Xis.Reports.Utils.inc"-->
<%
	If (HasUserRight(ACCESS_ADMIN, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	
	'Deklarer variabler
	Dim Conn
	Dim lPeriode
	Dim i
	Dim avd
	Dim avdnr
	Dim DDate
	Dim lYear
	Dim lMonth
	Dim lDay
	Dim strDay
	Dim strMonth
	Dim strYear
	Dim strSQL
	Dim rsLonn
	Dim Sum
	Dim rsAntAvd
	Dim AntallAvd
	Dim AvdClause
	
	' initierer variabelverdier
	lPeriode = Request("lPeriode")
	
	' Get a database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<title>Kontrollrapport lønn</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script type="text/javascript" language="javascript" src="../Js/javaScript.js"></script>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Kontrollrapport for overføring av lønn</h1>
			</div>
			<div class="content">
				<FORM ACTION="Loen_kontroll_vis.asp" METHOD="post">
					<p>Lønnsperiode (yyyymm): <INPUT SIZE="6" maxlength="6" TYPE="TEXT" NAME="lPeriode" VALUE="<% If lPeriode <> "" Then Response.Write lPeriode %>">
					<INPUT TYPE="SUBMIT" VALUE="Søk">
					</p>
				</form>
				<%
				If lPeriode <> "" Then
					strSQL = "SELECT stoerste = MAX(AvdelingID) FROM avdeling"
					set rsAntAvd = GetFirehoseRS(strSQL, Conn)
					If HasRows(rsAntAvd) then
						AntallAvd = rsAntAvd("stoerste")	
						rsAntAvd.close
					Else
						AntallAvd = 1
					End if
					set rsAntAvd = Nothing
					
					For i = 1 To AntallAvd + 1
						SELECT Case i
							Case AntallAvd + 1: avd = "ALLE"
							Case else
								Avd = UCase(HentAvdNavn(i))
						End Select
						Response.Write "<p><strong>Avdeling:</strong>&nbsp;" & avd & "<p>Lønn overført til Huldt &amp; Lillevik i periode " & lPeriode & "</p>"
						avdnr = i

						if i <> AntallAvd + 1 then 
							AvdClause = " AND L.Avdeling = " & avdnr
						Else
							AvdClause = ""	
						End if	
						' sql
						strSQL = "SELECT L.loennsartnr, LA.Loennsart, antall=Sum(Antall), belop=Sum(Beloep)" &_
							" FROM VIKAR_LOEN_VARIABLE L, H_LOENNSART LA" &_
							" WHERE Overfor_loenn_status = 3" &_	
							" AND L.Loennperiode = "& lPeriode &_	
							" AND L.Loennsartnr = LA.Loennsartnr" &_
							AvdClause &_
							" AND L.Loennsartnr not IN('29','40','41','42','44','45','46','51','80','83','85','90','93','96','98')" &_
							" GROUP BY L.loennsartnr, LA.Loennsart"

						set rsLonn = GetFirehoseRS(strSQL, Conn)
						' Vis data
						%>
						<div class="listing">
							<table>
								<tr>
									<th>L.ArtNr</th>
									<th>Lønnsart</th>
									<th>Antall</th>
									<th>Beløp</th>
								</tr>
								<%
								If HasRows(rsLonn) Then
									sum = 0
									do while not rsLonn.EOF 
										%>
										<tr>
											<td class="right"><% =rsLonn("loennsartnr") %></td>
											<td ><% =rsLonn("loennsart") %></td>
											<td class="right"><% =rsLonn("antall") %></td>
											<td class="right"><% =formatNumber(rsLonn("belop"),2) %></td>
										</tr>
										<%
										sum = sum + rsLonn("belop")
										rsLonn.MoveNext
									loop
									rsLonn.Close
									Set rsLonn = Nothing 
								End If ' ingen rader
								%>
								<tr>
									<td colspan="3">SUM LØNN</td>
									<td class="right"><% =formatNumber(sum, 2)%></td>
								</tr>
								<%
								' sql for kjøregodtgjørelse og annen godtgjørelse
								strSQL = "SELECT L.loennsartnr, LA.Loennsart, antall=Sum(Antall), belop=Sum(Beloep)" &_
									" FROM VIKAR_LOEN_VARIABLE L, H_LOENNSART LA" &_
									" WHERE Overfor_loenn_status = 3" &_
									" AND L.Loennperiode = "& lPeriode &_	
									" AND L.Loennsartnr = LA.Loennsartnr" &_
									AvdClause &_
									" AND L.Loennsartnr IN('29','40','41','65','80','83','85','90','93','96')" &_
									" GROUP BY L.loennsartnr, LA.Loennsart"
									
								set rsLonn = GetFirehoseRS(strSQL, Conn)
								sum = 0
								' Vis data
								If HasRows(rsLonn) Then
									do while not rsLonn.EOF 
										%>
										<tr>
											<td class="right"><% =rsLonn("loennsartnr") %></td>
											<td><% =rsLonn("loennsart") %></td>
											<td class="right"><% =rsLonn("antall") %></td>
											<td class="right"><% =formatNumber(rsLonn("belop"),2) %></td>
										</tr>
										<%
										sum = sum + rsLonn("belop")
										rsLonn.MoveNext
									loop
									rsLonn.Close
								End If ' ingen rader
								Set rsLonn = Nothing
								%>
								<tr>
									<td colspan="3">SUM ANNEN GODTGJ.</td>
									<td class="right"><% =formatNumber(sum,2) %></td>
								</tr>
								<%
								' sql for trekk
								strSQL = "SELECT L.loennsartnr, LA.Loennsart, antall=Sum(Antall), belop=Sum(Beloep)" &_
									" FROM VIKAR_LOEN_VARIABLE L, H_LOENNSART LA" &_
									" WHERE Overfor_loenn_status = 3" &_
									" AND L.Loennperiode = "& lPeriode &_	
									" AND L.Loennsartnr = LA.Loennsartnr" &_
									AvdClause &_
									" AND L.Loennsartnr in('42','44','45','46','51','98')" &_
									" GROUP BY L.loennsartnr, LA.Loennsart"


								set rsLonn = GetFirehoseRS(strSQL, Conn)
								sum = 0
								' Vis data
								If HasRows(rsLonn) Then
									do while not rsLonn.EOF 
										%>
										<tr>
											<td class="right"><% =rsLonn("loennsartnr") %></td>
											<td><% =rsLonn("loennsart") %></td>
											<td class="right"><% =rsLonn("antall") %></td>
											<td class="right"><% =formatNumber(rsLonn("belop"),2) %></td>
										</tr>
										<%
										sum = sum + rsLonn("belop")
										rsLonn.MoveNext
									loop
									rsLonn.Close
								End If ' ingen rader
								Set rsLonn = Nothing
								%>
								<tr>
									<td colspan="3">SUM TREKK</td>
									<td class="right"><% =formatNumber(sum,2) %></td>
								</tr>
							</table>
						</div>
						<%
					Next 'avdeling
				End If 'peiode er lagt inn 
				%>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>