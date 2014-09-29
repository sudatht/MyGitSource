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
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim strRegistrertVikar
	dim Conn
	dim strSQL
	dim lVikarID
	dim rsVikar
	dim rsBakgrunn
	dim strPoststed
	dim strHeading
	dim vikNavn
	
	strRegistrertVikar = request.querystring("vikarID")

	' Move parameters to local variables
	lVikarID = Request.Querystring("suspectID")

	' Check VikarID
	If lenb(lVikarID) = 0 Then
		AddErrorMessage("Etternavn ikke utfyllt!")
		call RenderErrorMessage()
	End If

	' Get a database connection
	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))	

	if request.querystring("aktuell")="ja" Then
		strSQL = "UPDATE V_SUSPECT SET Uaktuell = 0 WHERE  suspectID = "& lVikarID
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			CloseConnection(Conn)
			set Conn = nothing			
			AddErrorMessage("Kunne ikke sette uaktuell suspect aktuell!")
			call RenderErrorMessage()			
		End if		
		CloseConnection(Conn)
		set Conn = nothing		
		response.redirect "SuspectListUaktuelle.asp"
	End if

	if request.querystring("slett")="ja" Then
		strSQL = "UPDATE V_SUSPECT SET Slettes = 1 WHERE  suspectID = "& lVikarID
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Feil under sletting av uaktuell suspect!")
			call RenderErrorMessage()
		End if
		CloseConnection(Conn)
		set Conn = nothing				
		response.redirect "SuspectListUaktuelle.asp"
	End if

	' Get all data
	If lVikarID <> "" Then
		strSQL = "SELECT V.suspectID, V.regdato, V.Etternavn, V.Fornavn, V.foedselsdato,  " &_
			"A.Adresse, A.Postnr, A.PostSted, T.Vikartype, S.VikarStatus, V.AnsMedID,"&_
			"V.notat, Telefon, MobilTlf, Fax, Personsoek, EPost " &_
			"FROM V_SUSPECT V,H_VIKAR_TYPE T, H_VIKAR_STATUS S,V_SUSPECT_ADRESSE A " &_
			" WHERE V.suspectID = " & lVikarID  &_
			" AND V.TypeID *= T.VikarTypeID  " &_
			" AND V.StatusID *= S.VikarStatusID " &_                             
			" AND V.suspectID = A.adresseRelID " & _
			" AND A.AdresseType = 2 "

		SET rsVikar = GetFirehoseRS(strSQL, Conn)

		strsuspectID	= rsVikar("suspectID")
		strRegDato		= rsVikar("regdato")
		strEtternavn	= rsVikar("Etternavn") 
		strFornavn		= rsVikar("Fornavn") 
		strFoedselsdato = rsVikar("Foedselsdato")   
		strNotat		= rsVikar("Notat")   
		strStatus		= rsVikar("VikarStatus")
		strVikarType	= rsVikar("Vikartype")   
		strTelefon		= rsVikar("Telefon")
		strPersonsoek	= rsVikar("Personsoek") 
		strFax			= rsVikar("Fax") 
		strMobilTlf		= rsVikar("MobilTlf") 
		strEPost		= rsVikar("EPost")  
		strAnsMed		= rsVikar("AnsMedID")  
		strAdresse		= rsVikar("Adresse")

		' Create poststed
		strPoststed = rsVikar("Postnr") & " " & rsVikar("PostSted")
		strHeading = "Viser uaktuell suspect " & rsVikar("Fornavn") & " " &rsVikar("Etternavn")
		vikNavn = rsVikar("Fornavn") & " " &rsVikar("Etternavn")

		' Close AND release recordset
		rsVikar.Close
		Set rsVikar = Nothing

		' Get all connected AVDELINGSKONTOR
			strSQL = "SELECT avdKontor " &_
			" FROM V_SUSPECT_AVDKONTOR " &_
			" WHERE suspectID = " & strsuspectID &_ 
			" ORDER BY avdKontor"

		Set rsVikarAvdelingsKtr = GetFirehoseRS( strSQL, Conn )
	End If

	'sletter enkeltlinjer i kompetansen
	If Request.QueryString("slett")="ja" Then
		strID = Request.QueryString("ID")
		strSQL = "DELETE FROM V_SUSPECT_KOMPETANSE WHERE kompetanseID = " &strID
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Feil under sletting av kompentanse for uaktuell suspect!")
			call RenderErrorMessage()
		End if
		CloseConnection(Conn)
		set Conn = nothing
		Response.Redirect "v_suspectvis.asp?suspectID=" & strsuspectID
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"  "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<title><%=strHeading %></title>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>
					<%
					if lenb(strRegistrertVikar) =  0 then 
						Response.Write strHeading
					else 
						%>
						Response.Write vikNavn & " er nå registrert som vikar nr. " & strRegistrertVikar
						<% 
					end if 
					%>
				</h1>
			</div>
			<div class="content">
				<form name="suspekt" action="SuspectDB.asp?oppdat=ja&suspectID=<%=strsuspectID%>" METHOD="post">
					<table>
						<tr>
							<th>Suspectnr:</th>
							<td><%=strsuspectID %></td>
							<th>Sett ansvarlig:</th>
							<td>
								<SELECT NAME="dbxMedarbeider" onkeydown="typeAhead()">
									<option VALUE="0"></option>
									<%
									' Get ansvarlig medarbeider
									strSQL = "SELECT MedId, Fornavn, Etternavn FROM medarbeider ORDER BY Etternavn"
									Set rsMedarbeider = getFirehoseRS(strSQL, Conn)
					
									Do Until rsMedarbeider.EOF 
										If rsMedarbeider("MedID") = strAnsMed Then
											sel = " SELECTED"
										Else
											sel = ""
										End If 
										strName = rsMedarbeider("Etternavn") & "  " & rsMedarbeider("Fornavn") 
										%>
										<option VALUE="<% =rsMedarbeider("MedID")%>" <% =sel %>><% =strName %></option>
										<%   
										rsMedarbeider.MoveNext
									Loop
									rsMedarbeider.Close
									Set rsMedarbeider = Nothing
									%>
								</select>
	 						</td>
							<td><INPUT TYPE="submit" VALUE="Overfør til ansvarlig"></td>
						</tr>
						<tr>
							<th>Fornavn:</th>
							<td><%=strFornavn %></td>
							<th>Etternavn:</th>
							<td><%=strEtternavn %></td>
							<td><a href="SuspectvisUaktuelle.asp?aktuell=ja&suspectID=<%=strsuspectID%>">Sett søker aktuell</a></td>
						</tr>
						<tr>
							<th>Telefon:</th>
							<td><%=strTelefon %></td>
							<th>Mobil:</th>
							<td><%=strMobilTlf %></td>
							<td><a href="SuspectvisUaktuelle.asp?slett=ja&suspectID=<%=strsuspectID%>">Slett søker</a></td>
						</tr>
						<tr>
							<th align=left>E-Post:</th>
							<td><a href="mailto:<%=strEPost%>" title="Send e-post til suspect"><%=strEPost%></td>
							<th>Reg.dato:</th>
							<td><%=strRegDato%></TD>
						</tr>	
						<tr>
							<th>Hjemmeadresse:</th>
							<td colspan="2"><%=strAdresse%></td>
							<th>Poststed:</th>
							<td colspan="2"><%=strPoststed %></td>
						</tr>
						<%
						' Get all connected  AVDELING
						strSQL = "SELECT A.Avdeling " &_
								" FROM V_SUSPECT_AVDELING VA, AVDELING A " &_
								" WHERE VA.suspectID = " & strsuspectID &_
								" AND VA.AvdelingID = A.AvdelingID" &_
								" ORDER BY A.Avdeling"

						Set rsVikarAvdeling = getFirehoseRS(strSQL, Conn)

						' Print lead text
						Response.Write "<tr><th>" & "Avdelinger:" & "</th>"
						if HasRows(rsVikarAvdeling) then
							' loop on result and display in table
							Do Until rsVikarAvdeling.EOF      
								Response.Write "<td>" & rsVikarAvdeling("Avdeling" ) & "</td>"
								rsVikarAvdeling.MoveNext
							Loop
							' Close recordset
							rsVikarAvdeling.Close
						end if						
						Set rsVikarAvdeling = Nothing
						Response.Write "</tr>"
						'avdelingskontor

						' Print lead text
						Response.Write "<tr><th>" & "Avdelingskontor:" & "</th>"
							
						if HasRows(rsVikarAvdelingsKtr) then
							' loop on result and display in table
							Do Until rsVikarAvdelingsKtr.EOF     
								Response.Write "<td>" & rsVikarAvdelingsKtr("avdKontor" ) & "</td>"
								rsVikarAvdelingsKtr.MoveNext
							Loop
							Response.Write "</tr>"							  
							' Close recordset
							rsVikarAvdelingsKtr.Close
						end if
						Set rsVikarAvdelingsKtr = Nothing
						%>
					</table> 
				</form>   
				<table>
					<tr>
						<td>
							<table>
								<tr>
									<th>Kompetanse</th>
									<tr>&nbsp;</td>
								</tr>
								<% 
								Set rsKompetanse = Conn.Execute("SELECT KompetanseId, KType, V.K_TypeID, V.suspectID, KTittel, KLevel, Rangering, kommentar FROM V_SUSPECT_KOMPETANSE V, H_KOMP_TYPE TY, H_KOMP_TITTEL T, H_KOMP_LEVEL L  WHERE V.suspectID = " & lvikarID & " AND V.K_TypeID *= TY.K_TypeID AND V.K_TypeID *= T.K_TypeID AND V.K_TittelID *= T.K_TittelID AND V.K_LevelID *= L.K_LevelID order by V.K_TypeID,V.K_TittelID")
								if HasRows(rsKompetanse) then
									Do Until rsKompetanse.EOF
										%>
										<tr>
											<td><%=rsKompetanse("KTittel") %></td>
											<td></td>
										</tr>
										<%
									rsKompetanse.MoveNext
									Loop
									' Close AND release recordset
									rsKompetanse.Close
								end if
								Set rsKompetanse = Nothing
								%>
							</table>
						</td>
					</tr>
				</table>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>