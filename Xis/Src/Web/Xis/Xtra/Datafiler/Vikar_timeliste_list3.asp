<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	Function seType(t)
		SELECT Case t
			Case 1: seType = "Lønnsmottager"
			Case 2: seType = "Selvstendig"
			Case 3: seType = "Aksjeselskap"
			Case Else
		End Select
	End Function

	dim strMelding
	dim valgt_avd
	dim strEndret
	dim blnEndret
	dim strSQL
	dim Conn
	dim rsAvdnavn
	dim rsVikar
	dim firma
	
	' PARAMETERS
	strMelding = ""

	If Request.QueryString("viskode") = "" Then
		viskode = Cint(Request.Form("viskode"))
	Else
		viskode = Cint(Request.QueryString("viskode"))
	End If

	Session("viskode") = viskode

	mm = Datepart("m", Date): mm = mm + 1
	yy = Datepart("yyyy", DateValue( Date))
	If mm = 13 Then yy = yy + 1: mm = 1
	If mm < 10 Then dd = "01.0" Else dd="01."
	yy = Right(CStr(yy),2)
	dato1 = dd & mm & "." & yy
	dato2 = (Date - 60)

	If Request.Form("Dato1") <> "" Then
		dato1 = Request.Form("Dato1")
	End If
	
	If Request.QueryString("Dato1") <> "" Then
		dato1 = Request.QueryString("Dato1")
	End If
	
	If Request.Form("Dato2") <> "" Then
		dato2 = Request.Form("Dato2")
	End If
	
	If Request.QueryString("Dato2") <> "" Then
		dato2 = Request.QueryString("Dato2")
	End If

	If Request.form("velgAvdeling") <> "" Then
		valgt_avd = CInt(Request.form("velgAvdeling"))
	else
		valgt_avd = 0
	End If

	If (Request.QueryString("avd") <> "")  Then
		valgt_avd = CInt(Request.QueryString("avd"))
	End If

	If (Request.QueryString("avd2") <> "")  Then
		valgt_avd = CInt(Request.QueryString("avd2"))
	End If

	If (Request.form("avdel") <> "")  Then
		valgt_avd = CInt(Request.form("avdel"))
	End If

	Session("limitDato") = dato1
	Session("stoppDato") = dato2

	' BUTTONS
	If viskode = 1 Then k = "-> " Else k = "  "
	If viskode = 2 Then kk = "-> " Else kk = "        "
	If viskode = 3 Then kkk = "-> " Else kkk = "   "

	' Get a database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	strSQL = "SELECT avdeling FROM Avdeling WHERE avdelingID = " & valgt_avd
	Set rsAvdnavn = GetFirehoseRS(strSQL, Conn)

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
	<html>
		<head>
			<title>Timelister</title>
			<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
			<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
			<script type="text/javascript" language="javascript" src="../js/javascript.js"></script>
			<script type="text/javascript" language="javascript">
				function transferVals()
				{
					document.forms.varlloenn.velgAvdeling.value = document.forms.sok.avdel.value;
					return(true);
				};
			</script>
		</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Behandle timelister</h1>
			</div>
			<div class="content">
				<FORM NAME="HH" ACTION="Vikar_timeliste_list3.asp?viskode=1&dato2=<% =dato2 %>" METHOD=POST ID="Form1">
					<INPUT TYPE=SUBMIT VALUE="<% =k %> Ny timelister" ID="Submit1" NAME="Submit1">
					<input type=text size=7 NAME=DATO1 VALUE="<% =dato1 %>" ONBLUR="dateCheck(this.form, this.name)" ID="Text1">
				</form>
				<br>
	
				<FORM NAME="GG" ACTION="Vikar_timeliste_list3.asp?viskode=3&dato1=<% =dato1 %>" METHOD=POST ID="Form2">
					<INPUT TYPE=SUBMIT VALUE="<% =kkk %>Gamle" ID="Submit2" NAME="Submit2">
					<input type=text size=7 NAME="DATO2" VALUE="<% =dato2 %>"  ONBLUR="dateCheck(this.form, this.name)" ID="Text2">
				</form>
				<br>
				<table>
					<tr>
						<td>
							<form name="sok" id="sok" ACTION="Vikar_timeliste_list3.asp?viskode=1&avd=1&dato1=<% =dato1 %>" METHOD=POST>
							<SELECT NAME="avdel" ID="Select1">
								<OPTION VALUE="0" <% =sel %>></OPTION>
								<%
								strSQL = "SELECT avdelingID, avdeling FROM avdeling"
								set rsAvdeling = GetFirehoseRS(strSQL, Conn)
								while (not rsAvdeling.EOF)
									If valgt_avd = cint(rsAvdeling("AvdelingID")) Then
										sel = "Selected"
									Else
										sel = ""
									End if
									%>
									<OPTION VALUE="<% =rsAvdeling("AvdelingID")%>" <%=sel %>><%=rsAvdeling("Avdeling")%></OPTION>
									<% 
									rsAvdeling.MoveNext
								wend
								rsAvdeling.close
								set rsAvdeling = Nothing
								
								If viskode < 2 Or viskode = 3 Then
									If valgt_avd = "-2" Then sel = "SELECTED" Else sel = "" 
									%>
									<OPTION VALUE="-2" <% =sel %>>Selvstendig</OPTION>
									<% 
									If valgt_avd = "-1" Then sel = "SELECTED" Else sel = "" 
									%>
									<OPTION VALUE="-1" <% =sel %>>AS</OPTION>
									<% 
								End If 
								%>
							</SELECT>
							<input type=hidden name=DATO1 VALUE="<% =dato1 %>" ID="Hidden1">
						</td>
						<td>
							<INPUT TYPE="submit" value="   Søk   " ID="Submit3" NAME="Submit3">
						</td>
					</form>
					<form onSubmit="transferVals()" name="varlloenn" id="varlloenn" action="Vikar_timeliste_list3.asp?viskode=2&dato1=<%=dato1%>" name="varlloenn" method="post">
						<td align="left"><input type="hidden" name="velgAvdeling" value="1" ID="Hidden2"><input TYPE="submit" value="<% =kk %>Variabel lønn -  overføring " ID="Submit4" NAME="Submit4"></td>
					</form>
				</tr>
			</table>
<%
' Run different queries
If (viskode = 1) Then

	If Request.QueryString("Avd") = 1 Then
		opt = Request.Form("AVDEL")
	
	If opt < 0 Then   'selvstendige og AS
		strSQL = "SELECT DISTINCT VA.Ansattnummer, D.VikarID, D.OppdragID, Etternavn, Fornavn, Fakturastatus, Loennstatus, vikar.H_L_Status, " &_
		"D.FirmaID, stat=TimelisteVikarStatus, Firma, fDato=Foedselsdato, VIKAR.TypeID, SoCuid " &_
		"FROM  DAGSLISTE_VIKAR D, VIKAR, FIRMA, VIKAR_ANSATTNUMMER VA " &_
		"WHERE D.TimelisteVikarStatus < 6 "  &_
		" AND (D.Loennstatus < 3 or D.Fakturastatus < 3) " &_
		" AND Vikar.VikarID *= VA.VikarID " & _
		" AND VIKAR.TypeID = " & (opt + 4) &_
		" AND D.VikarID = VIKAR.VikarID " &_
		" AND D.Dato < " & DbDate(session("limitDato"))  &_
		" AND D.FirmaID = FIRMA.FirmaID " &_
		" ORDER BY VIKAR.TypeID, Etternavn"

	Else		'kurs, data, dokument
		strSQL = "SELECT DISTINCT VA.Ansattnummer, D.VikarID, D.OppdragID, Etternavn, Fornavn, Fakturastatus, Loennstatus, vikar.H_L_Status, " &_
		"D.FirmaID, stat=TimelisteVikarStatus, Firma, fDato=Foedselsdato, VIKAR.TypeID, SoCuid " &_
		"FROM DAGSLISTE_VIKAR D, VIKAR, FIRMA, OPPDRAG, VIKAR_ANSATTNUMMER VA  " &_
		"WHERE D.TimelisteVikarStatus < 6 "  &_
		" AND (D.Loennstatus < 3 or D.Fakturastatus < 3)" &_
		" AND OPPDRAG.OppdragID = D.OppdragID" &_
		" AND OPPDRAG.AvdelingID = " & opt &_
		" AND Vikar.VikarID *= VA.VikarID" & _
		" AND D.VikarID = VIKAR.VikarID" &_
		" AND D.Dato < " & DbDate(session("limitDato"))  &_
		" AND D.FirmaID = FIRMA.FirmaID " &_
		" ORDER BY VIKAR.TypeID, Etternavn"
	End If
Else
	strSQL = "SELECT DISTINCT VA.Ansattnummer, D.VikarID, D.OppdragID, Etternavn, Fornavn, Fakturastatus, Loennstatus, vikar.H_L_Status, " &_
	"D.FirmaID, stat=TimelisteVikarStatus, Firma, fDato=Foedselsdato, VIKAR.TypeID, SoCuid " &_
	"FROM DAGSLISTE_VIKAR D, VIKAR, FIRMA, VIKAR_ANSATTNUMMER VA " &_
	"WHERE D.TimelisteVikarStatus < 6 "  &_
	" AND (D.Loennstatus < 3 or D.Fakturastatus < 3)" &_
	" AND D.VikarID = VIKAR.VikarID" &_
	" AND Vikar.VikarID *= VA.VikarID" & _
	" AND D.Dato < " & DbDate(session("limitDato"))  &_
	" AND D.FirmaID = FIRMA.FirmaID " &_
	" ORDER BY VIKAR.TypeID, Etternavn"
End If 'søk på avdeling

set rsVikar = GetFirehoseRS(strSQL, conn)

Elseif viskode = 2 Then

	if valgt_avd > 0 then
		strSQL = "SELECT DISTINCT VA.Ansattnummer, VIKAR.VikarID, Etternavn, Fornavn, stat=Overfor_loenn_status, Loenndato, oppdragID, F.FirmaID, Avdeling, vikar.H_L_Status, F.SoCuid " &_
				"FROM VIKAR_LOEN_VARIABLE, Firma F, VIKAR, VIKAR_ANSATTNUMMER VA " &_
				"WHERE VIKAR_LOEN_VARIABLE.VikarID = VIKAR.VikarID " &_
				" AND VIKAR_LOEN_VARIABLE.FirmaID = F.FirmaID" & _				
				" AND Vikar.VikarID *= VA.VikarID" & _
				" AND Avdeling = " & valgt_avd &_
				" AND Overfor_loenn_status = 2  "

		'se strSQL & "<br>"
		set rsVikar = GetFirehoseRS(strSQL, conn)
	else
		strMelding = "Du må velge en avdeling!<br>"
	end if

Elseif viskode = 3 Then

	strSQL = "SELECT DISTINCT VA.Ansattnummer, v.VikarID, v.Etternavn, v.Fornavn, stat = vl.Overfor_loenn_status, vl.Loenndato, vl.OppdragID, vl.FirmaID, A.avdeling, v.H_L_Status, F.SoCuid " &_
		"FROM VIKAR_LOEN_VARIABLE vl, Firma F, VIKAR v, AVDELING A, VIKAR_ANSATTNUMMER VA " &_
		" WHERE vl.VikarID = v.VikarID" &_
		" AND vl.FirmaID = F.FirmaID" & _			
		" AND v.VikarID *= VA.VikarID" & _
		" AND a.avdelingID = vl.avdeling" &_
		" AND vl.Loenndato >= " & DbDate(dato2) & _
		" AND vl.Overfor_loenn_status = 3 " &_
		" ORDER BY VA.Ansattnummer"

	set rsVikar = GetFirehoseRS(strSQL, conn)	

End If

	if viskode = 2 AND valgt_avd = 0 then
		%>
		<tr>
			<td><%=strMelding%></td>
		</tr>
		<%
	elseif viskode = 2 AND valgt_avd > 0 then
		%>
		<tr>
			<td>Avdeling: <%=rsAvdnavn("avdeling")%></td>
		</tr>
		<%
	end if

If (viskode > 0 AND IsObject(rsVikar)) Then

	' No records found
	If (NOT rsVikar.EOF) Then

		If viskode = 1 Then 
		%>
			<A HREF=#A>A</A>
			<A HREF=#B>B</A>
			<A HREF=#C>C</A>
			<A HREF=#D>D</A>
			<A HREF=#E>E</A>
			<A HREF=#F>F</A>
			<A HREF=#G>G</A>
			<A HREF=#H>H</A>
			<A HREF=#I>I</A>
			<A HREF=#J>J</A>
			<A HREF=#K>K</A>
			<A HREF=#L>L</A>
			<A HREF=#M>M</A>
			<A HREF=#N>N</A>
			<A HREF=#O>O</A>
			<A HREF=#P>P</A>
			<A HREF=#R>R</A>
			<A HREF=#S>S</A>
			<A HREF=#T>T</A>
			<A HREF=#U>U</A>
			<A HREF=#V>V</A>
			<A HREF=#Ø>Ø</A>
			<A HREF=#Å>Å</A>
		<%
		End If 'viskode = 1
		%>
		</table>	
		<div class="listing">
			<table ID="Table2">
	<%
	If viskode = 1 Then
		' Display table headings
		%>
		<tr>
			<th>&nbsp;</th>
			<th>A.nr</th>
			<th>Oppd</th>
			<th>FDato</th>
			<th colspan="2">Navn</th>
			<th>Kontakt</th>
			<th>T.stat</th>
			<th>L.stat</th>
			<th>F.stat</th>
			<th>Lønn</th>
		</tr>
		<%

		' Show data
		Bok = Left(rsVikar("Etternavn"),1)

		TYID = ""
	Do Until rsVikar.EOF
		If TYID <> rsVikar("TypeID") Then 
		%>
			<tr>
				<th>&nbsp;</th>
				<th>&nbsp;</th>
				<th>&nbsp;</th>
				<th>&nbsp;</th>
				<th colspan="2"><% =seType(rsVikar("TypeID")) %></th>
			</tr>
			<%	
			TYID = rsVikar("TypeID")
		End If		
		%>
		<tr>
			<td>&nbsp;
			<%
			strFullName = rsVikar("Etternavn") & " " & rsVikar("Fornavn")
			If Bok <> Left(rsVikar("Etternavn"),1) Then
				Bok = ucase(Left(rsVikar("Etternavn"), 1))
				%><a name='<%=Bok%>' id='<%=Bok%>'></a>
				<%
			End If
			%>		
	  		 </td>
			<td><% =rsVikar("Ansattnummer") %></td>
			<td><% =rsVikar("OppdragID") %></td>
			<td><% =rsVikar("Fdato") %></td>
			<td><%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & rsVikar( "VikarID" ), strFullName, "Vis vikar " & strFullName )%></td>
			<% 
			If kode < 5 Then 
				%>
				<td><A HREF="Vikar_timeliste_vis3.asp?vikarID=<%=rsVikar("VikarID") %>&viskode=<% =viskode %>&tilgang=2&OppdragID=<% =rsVikar("OppdragId") %>&FirmaID=<% =rsVikar("FirmaID") %>&limitdato=<% =dato1 %>&frakode=2" TARGET=_new title="Vis timeliste" >Timeliste</A></td>
				<% 
			Else 
				%>
				<td>&nbsp;</td>
				<% 
			End If 
			firma = Left(rsVikar("Firma"), 28)
			if(isnull(rsVikar("SOCuid")) = false) then
				firma = CreateSONavigationLink(SUPEROFFICE_PANEL_CONTACT_URL, SUPEROFFICE_PANEL_CONTACT_URL, rsVikar("SOCuid"), firma, "Vis kontakt '" & rsVikar("Firma") & "'")
			end if			
			%>
			<td><%=firma%></td>
			<td>
				<% If rsVikar("stat") = 5 Then %>
					<FONT COLOR="GREEN" style="background:#e6e6e6; font-weight:bold;">
				<% ElseIf rsVikar("stat") = 4 Then %>
					<FONT COLOR="ORANGE" style="background:#e6e6e6; font-weight:bold;">
				<% ElseIf rsVikar("stat") = 2 Then %>
					<FONT COLOR="BLUE" style="background:#e6e6e6; font-weight:bold;">
				<% Else %>
					<FONT COLOR="#CB1700" style="background:#e6e6e6; font-weight:bold;">
				<% End If %>
				(<% =rsVikar("stat") %>)</font>
			</td>
			<td>
				<% If rsVikar("Loennstatus") = 2 Then %>
					<FONT COLOR=GREEN style="background:#e6e6e6; font-weight:bold;">
				<% ElseIf rsVikar("Loennstatus") = 1 Then %>
					<FONT COLOR=#CB1700 style="background:#e6e6e6; font-weight:bold;">
				<% End If %>
				(<% =rsVikar("Loennstatus") %>)</font>
			</td>
			<td>
				<% If rsVikar("Fakturastatus") = 2 Then %>
					<FONT COLOR=GREEN style="background:#e6e6e6; font-weight:bold;">
				<% ElseIf rsVikar("Fakturastatus") = 1 Then %>
					<FONT COLOR=#CB1700 style="background:#e6e6e6; font-weight:bold;">
				<% End If %>
				(<% =rsVikar("Fakturastatus") %>)</font>
			</td>
			<td>
				<% If rsVikar("stat") = 5 Then %>
					<A HREF=Vikar_varl_vis3.asp?VikarID=<% =rsvikar("VikarID") %>&OppdragID=<% = rsVikar("OppdragID") %>&FirmaID=<% = rsVikar("FirmaID") %>&frakode2=2 >Lønn</A>
				<% End If %>
				&nbsp;
			</td>
		</tr>
		<%
		rsVikar.MoveNext
		count = count + 1
	Loop

	ElseIf viskode = 2 Or viskode = 3 Then ' Vis gamle timelister
		' Display table headdings
		%>	
		<tr>
			<th>A.nr</th>
			<th>Vikar</th>
			<th>Var.Lønn</th>
			<th>Lønnsdato</th>
			<th>L.stat</th>
			<th>Avdeling</th>
			<th>Endret</th>
			<th>Nedgrader</th>
		</tr>
		<%
		' Show data
		count = 1
		VID = rsVikar("VikarID")

		strEndret = ""
		blnEndret = false
		Do Until rsVikar.EOF

			strFullName   = rsVikar("Etternavn") & " " & rsVikar("Fornavn")
			if (rsVikar("H_L_Status")=1) then
				strEndret = "Ja"
				blnEndret = true
			else
				strEndret = "Nei"
			end if

			optname = "opt" & count
			optname2 = optname & count
			If rsVikar("stat") = 2 Then
				merket = "checked"
			Else
				merket = ""
			End If

			If viskode = 3 Then
				If rsVikar("VikarID") <> VID Then 
				%>
					<th><A HREF="Vikar_varl_lagre3.asp?VikarID=<% =VID %>&nedgrad=Ja&Loenndato=<% =loenndato %>&Loennstatus=3&avd=<%=valgt_avd%>&viskode=2&dato1=<%=dato1%>" > nedgrader</A>
					<% 
					VID = rsVikar("VikarID")
				End If
			End If 
			%>
			<tr>
				<th class="right"><% = rsVikar("Ansattnummer") %></th>
				<%
				If viskode = 3 Then 
					loenndato = rsvikar("Loenndato") 
					%>
					<th><%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & rsVikar( "VikarID" ), strFullName, "Vis vikar " & strFullName )%></th>
					<th><A HREF="Vikar_varl_gml3.asp?vikarID=<%=rsVikar("VikarID") %>&viskode=<% =viskode %>&LoennDato=<% =rsVikar("Loenndato") %>&frakode=2&gmlloenn=Ja&OppdragID=<% =rsVikar("OppdragID") %>" TARGET=new>Varl.lønn</A></th>
					<th><% =loenndato %></th>
					<th>(<% =rsVikar("stat") %>)</th>
					<th><% =rsvikar("avdeling") %></th>
					<th><% =strEndret%></th>
					<% 
				Else 
					%>
					<th><%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & rsVikar( "VikarID" ), strFullName, "Vis vikar " & strFullName )%></th>
					<th><A HREF="Vikar_varl_vis3.asp?vikarID=<%=rsVikar("VikarID") %>&viskode=<% =viskode %>&OppdragID=<% =rsVikar("OppdragID") %>&FirmaID=<% =rsVikar("FirmaID") %>&frakode2=2" TARGET=new>Varl.Lønn</A></th>
					<th><% =rsVikar("Loenndato") %></th>
					<th>(<% =rsVikar("stat") %>)</th>
					<th><% =rsvikar("avdeling") %></th>
					<th><% =strEndret%></th>
					<th><A HREF="Vikar_varl_lagre3.asp?VikarID=<% =rsVikar("VikarID") %>&nedgrad=Ja&Loenndato=<% =rsVikar("Loenndato") %>&Loennstatus=<% =rsVikar("stat") %>&avd=<%=valgt_avd%>&viskode=2&dato1=<%=dato1%>"> nedgrader</A></th>
					<% 
				End If 
				%>
			</tr>
			<%
			rsVikar.MoveNext
		Loop
		rsVikar.Close
	End If  'vis gamle timelister
	set rsVikar = nothing
	If viskode = 3 Then 
		%>
		<th><A HREF="Vikar_varl_lagre3.asp?VikarID=<% =VID %>&nedgrad=Ja&Loenndato=<% =loenndato %>&Loennstatus=3"> nedgrader</A>
		<% 
	End If 
	
	If viskode = 2 AND HasUserRight(ACCESS_ADMIN, RIGHT_ADMIN) = true Then 
		%>
		</table>
		<br>
		</form>
		<FORM ACTION="Eksp_HL_02.asp?viskode=<%=viskode%>&avd=<%=valgt_avd%>" METHOD=POST ID="frmTransfer">
			<%
			if (blnEndret=true) then
				%>
				<span class="menu disabled" title="Generere Huldt &amp; Lillevik lønnsdatafil">
					&nbsp;Generere lønnsfil
				</span>			
				<span class="menu" title="Generere Huldt &amp; Lillevik vikarfil">
					<a href="eksp_konsulent.asp">&nbsp;Generere vikarfil</a>
				</span>
				<%
			else
				%>
				<span class="menu" title="Generere Huldt &amp; Lillevik lønnsdatafil">
					<a href="#" onClick="javascript:document.all.frmTransfer.submit();">&nbsp;Generere lønnsfil</a>
				</span>
				&nbsp;
				<span class="menu disabled" title="Generere Huldt &amp; Lillevik vikarfil">
					&nbsp;Generere vikarfil
				</span>			
				<%
			end if
			%>
		</form>
		<% 
	End If 
	%>
	</table>
	<%  
End If 'ingen treff 
	%>
	<br>
		<% 
		If viskode = 1 Then 
			%>
			<A HREF=#A>A</A>
			<A HREF=#B>B</A>
			<A HREF=#C>C</A>
			<A HREF=#D>D</A>
			<A HREF=#E>E</A>
			<A HREF=#F>F</A>
			<A HREF=#G>G</A>
			<A HREF=#H>H</A>
			<A HREF=#I>I</A>
			<A HREF=#J>J</A>
			<A HREF=#K>K</A>
			<A HREF=#L>L</A>
			<A HREF=#M>M</A>
			<A HREF=#N>N</A>
			<A HREF=#O>O</A>
			<A HREF=#P>P</A>
			<A HREF=#R>R</A>
			<A HREF=#S>S</A>
			<A HREF=#T>T</A>
			<A HREF=#U>U</A>
			<A HREF=#V>V</A>
			<A HREF=#Ø>Ø</A>
			<A HREF=#Å>Å</A>
			<% 
		end If 'viskode 
End If 'viskode > 0 
set rsVikar = nothing
%>
				</div>
			</div>	
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>