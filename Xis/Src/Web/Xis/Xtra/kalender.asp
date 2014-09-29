<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="includes\MailLib.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.utils.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.Economics.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim VikarID
	dim SourceVikarID
	dim FraDato
	dim TilDato
	dim rsVikar 
	dim SokOppdragID
	dim strAnsattnummer
	dim EtterNavn
	dim ForNavn
	dim VikarFound : VikarFound = false
	dim RecordsFound : RecordsFound = false
	dim strSQL

	' Check input values
	If lenb(Request( "tbxPageNo")) = 0 Then ' Triggers on first visit of page
		' Do we have VikarID ?
		If Request.QueryString( "VikarID" ) = "" Then
			VikarID = ""
		Else
			VikarID = cstr(Trim(Request.QueryString( "VikarID" )))
			SourceVikarID = VikarID
		End If

		If Session("FraDato") <> "" Then
			FraDato = Session("Fradato")
		Else
			FraDato = Date
		End If

		If Session("TilDato") <> "" Then
			TilDato = Session("Tildato")
		Else
			TilDato =  DateAdd("m", 2, Date())
		End If

		SokOppdragID = Request.QueryString( "OppdragID" )

	Else ' Triggers on postback
		' Add values from current page
		strAnsattnummer	= Request.Form("tbxAnsattnr")
		VikarID			= cstr(trim(Request.Form( "tbxVikarID" )))
		SourceVikarID	= cstr(trim(Request.Form( "SourceVikarID" )))
		EtterNavn		= Request.Form( "tbxEtterNavn" )
		ForNavn			= Request.Form( "tbxForNavn" )
		Fradato			= Request.Form( "tbxFraDato" )
		if(len(FraDato) = 0) then
			 AddErrorMessage("Fradato mangler!")			
		end if
		Tildato			= Request.Form( "tbxTilDato" )
		if(len(Tildato) = 0) then
			 AddErrorMessage("Tildato mangler!")			
		end if		
		SokOppdragID	= Request.Form( "tbxOppdragID" )
		Session("Fradato") = Fradato
		Session("Tildato") = Tildato
	End If

	if(HasError()) then
		call RenderErrorMessage()
	end if

	' First time page called and search value exist ?
	set Conn = GetConnection(GetConnectionstring(XIS, ""))	

	If lenb(SokOppdragID) > 0 And lenb(VikarID) = 0 Then

		' Get vikar from current Oppdrag
		strSQL =  "SELECT DISTINCT VikarID from OPPDRAG_VIKAR where OppdragID = " & SokOppdragID
		set rsOppdragVikar = GetFirehoseRS(strSQL, Conn)
		strSQL = vbNullString
		If (HasRows(rsOppdragVikar) = true) Then
			VikarID = cstr(trim(rsOppdragVikar("VikarID")))
			rsOppdragVikar.Close
		End if

		set rsOppdragVikar = Nothing
	End If

	' Build sql-statement
	If lenb(Request( "tbxPageNo")) = 0 Then 'If first visit get vikar from vikarid
		If (VikarID <> "") Then
			strSQL = "VIKAR.VikarID = " & VikarID
		End If
	Elseif lenb(strAnsattnummer) > 0 and strAnsattnummer <> "0" Then
		strSQL = "VIKAR_ANSATTNUMMER.ansattnummer = " & strAnsattnummer
	Else
		If Etternavn <> "" Then
			strSQL = "Etternavn LIKE " & Quote( Etternavn & "%" )
		End If

		If Fornavn <> "" Then
			If strSQL <> "" Then
				strSQL = strSql & " AND "
			End If
			strSQL = strSQL & "Fornavn like " & Quote( Fornavn & "%" )
		End If
	End If

    If strSQL = "" Then
        strSQL = "VIKAR.VikarID > 0"
    End If

   ' Get vikar
   strSQL = "SELECT " & _
			"VIKAR.VikarID, " & _
			"Navn = (VIKAR.Fornavn + ' ' + VIKAR.Etternavn), " & _
			"VIKAR.Loenn1, " & _
			"VIKAR.TypeID, " & _
			"VIKAR_ANSATTNUMMER.ansattnummer " & _
			"FROM VIKAR " & _
			"LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " & _
			"WHERE " & strSql

	set rsVikar = GetFirehoseRS(strSQL, Conn)
	
	If (HasRows(rsVikar) = true) Then

		Navn = rsVikar( "Navn" ).Value
		VikarID = rsVikar("VikarID").Value
		strAnsattnummer = rsVikar("ansattnummer").Value
		Timeloenn = rsVikar("Loenn1").Value
		VikarType = rsVikar("TypeID").Value

		VikarFound = true

		strSQL = "SELECT OV.OppdragVikarId, OV.OppdragID, OV.Fradato, OV.Tildato, S.Status "&_
			", O.Beskrivelse, O.Kurskode, F.SOCuID, F.Firma, Program = KO.KTittel, Kompniva=KL.KLevel " &_
			" FROM OPPDRAG_VIKAR OV, OPPDRAG O, FIRMA F "&_
			", H_OPPDRAG_VIKAR_STATUS S, H_KOMP_TITTEL KO, H_KOMP_LEVEL KL" &_
			" WHERE OV.VikarID = " & VikarID &_
			" and OV.OppdragID = O.OppdragID " &_
			" and OV.FirmaID = F.FirmaID " &_
			" and OV.StatusID = S.OppdragVikarStatusID " &_
			" and O.ProgramID *= KO.K_TittelID "&_
			" AND KO.K_TypeID=3 " &_
			" AND O.Kompniva *= KL.K_LevelID " &_
			" AND ( ( OV.Fradato>= " & DbDate( Fradato ) & " and  OV.Fradato<= " & DbDate( Tildato ) & ") " &_
			" OR ( OV.Tildato>= " & DbDate( Fradato ) & " and  OV.Tildato<= " & DbDate( Tildato ) & ") " &_
			" OR ( OV.Fradato < " & DbDate( Fradato ) & " and  OV.Tildato > " & DbDate( Tildato ) & ") ) " &_
			" AND OV.StatusID IN (4,5) " &_
			" ORDER BY OV.Fradato "

		set rsCalender = GetFirehoseRS(strSQL, Conn)

		' No records found ?
		If (HasRows(rsCalender)= True) Then
			RecordsFound = true
		End If
		rsVikar.Close
   End If

   ' Close recordset
   set rsVikar = nothing

If SokOppdragID <> "" then

	strSQL = "SELECT OppdragID, StatusID, FirmaID, FraDato, Frakl, Tilkl, Kurskode, Timepris, Timerprdag, Lunch FROM OPPDRAG where OppdragID = "& SokOppdragID

	set rsOppdrVikarInfo = GetFirehoseRS(strSQL, Conn)

	If not rsOppdrVikarInfo.EOF then
		strFraDato	= cDate(rsOppdrVikarInfo("FraDato"))
		OppdragID	= rsOppdrVikarInfo("OppdragID")
		StatusID	= rsOppdrVikarInfo("StatusID")
		FirmaID		= rsOppdrVikarInfo("FirmaID")
		Frakl		= FormatDateTime( rsOppdrVikarinfo("Frakl"), 4 )
 		Tilkl		= FormatDateTime( rsOppdrVikarinfo("Tilkl"), 4 )
		oppdragTid	= rsOppdrVikarInfo("Kurskode")
 		Timepris	= rsOppdrVikarInfo("Timepris")
		T_prdag		= rsOppdrVikarInfo("Timerprdag")
		Lunch		= FormatDateTime( rsOppdrVikarinfo("Lunch"), 4 )
		
		if VikarType = 1 AND Timepris <> "" AND Timeloenn <> "" AND Timepris <> 0 AND Timeloenn <> 0 Then
			Faktor = Timepris / Timeloenn
			Faktor = Mid(Faktor, 1, 4)
		ElseIf VikarType <> 1 AND Timepris <> "" AND Timeloenn <> "" AND Timepris <> 0 AND Timeloenn <> 0 Then
			Faktor = Timepris / (Timeloenn / XIS_FACTOR)
			Faktor = Mid(Faktor, 1, 4)
		Else
			Faktor = 0
		End If
	End If
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"  "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<title>Kalender</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script language="javascript" src="Js/javascript.js" type="text/javascript"></script>
		<script language='javascript' src='js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='js/navigation.js' id='navigationScripts'></script>
		<script language='javascript' src='js/contentMenu.js'></script>
		<script language="javaScript" type="text/javascript">
			function sjekk()
			{
				var countBx = document.forms[1].elements.length;

				for (var i=0; i < countBx; i++)
				{
					if (document.forms[1].elements[i].type == "checkbox")
					{
						if (document.forms[1].elements[i].checked == true)
							document.forms[1].elements[i].checked = false;
						else
							document.forms[1].elements[i].checked = true;
					}
				}
			}

			function Faktor()
			{
				var status = document.forms[1].elements['tbxVikarStatus'].value;
				var loenn = document.forms[1].elements['tbxTimeloenn'].value;
				pris = document.forms[1].elements['tbxTimepris'].value;
				var faktor = document.forms[1].elements['tbxFaktor'].value;

				if (status == 1 && pris !=0 && loenn !=0 ){
					faktor = parseFloat(pris) / parseFloat(loenn);
					faktor=faktor.toString().substring(0,4);
					document.forms[1].elements['tbxFaktor'].value = faktor;
				}
				else if (status !=1 && pris !=0 && loenn !=0) {
					faktor =   parseFloat(pris)/(parseFloat(loenn)/XIS_FACTOR);
					faktor=faktor.toString().substring(0,4);
					document.forms[1].elements['tbxFaktor'].value = faktor;
				}
				else
				document.forms[1].elements['tbxFaktor'].value="";
			}
			
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
	<body onLoad="fokus()">
	<div class="pageContainer" id="pageContainer">
		<div class="contentHead1">
			<h1>Kalender</h1>
			<div class="contentMenu">
				<table cellpadding="0" cellspacing="0" width="96%">
					<tr>
						<td>
							<table cellpadding="0" cellspacing="2">
								<tr>
									<td class="menu" id="menu1" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
										<%
										if lenb(SourceVikarID) > 0 then
											%>
											<a href="/xtra/vikarvis.asp?vikarid=<%=SourceVikarID%>" title="Vis vikar">
											<img src="/xtra/images/icon_consultant.gif" width="18" height="15" alt="" align="absmiddle">Vis</a>
											<%
										else
											%>
											<a href="/xtra/OppdragVis.asp?OppdragID=<%=SokOppdragID%>" title="Vis oppdrag">
											<img src="/xtra/images/icon_job.gif" width="18" height="15" alt="" align="absmiddle">Vis
											<%
										end if
										%>
									</td>
								</tr>
							</table>
						</td>
						<td class="right">
						<!--#include file="Includes/contentToolsMenu.asp"-->
						</td>
					</tr>
				</table>
			</div>
		</div>	
		<div class="content">
			<form name="en" ACTION="kalender.asp" METHOD="POST" ID="Form1">
				<input type="hidden" NAME="tbxPageNo" VALUE="1" ID="tbxPageNo">
				<input type="hidden" NAME="tbxOppdragID" VALUE="<%=SokOppdragID%>" ID="tbxOppdragID">
				<input type="hidden" NAME="tbxVikarID" VALUE="<%=VikarID%>" ID="tbxVikarID">
				<input type="hidden" name="SourceVikarID" value="<%=SourceVikarID%>" id="SourceVikarID">
				<table>
					<tr>
						<td>Ansattnummer:</td>
						<td><input NAME="tbxAnsattnr" size="5" MAXLENGTH="5" value="<%=strAnsattnummer%>" ID="Text1"></td>
						<td>Etternavn:</td>
						<td><input NAME="tbxEtterNavn" size="15" MAXLENGTH="50" value="<%=EtterNavn%>" ID="Text2"></td>
						<td>Fornavn:</td>
						<td><input NAME="tbxForNavn"  size="15" MAXLENGTH="50" value="<%=ForNavn%>" ID="Text3"></td>
						<td>Fradato:</td>
						<td><input NAME="tbxFraDato" size="10" MAXLENGTH="10" value="<%=Fradato%>" ONBLUR="dateCheck(this.form, this.name)" ID="Text4"></td>
						<td>Tildato:</td>
						<td><input NAME="tbxTilDato" size="10" MAXLENGTH="10" value="<%=Tildato%>" ONBLUR="dateCheck(this.form, this.name)" ID="Text5"></td>
						<td><input type="submit" name="pbnDataAction" value="Søk"  onclick="dateInterval(this.form, this.name)" ID="Submit1"></td>
					</tr>
				</table>
			</form>
			<%
			If  (VikarFound = true) Then
				%>
				<form NAME='OppdragVikarVariabler' ACTION='kalenderDB.asp?vikarID=<%=VikarID%>&oppdragID=<%=SokOppdragID%>' METHOD='POST' ID="Form2">
					<table cellspacing='1' cellpadding='1' ID="Table2">
						<tr>
							<td><strong>Timelønn</strong></td>
							<td><INPUT TYPE='TEXT' NAME='tbxTimeloenn' SIZE=5 onChange="Faktor();" VALUE="<%=Timeloenn%>" ID="Text6"></td>
							<%
							If SokOppdragID <> "" then
								%>
								<th>Timepris:</th>
								<th><INPUT TYPE='TEXT' NAME='tbxTimepris' SIZE=5 onChange="Faktor();" VALUE="<%=Timepris%>" ID="Text7"></th>
								<th>Fra kl:</th><td><INPUT TYPE='TEXT' NAME='tbxFraKl' SIZE=5 onchange='timeCheck(this.form, this.name), workTime(this.form, this.name)' VALUE="<%=Frakl%>" ID="Text8"></td>
								<th>Til Kl:</th><td><INPUT TYPE='TEXT' NAME='tbxTilKl' SIZE=5 onchange='timeCheck(this.form, this.name), workTime(this.form, this.name)' VALUE="<%=Tilkl%>" ID="Text9"></td>
								<th>Lunsj:</th><td><INPUT TYPE='TEXT' NAME='tbxLunsj'  SIZE=5 onchange='timeCheck(this.form, this.name), workTime(this.form, this.name)' VALUE="<%=Lunch%>" ID="Text10"></td>
								<th>Ant. timer:</th><td><INPUT TYPE='TEXT' NAME='tbxTimerPrDag' SIZE=5 VALUE="<%=T_prdag%>" ID="Text11"></th>
								<th>Faktor:</th><td><INPUT TYPE='TEXT' NAME='tbxFaktor'  SIZE=5 VALUE="<%=Faktor%>" ID="Text12"></th>
								<INPUT TYPE='HIDDEN' NAME='tbxVikarStatus' VALUE="<%=VikarType%>" ID="Hidden4">
								<INPUT TYPE='HIDDEN' NAME='tbxFirmaID' VALUE="<%=FirmaID%>" ID="Hidden5">
								<%
							end if
							%>
						</tr>
						<%
						If SokOppdragID <> "" then						
							%>
							<tr>
								<th>Status:</th>
								<td colspan="2">
									<select NAME="dbxStatus" ID="Select1">
										<% 
										' Get Status
										set rsStatus = GetFirehoseRS("SELECT OppdragVikarStatusID, Status FROM h_oppdrag_vikar_status", Conn)
										Do Until rsStatus.EOF
											strValueSelected = rsStatus("OppdragVikarStatusID") & " SELECTED"
											strValueSelected = rsStatus("OppdragVikarStatusID")
											%>
											<option VALUE="<%=strValueSelected %>"><%=rsStatus("Status") %>
											<%    
											rsStatus.MoveNext
										Loop
										rsStatus.Close
										Set rsStatus = Nothing
										%>
									</select>
								</td>
								<td colspan="3"><input TYPE="BUTTON" OnClick="sjekk()" VALUE="Av/på sjekkbokser" ID="Button1" NAME="Button1"></td>
							</tr>
							<%
						End if 'SokOppdragID<>""
						%>
					</table>				
					<div class="listing">
						<table ID="Table3">
							<tr>
								<th>Dato</th>
								<th>Dag</th>
								<th>Tid</th>
								<th>Oppdrag</th>
								<th>Beskrivelse</th>
								<th>
								<%
								If SokOppdragID<>"" Then
									Response.Write"<INPUT TYPE='SUBMIT' VALUE='Velg'>"
								End if
								%>
								</th>
								<th>Kurstype</th>
								<th>Program</th>
								<th>Nivå</th>
								<th>Status</th>
								<th>Kontakt</th>
							</tr>
							<%
							If (RecordsFound = true) Then
								' Move recordset to array
								ArrCalender = rsCalender.GetRows
								' Close recordset
								rsCalender.Close
								set rsCalender = Nothing
							End If

							FirstDate = CDate( FraDato )
							LastDate = CDate( TilDato )

							NewDate = FirstDate
							Do Until NewDate > LastDate
								%>
								<tr>
									<td rowspan='2'><%=NewDate%></td>
									<td rowspan='2'><%=WeekdayName( WeekDay(NewDate) )%></td>
									<%
									' Loop on each date to check if Vikar is free
									Beskrivelse= ""
									OppdragID = ""
									FirmaID = ""
									Firma = ""
									Kurskode = ""
									Program = ""
									KompNiva = ""
									Status = ""
									tidStatus = ""
									bolDag = false
									bolKveld = false
									valg = ""

									If  (Recordsfound = true)   Then
										For Counter = LBound( ArrCalender, 2 ) To UBound( ArrCalender , 2 )
											If NewDate >= ArrCalender( 2, Counter ) And NewDate <=  ArrCalender( 3, Counter )  Then
												OppdragVikarID = ArrCalender( 0, Counter )
												OppdragID = ArrCalender( 1, Counter )
												Status = ArrCalender( 4, Counter )
												Beskrivelse = ArrCalender( 5, Counter )
												tidStatus = ArrCalender( 6, Counter )
												SOCuID = ArrCalender( 7, Counter )
												Firma = ArrCalender( 8, Counter )
												Program = ArrCalender( 9, Counter )
												Kompniva = ArrCalender( 10, Counter )

												If ArrCalender( 6, Counter ) < 2 THEN
													bolDag = true
													if oppdragID <> "" And oppdragID <> SokOppdragID Then
														oppdrag = CreateSOLink(SUPEROFFICE_XIS_TASK_URL, "", "oppdragvis.asp?oppdragID=" & OppdragID, OppdragID, "Vis Oppdrag" )
                  									else
														oppdrag = ""
													end if

													If Beskrivelse <> "" then
														Beskrivelse = CreateSOLink(SUPEROFFICE_XIS_TASK_URL, "", "OppdragvikarNy.asp?OppdragVikarID=" & OppdragVikarID, Beskrivelse, "Vis" )
													else
														Beskrivelse = "Ledig"
													End If

													kunde = CreateSONavigationLink(SUPEROFFICE_PANEL_ASSOCIATE_URL, SUPEROFFICE_PANEL_CONTACT_URL, SOCuID, Firma, "Vis Kontakt '" & Firma & "'")
													
                  									If ArrCalender( 6, Counter ) = 1 Then
                     									Kurskode = "Dagkurs"
                  									Else
                     									Kurskode = "Oppdrag"
                  									End If

													dag = "<td>" & oppdrag &"</td><td>"& Beskrivelse &"</td><td>"& valg &"</td><td>"& kurskode &" </td><td>"& program &"</td><td>"& Kompniva &"</td><td>"& Status &"</td><td>"& kunde &"</td>"

											ElseIf ArrCalender( 6, Counter ) = 2 Then
												bolKveld = true

												if oppdragID <> "" And oppdragID <> SokOppdragID Then
													oppdrag = CreateSOLink(SUPEROFFICE_XIS_TASK_URL, "", "oppdragvis.asp?oppdragID=" & OppdragID, OppdragID, "Vis Oppdrag" )
                  								else
													oppdrag = ""
												end if

												If Beskrivelse <> "" THEN
													Beskrivelse = CreateSOLink(SUPEROFFICE_XIS_TASK_URL, "", "OppdragvikarNy.asp?OppdragVikarID=" & OppdragVikarID, Beskrivelse, "Vis" )													
												else
													Beskrivelse = "Ledig"
												End If
												
												kunde = CreateSONavigationLink(SUPEROFFICE_PANEL_ASSOCIATE_URL, SUPEROFFICE_PANEL_CONTACT_URL, SOCuID, Firma, "Vis Kontakt '" & Firma & "'")

                  								If ArrCalender( 6, Counter ) = 1 Then
                     								Kurskode = "Dagkurs"
                  								Else
                     								Kurskode = "Oppdrag"
                  								End If

												kveld = "<td>" & oppdrag &"</td><td>"& Beskrivelse &"</td><td>"& valg &"</td><td>"& kurskode &" </td><td>"& program &"</td><td>"& Kompniva &"</td><td>"& Status &"</td><td>"& kunde &"</td>"
              							End If
									End If
								Next
							End If
							'her skrives data for dagoppdrag inn i tabellen
							%>
							<td>dag</td>
							<%
							if bolDag = true Then
     							Response.Write dag
							ElseIf SokOppdragID <> "" and NewDate >= strFraDato and bolDag=false and oppdragTid < 2 Then
								Beskrivelse = "<A href='oppdragvikarny.asp?OppdragID=" & SokOppdragID & "&Dato=" & NEWDATE & "&VIKARID=" & VikarID &"'>" & "<font color='green'>Ledig" & "</a>"
								valg = "<INPUT class='checkbox' TYPE='checkbox' name='chkBx' VALUE="& NEWDATE &" >"
								response.write "<td>&nbsp;</td><td>"& Beskrivelse &"</td><td>"& valg &"</td><td colspan=5'>&nbsp;</td>"
							Else
								response.write "<td>&nbsp;</td><td>Ledig</td><td>&nbsp;</td><td colspan='5'>&nbsp;</td>"
							end If
							'Her skrives data for kveldsoppdrag inn i tabellen
							Response.Write "<tr><td>kveld</td>"	'kveld
							if bolKveld = true Then
								Response.Write kveld
							ElseIf SokOppdragID <> "" and NewDate >= strFraDato and bolKveld=false and oppdragTid = 2 Then
								Beskrivelse = "<A href='oppdragvikarny.asp?OppdragID=" & SokOppdragID & "&Dato=" & NEWDATE & "&VIKARID=" & VikarID &"'>" & "<font color='green'>Ledig" & "</a>"
								valg = "<INPUT class='checkbox' TYPE='checkbox' name='chkBx' VALUE="& NEWDATE &" >"
								response.write "<td>&nbsp;</td><td>"& Beskrivelse &"</td><td>"& valg &"</td><td colspan='5'>&nbsp;</td>"
							Else
								response.write "<td>&nbsp;</td><td>Ledig</td><td colspan='6'>&nbsp;</td>"
							end If

							NewDate = DateAdd("d", 1, NewDate)
						Loop
						Response.Write "<tr><TD colspan=5></td>"
						Response.Write "<th>"
						If SokOppdragID<>"" Then
							Response.Write"<INPUT TYPE='SUBMIT' VALUE='Velg'>"
						End if
   						Response.Write"</th>"
						Response.Write "<TD colspan=5></td>"
						Response.Write "</table></form>"
					End If
					%>
				</div>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(objCon)
set objCon = nothing
%>