<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="includes/xis.rights.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="includes/Xis.Economics.Constants.inc"-->
<%

    dim strValueAdditionID
    dim strValueInvTotal
    dim strValueComment
	dim rsOppdragStatus
	dim oppdragVikarStatus
	dim faid
	dim transferFeeCount
	dim SOPeID
	dim lBestilltAv
	dim strKontaktperson
	Dim cts
	dim SOcuID
	
	If (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim rsOppdrag

	If Request.QueryString("FaId") <> "" Then
		faid = Request.QueryString("FaId")
	End If

	' Move parameter OppdragsVikarID to variable
	If Request.QueryString("OppdragVikarID") <> "" Then
		strOppdragVikarID = Request.QueryString("OppdragVikarID")
	Else
		' Do we have a OppdragID ?
		If Request.QueryString("OppdragID") <> "" Then
			OppdragID = Request.QueryString("OppdragID")
			VikarID = Request.QueryString("VikarID")
			strFradato = Request.QueryString("Dato")
			strTildato = Request.QueryString("Dato")
			lStatusID = 4
		Else
			AddErrorMessage("Feil:Parameter for oppdragid mangler!")
			call RenderErrorMessage()			
		End If
	End If

	' Move parameter Aksjon to variable
	If Request.QueryString("Aksjon") = "UTVID" Then
		strAksjon = Request.QueryString("Aksjon")
	ElseIf Request.QueryString("Aksjon") = "FORKORT" Then
		strAksjon = Request.QueryString("Aksjon")
		
	End If

	Set Conn = GetConnection(GetConnectionstring(XIS, ""))
	
	If strOppdragVikarID <> "" Then

		' Get oppdragsvikar data
		strSQL = "SELECT " & _
				"OPPDRAG_VIKAR.OppdragVikarID, " & _
				"OPPDRAG_VIKAR.OppdragID, " & _
				"OPPDRAG_VIKAR.VikarID, " & _
				"VIKAR.Fornavn, " & _
				"VIKAR.Etternavn, " & _
				"VIKAR.TypeID, " & _
				"OPPDRAG_VIKAR.Faktor, " & _
				"OPPDRAG_VIKAR.StatusID, " & _
				"OPPDRAG_VIKAR.FraDato, " & _
				"OPPDRAG_VIKAR.TilDato, " & _
				"OPPDRAG_VIKAR.FraKl, " & _
				"OPPDRAG_VIKAR.TilKl, " & _
				"OPPDRAG_VIKAR.Timeloenn, " & _
				"OPPDRAG_VIKAR.Timepris, " & _
				"OPPDRAG_VIKAR.AntTimer, " & _
				"OPPDRAG_VIKAR.Timeliste, " & _
				"OPPDRAG_VIKAR.Lunch, " & _
				"OPPDRAG_VIKAR.Notat, " & _
				"OPPDRAG_VIKAR.direkteTelefon, " & _
				"OPPDRAG_VIKAR.jobbEpost, " & _
				"OPPDRAG_VIKAR.CategoryId, " & _
				"VIKAR_ANSATTNUMMER.ansattnummer " & _
				"FROM OPPDRAG_VIKAR " & _
				"LEFT OUTER JOIN VIKAR ON OPPDRAG_VIKAR.Vikarid = VIKAR.Vikarid " & _			
				"LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON OPPDRAG_VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " & _
				"WHERE OPPDRAG_VIKAR.Oppdragvikarid = '" & strOppdragVikarID & "' "

			set rsOppdragVikar = GetFirehoseRS(StrSQL, Conn)
		  
			' Move to local variables
			OppdragVikarID		= rsOppdragVikar("Oppdragvikarid")
			OppdragID			= rsOppdragVikar("OppdragID")
			lStatusID			= rsOppdragVikar("StatusID")
			strFraDato			= rsOppdragVikar("FraDato")
			strTilDato			= rsOppdragVikar("TilDato")
			strVikarID			= rsOppdragVikar("VikarID")
			strFraKl			= FormatDateTime( rsOppdragVikar("FraKl"), 4)
			strTilKl			= FormatDateTime( rsOppdragVikar("Tilkl"), 4)
			strVikarNavn			= rsOppdragVikar("Fornavn") & " " & rsOppdragVikar("Etternavn")
			strTimeloenn			= FormatNumber(rsOppdragVikar("Timeloenn"),2)			
			strTimePris			= FormatNumber(rsOppdragVikar("Timepris"),2)
			strNotat			= Trim( rsOppdragVikar("Notat") )
			strAntTimer			= rsOppdragVikar("AntTimer")
			lTimeliste			= CLng( rsOppdragVikar("Timeliste") )
			Lunsj				= FormatDateTime( rsOppdragVikar("Lunch"), 4)
			strDirTlf			= rsOppdragVikar("direkteTelefon")
			strJobbEpost		= rsOppdragVikar("jobbEpost")
			strVikarStatus		= rsOppdragVikar("TypeID")
			strFaktor			= Mid( CStr( rsOppdragVikar("Faktor") ), 1, 4 )
			strAnsattnummer		= rsOppdragVikar("ansattnummer").Value
			lCategoryId 		= rsOppdragVikar("CategoryID")

		' Close and release recordset
		rsOppdragVikar.Close
		Set rsOppdragVikar = Nothing

	End If

	'''STH bug fix start
	strSQL = "select  count(*) as count , max(case isnull(statusid,0) " & _ 
		 "when 4 then 1 " & _
		 "else 0 end) as status " & _
		 "from OPPDRAG_VIKAR where oppdragid=" & OppdragID
		 
	set rsOppdragStatus = GetFirehoseRS(strSQL, Conn)
	
	if (cint(rsOppdragStatus("count"))=0) then
		oppdragVikarStatus = 0	
	else
		oppdragVikarStatus = cint(rsOppdragStatus("status"))
	end if
	
	rsOppdragStatus.Close
	set rsOppdragStatus = Nothing
	''' End

	' Get oppdrag
	strSQL = "SELECT O.OppdragID, S.Oppdragsstatus, O.Kurskode,O.SOPeID,O.BestilltAv,Kontaktperson = K.Fornavn + ' ' + K.Etternavn, O.FirmaID, F.Firma,F.SOCUID, O.FraDato, O.TilDato, O.FraKl, O.TilKl, O.Beskrivelse, O.Timerprdag, O.Lunch, O.Timepris " &_
			"FROM Oppdrag AS O, Firma AS F, H_OPPDRAG_STATUS AS S, KONTAKT AS K " &_
			"WHERE O.OppdragID = " & OppdragID &_
			" AND O.FirmaID *= F.FirmaID " &_
			" AND O.StatusID = S.OppdragsStatusID " &_
			" AND O.Bestilltav *= K.KontaktID "

	set rsOppdrag = GetFirehoseRS(StrSQL, Conn)

	' Move from recordset to variables
	OppdragID				= rsOppdrag("OppdragID")
	strFirma				= rsOppdrag("Firma")
	strFirmaID				= rsOppdrag("FirmaID")
	strOppdragStatus		= rsOppdrag("OppdragsStatus")
	strOppdragFraDato		= rsOppdrag("FraDato")
	strOppdragTilDato		= rsOppdrag("TilDato")
	strOppdragFraKl			= FormatDateTime( rsOppdrag("FraKl"), 4)
	strOppdragTilKl			= FormatDateTime( rsOppdrag("Tilkl"), 4)
	strOppdragBeskrivelse	= rsOppdrag("Beskrivelse")
	strOppdragLunsj			= FormatDateTime( rsOppdrag("Lunch"), 4)
	strOppdragAntTimer		= rsOppdrag("Timerprdag")
	strOppdragTimepris		= rsOppdrag("Timepris")
	strKurskode				= rsOppdrag("Kurskode")
	SOPeID					= rsOppdrag("SOPeID")
	lBestilltAv				= rsOppdrag("BestilltAv")
	SOcuID					= rsOppdrag("SOCUID")
	
	strKontaktperson	= rsOppdrag("Kontaktperson")
		if(lenb(strKontaktperson) > 0) then
			SOPeID = 0
		else
			lBestilltAv = 0
		end if		
	
	If strAksjon = "UTVID" Then
	  strHeading		= "Utvid oppdrag " & rsOppdrag("Firma")
	  strFraDato		= Dateadd("d", 1, strTilDato )
	  strTilDato		= ""
	  lStatusID		= 4
	ElseIf strAksjon = "FORKORT" Then
	  strHeading	= "Forkort oppdrag " & rsOppdrag("Firma")
	End If

	If strVikarID = "" Then

		strVikar = "SELECT " & _
				"VIKAR.Etternavn , " & _
				"VIKAR.Fornavn, " & _
				"VIKAR.Loenn1, " & _
				"VIKAR.TypeID, " &_
				"VIKAR_ANSATTNUMMER.ansattnummer " & _
				"FROM VIKAR " & _
				"LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " & _
				"WHERE VikarID = " & VikarID

		set rsVikar = GetFirehoseRS(strVikar, Conn)
		strVikarNavn = rsVikar("Fornavn") & " " & rsVikar("Etternavn")
		strVikarID = VikarID
		VikarType = rsVikar("TypeID")
		strAnsattnummer = rsVikar("ansattnummer").Value
		rsVikar.close
		set rsVikar = nothing

		strFraKl   = strOppdragFrakl
		strTilKl   = strOppdragTilkl
		strAntTimer = strOppdragAntTimer
		Lunsj		= strOppdragLunsj
		strTimePris = strOppdragTimePris

		strSQL = "SELECT [Timeloenn] FROM [Oppdrag] WHERE [Oppdragid] =" & OppdragID
		set rsOppdrag = GetFirehoseRS(strSQL, Conn)

		strTimeloenn = rsOppdrag("Timeloenn")
		rsOppdrag.close
		set rsOppdrag = nothing

		' Find faktor
		If Vikartype = 1 And strTimeloenn > 0 And strTimepris > 0 Then
			' Ansatt
			strFaktor = CDbl( strTimepris / strTimeloenn )
		ElseIf Vikartype > 1 And strTimeloenn > 0 And strTimepris > 0 Then
			' Selvstendig og AS
			strFaktor = CDbl( strTimepris / ( strTimeloenn / XIS_FACTOR ) )
			strFaktor = FormatNumber(strFaktor,2)
		Else
			strFaktor = 0
			strFaktor = FormatNumber(strFaktor,2)
		End If
	End If

	' Create title and heading in page
	If OppdragID <> "" Then
		strHeading = "Endre oppdragsvikar " & rsOppdrag("Firma")
	Else
		strHeading = "Ny oppdragsvikar " & rsOppdrag("Firma")
	End If

	' Close and release recordset
	rsOppdrag.Close
	Set rsOppdrag = Nothing

	'her legges mulighet inn for å oppdatere dirTelefon til vikar på aksepterte oppdrag
	if Request.QueryString("nyTlf")="ja" Then
		strTlf		= Request.Form("tbxDirTlf")
		strEmail	= Request.Form("tbxWorkEmail")
		strCategory	= Request.Form("tbxVikarCategory")
		
		if strCategory <> "" then
			strNyTlf	= "UPDATE [OPPDRAG_VIKAR] SET [direkteTelefon] = '"& strTlf &"',[CategoryId] = " & strCategory & ",[JobbEpost] = '"& strEmail &"' WHERE [OppdragVikarID] = "& strOppdragVikarID
		else
			strNyTlf	= "UPDATE [OPPDRAG_VIKAR] SET [direkteTelefon] = '"& strTlf &"',[JobbEpost] = '"& strEmail &"' WHERE [OppdragVikarID] = "& strOppdragVikarID
		end if 

		'strNyTlf	= "UPDATE [OPPDRAG_VIKAR] SET [direkteTelefon] = '"& strTlf &"',[CategoryId] = " & strCategory & ",[JobbEpost] = '"& strEmail &"' WHERE [OppdragVikarID] = "& strOppdragVikarID
		'Response.Write(strNyTlf)
		If ExecuteCRUDSQL(strNyTlf, Conn) = false then
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Oppdatering av telefon nummer feilet.")
			call RenderErrorMessage()
		End if		
		response.redirect "WebUI/OppdragView.aspx?OppdragID="&OppdragID
	End If
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<title><%=strHeading %></title>
		<script language="javascript" src="/xtra/Js/javascript.js" type="text/javascript"></script>
		<script language='javascript' src='/xtra/js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='/xtra/js/navigation.js' id='navigationScripts'></script>
		<script language="javaScript" type="text/javascript">
		
		
		   function SetDiv()
         {
            
          if(document.getElementById("TransferTemps")!=null)
          {
               if( document.getElementById("TransferTemps").checked )
           {
             document.getElementById("TransferDiv").style.display = "Block"; 
             
           }
           else 
           {
            document.getElementById("TransferDiv").style.display = "None"; 
            
            
            
            <%
			' Get Transfer Fee, if added
				StrSQL = "SELECT AdditionID,InvRate,InvTotal,Comment FROM ADDITION WHERE OppdragID = " & OppdragID & " AND ArticleID = 6"
				strValueAdditionID = 0
				set rsCategory = GetFirehoseRS(StrSQL, Conn)
				Do Until rsCategory.EOF											
					strValueAdditionID = rsCategory("AdditionID")
					strValueInvTotal = rsCategory("InvTotal")
					strValueComment = rsCategory("Comment")
					rsCategory.MoveNext
				Loop										
				rsCategory.Close
				Set rsCategory = Nothing			
			%>								 			
 		
 			<% 			
 			
 			If strValueAdditionID > 0 Then
 			
 			%>
 			
 			document.getElementById("tbxComment").value  = "<%=strValueComment%>"; 
            document.getElementById("tbxComment").disabled  = true; 
            document.getElementById("tbxAmount").value  = "<%=strValueInvTotal%>"; 
            document.getElementById("tbxAmount").disabled  = true; 
            document.getElementById("tbxAdditionID").value  = "<%=strValueAdditionID%>"; 
														
			<%
			Else
			%>
			
			document.getElementById("tbxComment").value  = "";             
            document.getElementById("tbxAmount").value  = "";             
            document.getElementById("tbxAdditionID").value  = "0"; 
				
			<%	
			End If  
			%>            
            
          
           }
          }	 
          
           
          }
                
			function Faktor()
			{
				var status = document.forms[0].elements['tbxVikarStatus'].value;
				var loenn = document.forms[0].elements['tbxTimeloenn'].value;
				pris = document.forms[0].elements['tbxTimepris'].value;
				var faktor = document.forms[0].elements['tbxFaktor'].value;
				var xis_factor = document.forms[0].elements['tbxXis_factor'].value;

				if (status == 1 && pris !=0 && loenn !=0 )
				{
					faktor = parseFloat(pris) / parseFloat(loenn);
					faktor=faktor.toString().substring(0,4);
					document.forms[0].elements['tbxFaktor'].value = faktor;
				}
				else if (status !=1 && pris !=0 && loenn !=0) 
				{
					faktor =  (parseFloat(pris)/(parseFloat(loenn)/parseFloat(xis_factor)));
					faktor=faktor.toString().substring(0,4);
					document.forms[0].elements['tbxFaktor'].value = faktor;
				}else
				document.forms[0].elements['tbxFaktor'].value="0";
			}
			
			function shortKey(e) 
			{	
				var keyChar = String.fromCharCode(event.keyCode);
				var modKey = event.ctrlKey;
				var modKey2 = event.shiftKey;
				
				if (modKey && modKey2 && keyChar=="S")
				{	
					parent.frames[funcFrameIndex].location=("/xtra/OppdragSoek.asp");
				}
			}
			
			function OnSave()
			{
				var selStatus = document.forms[0].elements['dbxStatus'].value;				
				<% 
				If lStatusID <> 4 Then
				%>
					if (<%=oppdragVikarStatus%> && selStatus == 4)
					{
						
						//alert("Kun en vikar kan knyttes til et oppdrag");
						
						alert("Et oppdrag kan bare ha en tilknyttet vikar med status 'aksept'. Dette oppdraget har allerede en tilknyttet vikar med status 'aksept'.");
						return false;
					}
					else
					{
						<% If faid > 0 Then%>
							if(document.forms[0].elements['dbxCategory'].value == 0 && selStatus == 4)
							{
								alert("Vennligst velg gyldig Kategori");
								return false;
							}
						<%End If %>
					}
				<%Else%>
					<% If faid > 0 Then%>
						if(document.forms[0].elements['dbxCategory'].value == 0 && selStatus == 4)
						{
							alert("Vennligst velg gyldig Kategori");
							return false;
						}
					<%End If %>
				<%End If %>
				return true;
			}
			
			function OnOppdater()
			{
				var selStatus = document.forms[0].elements['dbxStatus'].value;				
				<% 
				If lStatusID <> 4 Then
				%>
					if (<%=oppdragVikarStatus%> && selStatus == 4)
					{
						
						alert("Kun en vikar kan knyttes til et oppdrag");
						return false;
					}
					else
					{
						<% If faid > 0 Then%>
							if(document.forms[0].elements['dbxCategory'].value == 0 && selStatus == 4)
							{
								alert("Vennligst velg gyldig Kategori");
								return false;
							}
						<%End If %>
					}
				<%Else%>
					<% If faid > 0 Then%>
						if(document.forms[0].elements['dbxCategory'].value == 0 && selStatus == 4)
						{
							alert("Vennligst velg gyldig Kategori");
							return false;
						}
					<%End If %>
				<%End If %>
				document.oppdrVikTlf.elements['tbxVikarCategory'].value = document.forms[0].elements['dbxCategory'].value;
				return true;
			}
			
			//her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;			
		</script>		
	</head>
	<body onLoad="fokus();SetDiv();">
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1><%=strHeading %></h1>
			</div>
			<div class="content">				
				<table cellspacing="1" cellpadding="3">
					<tr>
						<th>KontaktNr:</th>
						<td><%=strFirmaID%></td>
						<th>Kontakt:</th>
						<td COLSPAN="5"><%=strFirma%></td>
					</tr>
					<tr>
						<th>Oppdragnr:</th>
						<td><%=OppdragID%></td>
						<th>Beskrivelse:</th>
						<td COLSPAN="5"><%=strOppdragBeskrivelse%></td>
					</tr>
					<tr>
						<th>StartDato:</th>
						<td><%=strOppdragFraDato%></td>
						<th>Klokken:</th>
						<td><%=strOppdragFrakl%></td>
						<th>SluttDato:</th>
						<td><%=strOppdragTilDato%></td>
						<th>Klokken:</th>
						<td><%=strOppdragTilkl%></td>
					</tr>
					<tr>
						<th>Status:</th>
						<td colspan="7"><%=strOppdragStatus%></td>
					</tr>
					<tr>
						<th>Ansattnummer:</th>
						<td><%=strAnsattnummer%></td>
						<th>Navn:</th>
						<td colspan="5"><%=strVikarNavn%></td>
					</tr>
				</table>
			</div>
			<form name="OPPDRAGVIKAR" Action="OppdragVikarDB.asp" method="post">
				<input name=tbxOppdragVikarID type=HIDDEN value=<%=strOppdragVikarID %> id="Hidden1">
				<input name=tbxOppdragID type=HIDDEN value=<%=OppdragID %> id="Hidden2">
				<input name=tbxVikarID type=HIDDEN value=<%=strVikarID %> id="Hidden3">
				<input name=tbxFirmaID type=HIDDEN value=<%=strFirmaID %> id="Hidden4">
				<input name=tbxKurskode type=HIDDEN value=<%=strKurskode %> id="Hidden5">
				<input name=tbxAksjon type=HIDDEN value=<%=strAksjon%> id="Hidden6">
				<input name="tbxXis_factor" type="HIDDEN" size="4" MAXLENGTH="6" VALUE="<%=XIS_FACTOR%> id="Hidden7">			
				<% 
				
				If lTimeliste < 1 or strAksjon = "UTVID" Then 		
					%>
					<div class="contentHead"><h2>Utvid oppdrag</h2></div>
					<div class="content">
						<table>
							<tr>
								<td>Dato:</td>
									<td nowrap><input class="mandatory" name="tbxFraDato" TYPE=TEXT SIZE=8 MAXLENGTH=8 VALUE="<%=strFraDato%>" ONBLUR="dateCheck(this.form, this.name)">
									-
									<input class="mandatory" name="tbxTilDato" TYPE=TEXT SIZE=8 MAXLENGTH=8 VALUE="<%=strTilDato%>" ONBLUR="dateCheck(this.form, this.name), dateInterval(this.form, this.name)">
								</td>
								<td>Klokken:</td>
								<td nowrap><input ID=lnk2 name="tbxFraKl"  TYPE=TEXT SIZE=4 MAXLENGTH=5 VALUE="<%=strFrakl%>" ONchange="timeCheck(this.form, this.name), workTime(this.form, this.name)">
									-
									<input ID=lnk3 name="tbxTilKl" TYPE=TEXT SIZE=4 MAXLENGTH=5 VALUE="<%=strTilkl%>" onchange="timeCheck(this.form, this.name), workTime(this.form, this.name)">
								</td>
								<td nowrap>Lunsj:</td>
								<td><input ID=lnk4 name=tbxLunsj TYPE=TEXT SIZE=5 MAXLENGTH=9 VALUE="<%=Lunsj%>" onchange="timeCheck(this.form, this.name), workTime(this.form, this.name)"></td>
								<td nowrap>Timer pr.dag:</td>
								<td><input name="tbxTimerPrDag" class="mandatory" TYPE=TEXT SIZE=4 MAXLENGTH=9 VALUE="<%=strAntTimer%>"</td>
							</tr>			
							<tr>
								<td>Status:</td>
								<td>
									<select name="dbxStatus" ID=lnk6 >
										<option VALUE="0"></option>
										<% 
										' Get Status
										StrSQL = "SELECT OppdragVikarStatusID, Status FROM h_oppdrag_vikar_status"
										set rsStatus = GetFirehoseRS(StrSQL, Conn)
										Do Until rsStatus.EOF
											If rsStatus("OppdragVikarStatusID") = lStatusID Then
												strValueSelected = rsStatus("OppdragVikarStatusID") & " SELECTED"
											Else
												strValueSelected = rsStatus("OppdragVikarStatusID")
											End If %>
											<option value=<%=strValueSelected %>><%=rsStatus("Status") %></option>
											<%  
											rsStatus.MoveNext
										Loop
										' Close and release recordset
										rsStatus.Close
										Set rsStatus = Nothing
										%>
									</select>
								</td>
								<td>Timelønn:</td>
								<td><input name="tbxTimeloenn" onChange="Faktor()" TYPE=TEXT SIZE=5 MAXLENGTH=9 VALUE="<%=strTimeloenn%>" </td>
								<td>Timepris:</td>
								<td><input name="tbxTimepris" onChange="Faktor()" TYPE=TEXT SIZE=5 MAXLENGTH=9 VALUE="<%=strTimePris%>" </td>
								<td>Faktor:</td>
								<td><input name="tbxFaktor" TYPE=TEXT SIZE=4 MAXLENGTH=6 VALUE="<%=strFaktor%>" </td>
								<td><input name="tbxVikarStatus" TYPE=HIDDEN SIZE=4 MAXLENGTH=6 VALUE="<%=strVikarStatus%>" </td>
							</tr>
							<tr>
							<td>Kontaktperson:</td>
							<td colspan="8">
							<%

							Set oXmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
							oXmlHTTP.Open "POST", "http://xis/Xtra/CRMIntegration/Responder.aspx?PageMode=GetContacts&Socuid=" + Cstr(SOCuID) + "&Sopeid=" +  Cstr(SOPeID), False , Application("Xtra_Domain") +"\" + Application("Xtra_User"), Application("Xtra_Password")

							if (clng(SOPeID) > 0 OR clng(lBestilltAv) = 0) then												
										

								oXmlHTTP.send ""
								Response.Write "<select name='dbxKontaktP' class='mandatory' ><option value='0'>(Ingen valgt)</option>"
								Response.Write oXmlHTTP.responseText
							end if
								if (clng(lBestilltAv) > 0) then
									%>
									<%=strKontaktperson%><input type="hidden" value="<%=lBestilltAv%>" name="dbxKontaktP" id="dbxKontaktP">
									<input type="hidden" value="<%=strKontaktperson%>" name="dbxKontaktName" id="dbxKontaktName">
									<%
								end if
							%>
							
							</td>
							<tr>
							<tr>
								<td>Direkte tlf:</td>
								<td COLSPAN="7"><input ID=lnk10 name="tbxDirTlf" TYPE=TEXT SIZE=8 MAXLENGTH=8 VALUE="<% =strDirTlf %>"></td>
							</tr>
							<tr>
								<td>Jobb E-post:</td>
								<td colspan="7"><input name="tbxWorkEmail" type="text" size="16" maxlength="256" value="<%=strJobbEpost%>" id="Text1"></td>
							</tr>														
							<% If faid > 0 Then%>
							<tr>
							<td >Kategori:</td>
							<td>
							<select name="dbxCategory" id="dbxCategory" class="form_list_menu">
								<option value="0">- None -</option>
								<%
     								' Get Categories
										StrSQL = "SELECT F.* FROM FrameworkCategory F WHERE F.faid = " & faid & " ORDER BY F.CategoryCode" 
										set rsCategory = GetFirehoseRS(StrSQL, Conn)
										Do Until rsCategory.EOF
											If rsCategory("CategoryId") = lCategoryId Then
												strValueSelected = rsCategory("CategoryID") & " SELECTED"
     											Else
												strValueSelected = rsCategory("CategoryID")
											End If %>
											<option value=<%=strValueSelected %>><%=rsCategory("CategoryCode") %> -
											<% if Not isNull(rsCategory("CategoryName")) Then %>
												<%=rsCategory("CategoryName") %>
											<%End if%>
											
											</option>
											<%  
											rsCategory.MoveNext
										Loop
										' Close and release recordset
										rsCategory.Close
										Set rsCategory = Nothing
								%>
							</select>
							</td>
							</tr>   													
							<%End If%>   													
							<tr>
								<td>Notat:</td>
								<TD COLSPAN=7><textarea name="tbxNotat" ROWS="2" COLS="60"><%=strNotat%></textarea></td>
							</tr>
						</table>
					<%
				elseif strAksjon="FORKORT" Then 
					%>
					<div class="contentHead"><h2>Forkort oppdrag</h2></div>
					<div class="content">
						<table>
							<tr>
								<td>Dato:</td>
								<td nowrap>
									<input name="tbxFraDato" class="mandatory" type=text SIZE=8 MAXLENGTH=8 VALUE="<%=strFraDato%>" ONBLUR="dateCheck(this.form, this.name), dateInterval(this.form, this.name)" >
									-
									<input name="tbxTilDato" class="mandatory" type=text SIZE=8 MAXLENGTH=8 VALUE="<%=strTilDato%>" ONBLUR="dateCheck(this.form, this.name), dateInterval(this.form, this.name)">
								</td>
								<td>Klokken:</td>
								<TD nowrap>
									<input name="tbxFraKl" disabled type=text SIZE=4 MAXLENGTH=5 VALUE="<%=strFrakl%>" ONBLUR="timeCheck(this.form, this.name), dateInterval(this.form, this.name)">
									-
									<input name="tbxTilKl" disabled type=text SIZE=4 MAXLENGTH=5 VALUE="<%=strTilkl%>" ONBLUR="timeCheck(this.form, this.name)">
								</td>
								<TD nowrap>Lunsj:</td>
								<td><input name=tbxLunsj disabled type=text SIZE=5 MAXLENGTH=9 VALUE="<%=Lunsj%>"></td>
								<TD nowrap>Timer pr.dag:</td>
								<td><input name=tbxTimerPrDag disabled type=text SIZE=5 MAXLENGTH=9 VALUE="<%=strAntTimer%>"></td>
							</tr>							
							<tr>
								<td>Status:</td>
								<td>
									<select name="dbxStatus" disabled>
										<option VALUE="0"></option>
										<% 
										' Get Status
										StrSQL = "SELECT OppdragVikarStatusID, Status from h_oppdrag_vikar_status"
										set rsStatus = GetFirehoseRS(StrSQL, Conn)
										Do Until rsStatus.EOF
											If rsStatus("OppdragVikarStatusID") = lStatusID Then
												strValueSelected = rsStatus("OppdragVikarStatusID") & " SELECTED"
											Else
												strValueSelected = rsStatus("OppdragVikarStatusID")
											End If  
											%>
											<OPTION VALUE=<%=strValueSelected %>><%=rsStatus("Status") %></option>
											<%  
											rsStatus.MoveNext
										Loop

										' Close and release recordset
										rsStatus.Close
										Set rsStatus = Nothing
										%>
									</select>
								</td>												
								<td>Timelønn:</td>
								<td><input name="tbxTimeloenn" disabled type=text SIZE=4 MAXLENGTH=6 VALUE="<%=strTimeloenn%>"></td>
								<td>TimePris:</td>
								<td><input name="tbxTimepris" disabled type=text SIZE=4 MAXLENGTH=6 VALUE="<%=strTimepris%>"></td>
								<td>Faktor:</td>
								<td><input name="tbxFaktor" disabled type=text SIZE=4 MAXLENGTH=6 VALUE="<%=strFaktor%>" </td>
								<td><input name="tbxVikarStatus" type=HIDDEN SIZE=4 MAXLENGTH=6 VALUE="<%=strVikarStatus%>" </td>
							</tr>
							<tr>
								<td>Direkte tlf:</td>
								<TD COLSPAN=7><input name="tbxDirTlf" disabled type=text SIZE=8 MAXLENGTH=8 VALUE="<%=strDirTlf%>"></td>
							</tr>
							<% If faid > 0 Then%>
							<tr>
							<td >Kategori:</td>
							<td>
							<select name="dbxCategory" id="dbxCategory" class="form_list_menu">
								<option value="0"> - None -</option>
								<%
									' Get Categories
										StrSQL = "SELECT F.* FROM FrameworkCategory F WHERE F.faid = " & faid & " ORDER BY F.CategoryCode"
										set rsCategory = GetFirehoseRS(StrSQL, Conn)
										Do Until rsCategory.EOF
											If rsCategory("CategoryId") = lCategoryId Then
												strValueSelected = rsCategory("CategoryID") & " SELECTED"
     									Else
												strValueSelected = rsCategory("CategoryID")
											End If %>
											<option value=<%=strValueSelected %>><%=rsCategory("CategoryCode") %> -
											<% if Not isNull(rsCategory("CategoryName")) Then %>
												<%=rsCategory("CategoryName") %>
											<%End if%>
											</option>
											<%  
											rsCategory.MoveNext
										Loop
										' Close and release recordset
										rsCategory.Close
										Set rsCategory = Nothing
								%>
							</select>
							</td>
							</tr>  
							<%End If%>
							<tr >
							
                      
								
							<TD COLSPAN=8>
							
							<table border="0" cellpadding ="0" cellpadding ="0" width ="100%">
								<tr>
									<td style="width:5%">							
								  		<INPUT TYPE=CHECKBOX NAME="TransferTemps" id ="TransferTemps" onclick ="SetDiv()">
									</td>
								 	<td style="width:25%">Overgangsbeløp</td>
								 	<td style="width:70%">
								 		<div id ="TransferDiv" style="display:none" >  
								 				<input name="tbxAdditionID" id="tbxAdditionID" type="hidden" />
												<table border="0" cellpadding ="0" cellpadding ="0" width ="100%">
													<tr>
														<td>&nbsp;&nbsp;&nbsp;Fakturakommentar : </td>
														<td>
														<input maxlength="250" name="tbxComment" size="50" type="text" />														
														
														</td>
														<td>Beløp :	</td>
														<td><input maxlength="10" name="tbxAmount" size="10" type="text" />
									    				</td>
													</tr>
												</table>												
											
							 			</div>
								 	</td>
								</tr>
							</table>
								
								            
						 </td>
							</tr>  
							<tr>
								<td>Kommentar 						<TD COLSPAN=7><textAREA name="tbxNotat" ROWS=2 COLS=60 disabled><%=strNotat%></textAREA></td>
							</tr>
						</table>
					<%
				else
					%>
					<div class="contentHead"><h2>Vis oppdrag</h2></div>
					<div class="content">
						<table>
							<tr>
								<td>Dato:</td>
								<TD nowrap>
									<input name="tbxFraDato" disabled type=text size=8 MAXLENGTH=8 VALUE="<%=strFraDato%>" ONBLUR="dateCheck(this.form, this.name), dateInterval(this.form, this.name)" >
									-
									<input name="tbxTilDato" disabled type=text size=8 MAXLENGTH=8 VALUE="<%=strTilDato%>" ONBLUR="dateCheck(this.form, this.name), dateInterval(this.form, this.name)">
								</td>
								<td>Klokken:</td>
								<TD nowrap>
									<input name="tbxFraKl" disabled type=text size=4 MAXLENGTH=5 VALUE="<%=strFrakl%>" ONBLUR="timeCheck(this.form, this.name), dateInterval(this.form, this.name)">
									-
									<input name="tbxTilKl" disabled type=text size=4 MAXLENGTH=5 VALUE="<%=strTilkl%>" ONBLUR="timeCheck(this.form, this.name)">
								</td>
								<TD nowrap>Lunsj:</td>
								<td>
									<input name=tbxLunsj disabled type="text" size="5" MAXLENGTH="9" VALUE="<%=Lunsj%>">
								</td>
								<TD nowrap>Timer pr.dag:</td>
								<td>
									<input name=tbxTimerPrDag disabled type="text" size="5" MAXLENGTH="9" VALUE="<%=strAntTimer%>">
								</td>
							</tr>
							<tr>
								<td>Status:</td>
								<td>
									<select name="dbxStatus" disabled>
										<option VALUE="0"></option>
											<%
											' Get Status
											StrSQL = "SELECT OppdragVikarStatusID, Status FROM h_oppdrag_vikar_status"
											set rsStatus = GetFirehoseRS(StrSQL, Conn)
											Do Until rsStatus.EOF
												If rsStatus("OppdragVikarStatusID") = lStatusID Then
													strValueSelected = rsStatus("OppdragVikarStatusID") & " selectED"
												Else
													strValueSelected = rsStatus("OppdragVikarStatusID")
												End If %>
												<option VALUE=<%=strValueSelected %>><%=rsStatus("Status") %></option>
												<%
												rsStatus.MoveNext
											Loop
											rsStatus.Close
											Set rsStatus = Nothing
											%>
									</select>
								</td>
								<td>Timelønn:</td>
								<td><input name="tbxTimeloenn" disabled type="text" size="4" MAXLENGTH="6" VALUE="<%=strTimeloenn%>"></td>
								<td>TimePris:</td>
								<td><input name="tbxTimepris" disabled type="text" size="4" MAXLENGTH="6" VALUE="<%=strTimepris%>"></td>
								<td>Faktor:</td>
								<td><input name="tbxFaktor" disabled type="text" size="4" MAXLENGTH="6" VALUE="<%=strFaktor%>" </td>
								<td><input name="tbxVikarStatus" type="HIDDEN" size="4" MAXLENGTH="6" VALUE="<%=strVikarStatus%>" </td>
								
							</tr>
							<% If faid > 0 Then%>
							<tr>
							<td >Kategori:</td>
							<td>
							<select name="dbxCategory" id="dbxCategory" class="form_list_menu">
								<option value="0"> - None - </option>
								<%

     								' Get Categories
										StrSQL = "SELECT F.* FROM FrameworkCategory F WHERE F.faid = " & faid & " ORDER By F.CategoryCode"
										set rsCategory = GetFirehoseRS(StrSQL, Conn)
										Do Until rsCategory.EOF
											If rsCategory("CategoryId") = lCategoryId Then
												strValueSelected = rsCategory("CategoryID") & " SELECTED"
     											Else
												strValueSelected = rsCategory("CategoryID")
											End If %>
											<option value=<%=strValueSelected %>><%=rsCategory("CategoryCode") %>- 
											<% if Not isNull(rsCategory("CategoryName")) Then %>
												<%=rsCategory("CategoryName") %>
											<%End if%>
											</option>
											<%  
											rsCategory.MoveNext
										Loop
										' Close and release recordset
										rsCategory.Close
										Set rsCategory = Nothing

								%>
							</select>
							</td>
							</tr>  
							<%End If%>  
							<tr>
								<td>Kommentar <br> til vikar:</td>
								<td COLSPAN=7><textarea name="tbxNotat" ROWS=2 COLS=60 disabled><%=strNotat%></textarea></td>
							</tr>
						</table>
						<%
					end If
					%>
					<table>
						<tr>
							<%
							' Show appropriate buttons depening on key value
							If lTimeliste < 1 or strAksjon = "UTVID" or strAksjon = "FORKORT" Then
								If strOppdragVikarID = "" Then
									' No key value: save new and clear
									Response.write "<td><input ID='lnk12' name='pbnDataAction' TYPE='SUBMIT' VALUE=' Lagre ' onclick='return OnSave()'></td>"
								Else													
									' Key value: save and clear
									Response.write "<td><input ID='lnk12' name='pbnDataAction' TYPE='SUBMIT' VALUE=' Lagre ' onclick='return OnSave()'></td>"
									If lTimeliste < 1 Then
										Response.write "<td><input name=pbnDataAction TYPE=SUBMIT VALUE=' Slette '></td>"
									End If
								End If
							End If
							%>
						</tr>
					</table>
				</div>
			</form>
			<%
			If not lTimeliste < 1 or strAksjon = "UTVID" or strAksjon = "FORKORT" Then
				%>
				<div class="contentHead"><h2>Vikar kontaktinformasjon</h2></div>
				<div class="content">
					<form name="oppdrVikTlf" ACTION="OppdragVikarNy.asp?OppdragVikarID=<%=strOppdragVikarID%>&nyTlf=ja" METHOD=POST>
						<table>
							<tr>
								<td>Direkte telefon: </th>
								<td><input name="tbxDirTlf" TYPE="text" size="8" maxlength="8" VALUE="<%=strDirTlf%>"></td>
								<td>Jobb E-post: </th>
								<td><input name="tbxWorkEmail" TYPE="text" size="16" maxlength="256" VALUE="<%=strJobbEpost%>"></td>
								<td><input name="tbxVikarCategory" type="HIDDEN" size="4" MAXLENGTH="6" VALUE=""</td>
								<td><input type=submit VALUE='Oppdater' onclick='return OnOppdater()'></td>
							</tr>
						</table>
					</form>
				</div>
				<%
			End If
			%>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>

