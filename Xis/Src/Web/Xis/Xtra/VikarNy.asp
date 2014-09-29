<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.Settings.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim lngVikarID
	dim iAarTeller
	dim iAar
	dim strPresentasjon
	dim strSelected
	dim strSelected2
	dim strSelected3
	dim strSelected4
	dim strSelected5
	dim strSelected6
	dim strSelected7
	dim intTilgjengelig
	dim strAccountNr
	dim strSQL

	dim hasCV
	dim isCVLocked
	dim cons
	dim cv

	dim nTeller

	'Consultant menu variables
	dim strClass
	dim strJSEvents
	dim strDisabled

	' Get a database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

	'Henter data for vikar ut fra vikarid
	lngVikarID = clng(Request("VikarID"))

	If lngVikarID > 0 Then
		strSQL = "select VikarID, Etternavn, Fornavn, foedselsdato, kjoenn, personnummer, " &_
		"loenn1, Ansmedid, typeid, statusid, IntervjuDato, Telefon, MobilTlf, kundepresentasjon, " &_
		"Fax, Epost, Hjemmeside," &_
		"kontraktsendt, kontraktmottatt, kurskode,WorkType,Link1URL,Link2URL,Link3URL,notat,InterestedJobs," &_
		"oppsummering_intervju, oppsummering_ref_sjekk, MottattSkattekort, " &_
		"Foererkort, Bil, hasCar, Country, Oppsigelsestid, bankkontonr " &_
		"FROM VIKAR " &_
		"WHERE VikarID = " & Request("VikarID")
		
		set rsVikar = GetFirehoseRS(strSQL, Conn)

		'Flytter verdier til variable
		lngVikarID				= rsVikar("VikarID").Value
		strEtternavn			= rsVikar("Etternavn").Value
		strFornavn				= rsVikar("Fornavn").Value
		strFoedselsdato			= rsVikar("Foedselsdato").Value
		strPersonnummer			= rsVikar("Personnummer").Value
		strMottattSkattekort	= rsVikar("MottattSkattekort").Value
		strNotat				= rsVikar("Notat").Value
		strInterestedJobs       = rsVikar("InterestedJobs").Value
		strOppsumIntervju		= rsVikar("Oppsummering_intervju").Value
		strOppsumRef			= rsVikar("Oppsummering_ref_sjekk").Value
		lTimeloenn				= rsVikar("loenn1").Value
		lAnsMedID				= rsVikar("AnsMedID").Value
		lTypeID					= rsVikar("TypeID").Value
		strStatus				= rsVikar("StatusID").Value
		strIntervjudato			= rsVikar("IntervjuDato").Value
		strTelefon				= rsVikar("Telefon").Value
		strFax					= rsVikar("Fax").Value
		strMobilTlf				= rsVikar("MobilTlf").Value
		strEPost				= rsVikar("EPost").Value
		strHjemmeside			= rsVikar("Hjemmeside").Value
		strKontraktSendt		= rsVikar("KontraktSendt").Value
		strKontraktmottatt		= rsVikar("Kontraktmottatt").Value
		strPresentasjon			= rsVikar("kundepresentasjon").Value
		strFoererkort			= rsVikar("Foererkort").Value
		strBil					= rsVikar("hasCar").Value
		strOppsigelsestid		= rsVikar("Oppsigelsestid").Value
		strWorkType			    = rsVikar("WorkType").Value
		strLink1URL			    = rsVikar("Link1URL").Value
		strLink2URL			    = rsVikar("Link2URL").Value
		strLink3URL			    = rsVikar("Link3URL").Value
		lCountryID			= rsVikar("Country").Value
		intTilgjengelig			= rsVikar("KursKode")
		strAccountNr			= rsVikar("bankkontonr")

		'kjønn
		If ucase(rsVikar("Kjoenn")) = "M" Then
			strMann = "CHECKED"
		ElseIf ucase(rsVikar("Kjoenn")) = "K" Then
			strKvinne = "CHECKED"
		End If

		'Lukker recordsett for vikaropplysninger
		rsVikar.Close
		Set rsVikar = Nothing

		'Henter data for adresse for vikar
		set rsAdresse = GetFirehoseRS("select AdrId, AdresseType, Adresse, Postnr, Poststed from ADRESSE where AdresseRelID = " & Request("VikarID") & " AND AdresseType=1" , Conn)

		lAdrID		= rsAdresse("AdrID")
		strAdress	= rsAdresse("Adresse")
		strPostnr   = rsadresse("Postnr")
		strPostSted = rsadresse("Poststed")

		'Lukker recordset for adresse
		rsAdresse.Close
		Set rsAdresse = Nothing

		set cons = Server.CreateObject("XtraWeb.Consultant")
		cons.XtraConString = Application("XtraWebConnection")
		cons.GetConsultant(lngVikarID)

		set cv	= cons.CV
		cv.XtraConString = Application("Xtra_intern_ConnectionString")
		cv.XtraDataShapeConString = Application("ConXtraShape")
		cv.Refresh

		if cons.CV.DataValues.Count = 0 then
			cons.CV.Save
		end if

		hasCV = true
		isCVLocked = cv.islocked

		set cv = nothing
		cons.CV.cleanup
		cons.cleanup
		set cons = nothing

		'Header for siden
		strHeading = "Endre vikar " & strFornavn & " " & strEtternavn

		endring = 1

	Else
		lngVikarID = 0
   		strHeading = "Ny vikar"
		strStatus = 1
		' Set default radiobutton
		strSelected1 = "CHECKED"
		endring = 2
		strEtternavn		= request("txtEtternavn")
		strFornavn			= request("txtFornavn")
		strFoedselsdato		= request("dtFodselsDato")
		hasCV				= false
		isCVLocked			= false
	End If

	'DNN Brukerhåndtering

	' - Retrieves AND stores DNN Publish  username in variable if it exists.

	'Variables used by DNN user routines.
	dim strUsername 'string
	Dim objUserDom
	Dim objUserProxy 
	Dim sUserXml
         

	'dersom vikar har status "ANSATT" (3), "SØKER" (1) eller "KANDIDAT" (2) pt.optatt(8)
	'if ((strStatus = "3") OR (strStatus = "2") OR (strStatus = "1") OR (strStatus = "8"))  then
		sUserServiceURL = Application("DNNUserServiceURL")

		iApp = Cint(Application("Application"))
		
		Set objUserProxy = Server.CreateObject("ECDNNUserProxy.UserServiceProxy")
	    objUserProxy.Url = sUserServiceURL

		sUserXml = objUserProxy.GetUser(iApp, lngVikarID,"V")

		if sUserXml <> "" then
			Set objUserDom = Server.CreateObject("Microsoft.XmlDom")
			objUserDom.LoadXml sUserXml
    		strUsername = objUserDom.selectSingleNode("/user/userName").Text
               
		end if
	'end if

'/IMP Brukerhåndtering
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
		<script language="javascript" type="text/javascript" src="/xtra/Js/ajax.js"></script>
		<script type="text/javascript" src="/xtra/js/CVFunctions.js"></script>
		<script type="text/javascript" src="/xtra/js/contentMenu.js"></script>
		<script language='javascript' src='/xtra/js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='/xtra/js/ecajax.js'></script>
		<script language="javaScript">
			var i=0;
			var hasCV = <%=lcase(hasCV)%>;
			var isCVLocked = <%=lcase(isCVLocked)%>;

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
			// her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;
		
			//kode for å sette fokus ved onLoad()
			function fokus()
			{

				var strVikarFunksjon = document.all.dbxStatus.options[document.all.dbxStatus.selectedIndex].text;
				var bSkattekortEndring = <%=LCase(Session("EndreSkattekort"))%>

				if (bSkattekortEndring == true && strVikarFunksjon == "Ansatt") 
				{
					document.all.skattekort.disabled = false;
					document.all.tbxSkattekortDisabled.value = "false";
				} 
				else 
				{
					document.all.skattekort.disabled = true;
					document.all.tbxSkattekortDisabled.value = "true";
				}

				ToggleCreateUser();
				if(document.all('lnk0'))
				{
					document.all('lnk0').focus();
				}
			}

			function ToggleTermOfNotice()
			{
				if (document.all.dbxStatus.value=='3')
				{
					document.all.cboOppsigelse.disabled = true;
				}
				else
				{
					document.all.cboOppsigelse.disabled = false;
				}

			}

			// Fred 03.08.00 IMP brukerhåndering
			function statusChange()
			{
				ToggleCreateUser();
				ToggleTermOfNotice();

				var bSkattekortEndring = <%=LCase(Session("EndreSkattekort"))%>

				if (bSkattekortEndring == true && ((document.all.dbxStatus.value=='2') ||(document.all.dbxStatus.value=='3')) )
				{
					document.all.skattekort.disabled = false;
					document.all.tbxSkattekortDisabled.value = "false";
				}
				else
				{
					document.all.skattekort.disabled = true;
					document.all.tbxSkattekortDisabled.value = "true";
				}
			};

			function ckChange()
			{
				if (document.all.chkCreateuser.checked == true)
				{
					document.all.Username.disabled=false;
					if (document.all.tbxEPost.value != "")
					{
						document.all.Username.value = document.all.tbxEPost.value;
					}
					document.all.Username.focus();
				}
				else
				{
					document.all.Username.value = '';
					document.all.Username.disabled = true;
				}
				VerifyEditCv();
			};


			//Enable web user creation only if status is 'ANSATT' (3) or 'KANDIDAT' (2) or 'Søker' (1)
			function ToggleCreateUser()
			{
					if (((document.all.dbxStatus.value=='2') || (document.all.dbxStatus.value=='3') || (document.all.dbxStatus.value=='1') || (document.all.dbxStatus.value=='8')) && (document.all.tbxEPost.value.length>6) )
					{
						EnableCreateUser(1);
					}
					else
					{
						//document.all.Username.value = '';
						EnableCreateUser(1);
					}
					VerifyEditCv();

			}

			function EnableCreateUser(CreateUser)
			{
					if (CreateUser == 0)
					{
						document.all.Username.disabled = true;
						document.all.chkCreateuser.checked = false;
						document.all.chkCreateuser.disabled = true;
					}
					else
					{
						
						if (document.all.Username.length > 0)
						{
						        document.all.chkCreateuser.disabled = true;
							document.all.chkCreateuser.checked = true;
							document.all.Username.disabled = false;
						}
						else if (document.all.Username.length == 0)
						{
						        document.all.chkCreateuser.disabled = false;
							document.all.Username.disabled = true;
							document.all.chkCreateuser.checked = false;
						}
					}
			}

			function VerifyEditCv()
			{
				/*if (document.all.chkCreateuser.checked == true)
				{
					document.all.chkEditCV.disabled = false;
				}
				else
				{
					document.all.chkEditCV.disabled = true;
					document.all.chkEditCV.checked = false;
				}*/
			}

			/*function ToggleEditCV()
			{
				if (document.all.chkCreateuser.checked == true)
				{
					if (hasCV == true && document.all.chkEditCV.checked == true)
					{
						return;
						// Kommentert ut ettersom denne boksen aldri kom opp riktig kjørt i en pane i SuperOffice
						var response = confirm("Du har valgt å åpne CVen til vikar for redigering.\nDersom du lagrer nå vil HTML formatering komme til å forsvinne\nnår brukeren redigerer CVen på web.\nFor å åpne CV''en trykk på 'OK', trykk på 'Avbryt' for å avbryte.");
						if(response == true)
						{
							return;
						}
						else if(response == false)
						{
							document.all.chkEditCV.checked = false;
							return;
						}
						
					}
				}
			}*/

			function ValidateBA(oFld)
			{
				var sValue = oFld.value
				if (sValue.length==0)
				{
					return
				}
				var Reg = /^\d{4}\.\d{2}\.\d{5}$/
				var RegSeq = /^\d{11}$/
				if (Reg.test(sValue) == false)
				{
					if (RegSeq.test(sValue) == false)
					{
						alert("Kontonummeret er på ugyldig format!");
						oFld.select();
						oFld.focus();
					}
					else
					{
						oFld.value = sValue.substr(0,4) + '.' + sValue.substr(4,2) + '.' + sValue.substr(6,5)
					};
				}
			}
			
			function PopulateCategories(CONTROL)
			{
				var strTemp = "";
				var url;
				for(var i = 0;i < CONTROL.length;i++)
				{
					if(CONTROL.options[i].selected == true)
						strTemp += CONTROL.options[i].value + ",";
				}
				if(strTemp != "")
					strTemp = strTemp.substr(0,strTemp.length - 1);
				
				url = "Callback.asp?FuncName=GetVikarCategories&TomIDs=" + strTemp + "&SelectedCategories=" + document.all.tbxSelectedCategories.value;
				asynchronousCall(url,handleCategories);
				
			}
	
		var handleCategories = function (result)
		{
			if(result) 
			{
				var str = "<select class='mandatory' id='VikarCategory' multiple name='VikarCategory' size='10' style='width:220px;'>";
				str = str.concat(result,"</SELECT>");
				document.getElementById('divVikarCategory').innerHTML = str;
			}
		};
		</script>
	</head>
	<body onLoad="fokus()">
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1><%=strHeading %></h1>
				<div class="contentMenu">
					<table cellpadding="0" cellspacing="0" width="96%" ID="Table1">
						<tr>
							<td>
								<table cellpadding="0" cellspacing="2" ID="Table2">
									<tr>
										<%
										if (lngVikarID>0) then
											strClass = "menu"
											strJSEvents = "onMouseOver=""menuOver(this.id);"" onMouseOut=""menuOut(this.id);"""
											strDisabled = ""
										else
											strClass = "menu disabled"
											strJSEvents = ""
											strDisabled = "disabled"
										end if
										%>
										<td class="menu" id="menu1" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
											<a onClick="javascript:document.all.frmVikar.submit();" href="#" title="Lagre vikar">
											<img src="/xtra/images/icon_save.gif" width="18" height="15" alt="" align="absmiddle">Lagre</a>
										</td>
										<td class="<%=strClass%>" id="menu2" <%=strJSEvents%>>
											<a href="/xtra/vikarvis.asp?vikarid=<%=lngVikarID%>" title="Vis vikar">
											<img src="/xtra/images/icon_consultant.gif" width="18" height="15" alt="" align="absmiddle">Vis</a>
										</td>
										<td class="menu disabled" id="menu3">
											<img src="/xtra/images/icon_changeInfo.gif" width="18" height="15" alt="" align="absmiddle">Endre
										</td>
										<td class="<%=strClass%>" id="menu4" <%=strJSEvents%>><strong>CV</strong>&nbsp;<select <%=strDisabled%> id="cboCVChoice" onChange="javascript:Vis_CV(<%=lngVikarID%>);" NAME="cboCVChoice"><option value="0"></option><option value="1">Se</option><%If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then%><option value="2">Endre</option><%end if%><option value="3">Presentere</option></select></td>
										 
										<td class="<%=strClass%>" id="menu6" <%=strJSEvents%>>
											<form ACTION="vikar-kunder.asp?VikarID=<%=lngVikarID%>" METHOD="POST" id="frmConsultantFormerClients">
												<a href="javascript:document.all.frmConsultantFormerClients.submit();" title="Vis tidligere oppdragsgivere"><img src="/xtra/images/icon_tidl-kunder.gif" alt="" width="18" height="15" border="0" align="absmiddle">Tidligere oppdragsgivere</a>
											</form>
										</td>
										<td class="<%=strClass%>" id="menu7" <%=strJSEvents%>>
											<form ACTION="AktivitetVikar.asp?VikarID=<%=lngVikarID%>" METHOD="POST" id="frmConsultantActivities">
												<a href="javascript:document.all.frmConsultantActivities.submit();" title="Vis aktiviteter for vikaren"><img src="/xtra/images/icon_activities.gif" alt="" width="18" height="15" border="0" align="absmiddle">Aktiviteter</a>
											</form>
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
				<form name="frmVikar" action="VikarDB.asp" method="post" ID="Form1">
					<input NAME="tbxVikarID" 	TYPE="HIDDEN" VALUE="<%=lngVikarID%>" ID="Hidden1">
					<input NAME="tbxStatus"		TYPE="HIDDEN" VALUE="<%=strStatus%>" ID="Hidden2">
					<input NAME="tbxAdrID"		TYPE="HIDDEN" Value="<%=lAdrID%>" ID="Hidden3">
					<input NAME="tbxAdrType"	TYPE="HIDDEN" Value="1" ID="Hidden4">
					<input NAME="endring"		TYPE="HIDDEN" VALUE="<% =endring %>" ID="Hidden5">
					<input type="hidden" id="tbxSkattekortDisabled" name="tbxSkattekortDisabled" value="">
					<table class="layout" cellpadding="0" cellspacing="0" ID="Table3">
						<col width="33%">
						<col width="34%">
						<col width="33%">
						<tr>
							<td>
								<table ID="Table4">
									<tr>
										<td>Fornavn:</td>
										<td><input class="mandatory" NAME="tbxFornavn" ID=lnk0  type='text' MAXLENGTH=50 VALUE="<%=strFornavn%>"></td>
									</tr>
									<tr>
										<td>Etternavn:</td>
										<td><input class="mandatory" NAME="tbxEtternavn" ID=lnk1  type='text' MAXLENGTH=50 VALUE="<%=strEtternavn %>"></td>
									</tr>
									<tr>
										<td>Fødselsdato:</td>
										<td><input NAME="tbxFoedselsdato" ID=lnk2  type='text' MAXLENGTH=10 VALUE="<%=strFoedselsdato %>" ONBLUR="dateCheck(this.form, this.name)" >
									</tr>
									<tr>
										<td>Kjønn:</td>
										<td class="mandatory">
											<input NAME="rbnKjoenn" ID=lnk20  TYPE='RADIO' class='radio' VALUE="M" <%=strMann%>>Mann
											<input NAME="rbnKjoenn" TYPE='RADIO' class='radio' VALUE="K" <%=strKvinne%> ID="Radio1">Kvinne
										</td>
									</tr>
									<tr>
										<td>Nasjonalitet:</td>
										<td>
											<SELECT NAME="dbxCountry" ID=lnk12 >
											<%
											set rsType = GetFirehoseRS("SELECT CountryID,PrintableName FROM COUNTRY", Conn)

											Do Until rsType.EOF
												If rsType("CountryID") = lCountryID Then
													strValueSelected = rsType("CountryID") & " SELECTED"
												Else
													strValueSelected = rsType("CountryID")
												End If
												%>
												<OPTION VALUE=<%=strValueSelected %>><%=rsType("PrintableName") %>
												<%
												rsType.MoveNext
											Loop
											' Close AND release recordset
											rsType.Close
											Set rsType = Nothing
											%>
											</select>										
										</td>
									</tr>
								</table>
							</td>
							<td>
								<table ID="Table5">
									<tr>
										<td>Telefon:</td>
										<td><input NAME="tbxTelefon" ID=lnk3  type='text' MAXLENGTH=20 VALUE="<%=strTelefon %>"></td>
									</tr>
									</tr>
										<td>Mobil:</td>
										<td><input NAME="tbxMobilTlf" ID=lnk4  type='text' MAXLENGTH=20 VALUE="<%=strMobilTlf %>"></td>
									</tr>
									<tr>
										<td>Fax:</td>
										<td><input NAME="tbxFax" ID=lnk6  type='text' MAXLENGTH=20 VALUE="<%=strFax %>"></td>
									</tr>
									</tr>
										<td>E-Post:</td>
										<td><input NAME="tbxEPost" ID=lnk7  type='text' MAXLENGTH=50 onChange='statusChange()' VALUE="<%=strEPost %>"></td>
									</tr>
									<!--
									</tr>
										<td>Hjemmeside:</td>
										<td><input NAME="tbxHjemmeside" ID=lnk8  type='text' MAXLENGTH=50 VALUE="<%=strHjemmeside %>"></td>
									</tr>
									-->
									<tr>
										<td>Har f&oslash;rerkort:</td>
										<td>
											<%
												if strFoererkort = "0" then
													strSelected	 = ""
													strSelected2 = ""
													strSelected3 = "selected"
												elseif strFoererkort = "1" then
													strSelected	 = ""
													strSelected2 = "selected"
													strSelected3 = ""
												else
													strSelected	 = "selected"
													strSelected2 = ""
													strSelected3 = ""
												end if
											%>
											<SELECT NAME="cboForerkort" ID=lnk21 >
											<OPTION <%=strSelected%> VALUE=""></OPTION>
											<OPTION <%=strSelected2%> VALUE="1">Ja</OPTION>
											<OPTION <%=strSelected3%> VALUE="0">Nei</OPTION>
										</td>
									</tr>
									<tr>
										<td>Disponerer bil:</td>
										<td>
											<%
											if strBil = "0" then
												strSelected	 = ""
												strSelected2 = ""
												strSelected3 = "selected"
											elseif strBil = "1" then
												strSelected	 = ""
												strSelected2 = "selected"
												strSelected3 = ""
											else
												strSelected	 = "selected"
												strSelected2 = ""
												strSelected3 = ""
											end if
											%>
											<SELECT NAME="cboDispBil" ID=lnk22 >
											<OPTION <%=strSelected%> VALUE=""></OPTION>
											<OPTION <%=strSelected2%> VALUE="1">Ja</OPTION>
											<OPTION <%=strSelected3%> VALUE="0">Nei</OPTION>
										</td>
									</tr>
									<tr>
										<td>Oppsigelsestid:</td>
										<td>
											<%
											strSelected = ""
											strSelected2 = ""
											strSelected3 = ""
											strSelected4 = ""
											strSelected5 = ""
											strSelected6 = ""
											strSelected7 = ""
											select case strOppsigelsestid
											case ""
												strSelected = "selected"
											case "0"
												strSelected2 = "selected"
											case "1"
												strSelected3 = "selected"
											case "2"
												strSelected4 = "selected"
											case "3"
												strSelected5 = "selected"
											case "4"
												strSelected6 = "selected"
											case "5"
												strSelected7 = "selected"
											end select
											'strOppsigelsestid
											%>
											<SELECT NAME="cboOppsigelse" ID="Select1" >
											<OPTION VALUE=""  <%=strSelected%>></OPTION>
											<OPTION VALUE="0" <%=strSelected2%>>Ingen</OPTION>
											<OPTION VALUE="1" <%=strSelected3%>>14 dager</OPTION>
											<OPTION VALUE="2" <%=strSelected4%>>1 m&aring;ned</OPTION>
											<OPTION VALUE="3" <%=strSelected5%>>2 m&aring;neder</OPTION>
											<OPTION VALUE="4" <%=strSelected6%>>3 m&aring;neder</OPTION>
											<OPTION VALUE="5" <%=strSelected7%>>Over 3 m&aring;neder</OPTION>
										</td>
									</tr>
									<tr>
										<td>Stillingsbrøk:</td>
										<%
										strSelected = ""
										strSelected1 = ""
										strSelected2 = ""
										strSelected3 = ""
										strSelected4 = ""
										select case strWorkType
											case ""
												strSelected = "selected"
											case "0"
												strSelected1 = "selected"
											case "1"
												strSelected2 = "selected"
											case "2"
												strSelected3 = "selected"
											case "3"
												strSelected4 = "selected"											
										end select
										%>
										<td>
											<SELECT NAME="cboWorkType" ID="cboWorkType">
											<OPTION VALUE=""  <%=strSelected%>></OPTION>
											<OPTION VALUE="0" <%=strSelected1%>>Ikke angitt</OPTION>
											<OPTION VALUE="1" <%=strSelected2%>>Kun fulltid</OPTION>
											<OPTION VALUE="2" <%=strSelected3%>>Kun deltid</OPTION>
											<OPTION VALUE="3" <%=strSelected4%>>Fulltid og deltid</OPTION>
											</SELECT>
											
										</td>
										<%
										strSelected = ""
										strSelected1 = ""
										strSelected2 = ""
										strSelected3 = ""
										strSelected4 = ""
										%>
									</tr>
								</table>
							</td>
							<td>
								<table ID="Table6">
									<tr>
										<td>Hjemmeadresse:</td>
										<td><input class="mandatory" NAME="tbxAdresse" ID=lnk9  type='text' MAXLENGTH=50 VALUE="<%=strAdress %>"></td>
									</tr>
									<tr>
										<td>Postnr:</td>
										<td>
											<input NAME="tbxPostnr"  ID=lnk10  onBlur="ajax('GetPostOffice.asp' ,'lnk10',fillPostOffice)" type='text' MAXLENGTH=5 Value="<%=strPostnr%>">
										</td>
									</tr>
									<tr>
										<td>Poststed:</td>
										<td>
											<input NAME="tbxPoststed" ID=lnk11  type='text' MAXLENGTH=50 Value="<%=strPoststed%>">
										</td>
									</tr>
									<tr>
										<%
										if len(strUsername)>0 then
									 
											strChecked = "Checked"
											 strCBstate = "disabled" 

										else
											strChecked = ""
										end if
										%>
										<td class="nowrap">Opprett webbruker:</td>
										<%

										'if (isCVLocked) then
											'strChecked = ""
										'else
											'strChecked = "Checked"
										'end if
										%>
										<!--<td>Redigere web CV:<input id="chkEditCV" name='chkEditCV' type='checkbox' class='checkbox' <%=strChecked%> Value='1' onClick='ToggleEditCV()'></td>-->
										<td><input id='chkCreateuser'  name='chkCreateuser' type='checkbox' class='checkbox'   <%=strChecked%> Value='1' onClick='ckChange()' <%=strCBstate%>></td>
									</tr>
									<tr>
										<td>Brukernavn:</td>
										<td><input id='Username' name='Username' type='TEXT' value='<%=strUserName%>'></td>
									</tr>
									
									<tr>
										<td>Link 1:</td>
										<td><input id='Link1URL' name='Link1URL' type='TEXT' value='<%=strLink1URL%>' /></td>
									</tr>
									
									<tr>
										<td>Link 2:</td>
										<td><input id='Link2URL' name='Link2URL' type='TEXT' value='<%=strLink2URL%>' /></td>
									</tr>
									
									<tr>
										<td>Link 3:</td>
										<td><input id='Link3URL' name='Link3URL' type='TEXT' value='<%=strLink3URL%>' /></td>
									</tr>
									
								</table>
							</td>
						</tr>
					</table>
				</div>
				<div class="contentHead">
					<h2>Ansattinformasjon, avdelingskontor og tjenesteområder</h2>
				</div>
				<div class="content">
					<table  cellpadding="0" cellspacing="0" ID="Table7"  width="75%">
						<col width="35%">
						<col width="22%">
						<col width="22%">
						<col width="21%">
						<tr>
							<td>
								<table ID="Table8">
									<tr>
										<td>Intervjudato:</td>
										<td><input NAME="tbxIntervjudato" ID=lnk15  type='text' MAXLENGTH=10 VALUE="<%=strIntervjudato %>" ONBLUR="dateCheck(this.form, this.name)"></td>
									</tr>
									<tr>
										<td>Ansvarlig:</td>
										<%
										if strHeading = "Ny vikar" then
											if (not ISNULL(Session("medarbID")) ) AND (Not ISEMPTY(Session("medarbID")) ) then
												lAnsMedID = Session("medarbID")
											end if
										end if
										%>
										<td>
											<SELECT NAME="dbxMedarbeider" ID=lnk16 >
												
												<%
												response.write GetCoWorkersAsOptionList(lAnsMedID)
												%>
											</select>
										</td>
									</tr>
									<tr>
										<td>Ønsket timelønn:</td>
										<td><input NAME="tbxTimeloenn" ID=lnk17  type='text' MAXLENGTH=5 VALUE="<%=lTimeloenn %>"></td>
									</tr>
									<tr>
										<td>Kontonummer:</td>
										<td><input NAME="tbxBankkontonr" ID="tbxBankkontonr"  onBlur="javascript:ValidateBA(this);" type='text' MAXLENGTH=50 Value="<%=strAccountNr%>"></td>
									</tr>
									<tr>
										<td>Kontr.Sendt:</td>
										<td><input NAME="tbxKontraktsendt" ID=lnk18  type='text' MAXLENGTH=10 VALUE="<%=strkontraktsendt %>" ONBLUR="dateCheck(this.form, this.name)"></td>
									</tr>
									<tr>
										<td>Kontr.Mottatt:</td>
										<td><input NAME="tbxKontraktmottatt" ID=lnk19  type='text' MAXLENGTH=10 VALUE="<%=strKontraktmottatt %>" ONBLUR="dateCheck(this.form, this.name)"></td>
									</tr>
									<tr>
										<%
										strSelected = ""
										strSelected2 = ""
										strSelected3 = ""
										select case strOppsigelsestid
										case ""
											strSelected = "selected"
										case "0"
											strSelected2 = "selected"
										case "1"
											strSelected3 = "selected"
										case "2"
											strSelected4 = "selected"
										end select
										%>
										<td>Tilgjengelig:</td>
										<%
										strSelected1 = ""
										strSelected2 = ""
										strSelected3 = ""
										If (intTilgjengelig = 1) Then
											strSelected1 = "CHECKED"
										ElseIf (intTilgjengelig = 2) Then
											strSelected2 = "CHECKED"
										ElseIf (intTilgjengelig = 3) Then
											strSelected3 = "CHECKED"
										End If
										%>
										<td>
											<input NAME="rbnKurskode"  TYPE='RADIO' class='radio' VALUE="1" <%=strSelected1%> ID="Radio2">Dag
											<input NAME="rbnKurskode" TYPE='RADIO' class='radio' VALUE="2" <%=strSelected2%> ID="Radio3">Kveld
											<input NAME="rbnKurskode" TYPE='RADIO' class='radio' VALUE="3" <%=strSelected3%> ID="Radio4">Dag og Kveld
										</td>
									</tr>
									<tr>
										<td>Type:</td>
										<td>
											<SELECT NAME="dbxType" ID=lnk12 >
											<%
											set rsType = GetFirehoseRS("SELECT VikarTypeID, VikarType FROM H_VIKAR_TYPE", Conn)

											Do Until rsType.EOF
												If rsType("VikarTypeID") = lTypeID Then
													strValueSelected = rsType("VikarTypeID") & " SELECTED"
												Else
													strValueSelected = rsType("VikarTypeID")
												End If
												%>
												<OPTION VALUE=<%=strValueSelected %>><%=rsType("VikarType") %>
												<%
												rsType.MoveNext
											Loop
											' Close AND release recordset
											rsType.Close
											Set rsType = Nothing
											%>
											</select>
										</td>
									</tr>
									<tr>
										<td>Skattekort:</td>
										<td><select name="skattekort" id="skattekort">
												<option value="-1" <%If (strMottattSkattekort = "") Then Response.Write "selected"%>>Ikke levert</option>
												<%
												' Legger inn options for ifjor, i år og tre år fremover i tid.
												iAarTeller = 0
												iAar = Year(DateAdd("yyyy", -1, Date))
												While iAarTeller < 5
													If (strMottattSkattekort = iAar) Then
														Response.Write "<option value='" & iAar & "' selected>" & iAar & "</option>"
													Else
														Response.Write "<option value='" & iAar & "'>" & iAar & "</option>"
													End If
													iAar = Year(DateAdd("yyyy", iAarTeller, Date))
													iAarTeller = iAarTeller + 1
												Wend
												%>
											</select>
										</td>
									</tr>
									<tr>
										<td>Status:</td>
										<td>
											<select class="mandatory" name="dbxStatus"  ID="Select2">
												<option value=''> </option>
												<%
												set rsStatus = GetFirehoseRS("SELECT VikarStatusID, VikarStatus FROM H_VIKAR_STATUS ORDER BY VikarStatus", Conn)
												while (not rsStatus.EOF)
													If clng(rsStatus("VikarStatusID")) = clng(strStatus) Then
														strValueSelected = rsStatus("VikarStatusID") & " SELECTED"
													Else
														strValueSelected = rsStatus("VikarStatusID")
													End If
													Response.Write "<option value=" & strValueSelected & ">" & rsStatus("VikarStatus") & "</option>"
													rsStatus.MoveNext
												wend
												' Close AND release recordset
												rsStatus.Close
												Set rsStatus = Nothing
												%>
											</select>
										</td>
									</tr>
								</table>
							</td>
							<td>
								<table ID="Table9">
									<tr>
										<td valign='top'>Avdelingskontor:</td>
										<%
										'Avdelingskontor Fred 31.08.2000						
										strSQL = "SELECT distinct t2.id, t2.navn, t1.vikarid " & _
												" FROM vikar_arbeidssted t1, avdelingskontor t2 " & _
												" WHERE t2.show_hide = 1 and t1.AvdelingskontorID =* t2.id "
										
										if lngVikarID > 0 then
											strSQL = strSQL & "AND t1.vikarid =  " & lngVikarID
										else
											strSQL = strSQL & "AND t1.vikarid IS NULL "
										end if

										set rsVikarAvdeling = GetFirehoseRS(strSQL, Conn)
										%>
										</tr>
										<tr>
										<td>
											<select class='mandatory medium' id='LstAvdelingskontor' multiple name='lstAvdelingskontor' size='10'>
											<%
											' loop on result AND display in table
											while not rsVikarAvdeling.EOF
													If  rsVikarAvdeling("VikarID" ) > 0 Then
														strSelection = " SELECTED "
													Else
														strSelection = " "
													End If
													Response.Write "<OPTION " & strSelection & "VALUE=" & rsVikarAvdeling("ID" ) & "> " &  rsVikarAvdeling("navn" ) & "</OPTION>"
													rsVikarAvdeling.MoveNext
											wend
											' Close AND release recordset
											rsVikarAvdeling.Close
											Set rsVikarAvdeling = Nothing
											%>
											</select>
										</td>
									</tr>
								</table>
							</td>
							<td>
								<table ID="Table10">
									<tr>
										<td valign='top'>Tjenesteomr&aring;der:</td>
										<%
										strSQL = "SELECT DISTINCT t2.tomid, t2.navn, t1.vikarid " & _
													"FROM vikar_tjenesteomrade t1, tjenesteomrade t2 " & _
													"WHERE t1.tomID =* t2.tomid "

										if lngVikarID > 0 then
												strSQL = strSQL & "AND t1.vikarid =  " & lngVikarID & " "
										else
												strSQL = strSQL & "AND t1.vikarid IS NULL "
										end if

										strSQL = strSQL	& "order by t2.tomid"

										set rsVikartjenesteomrade = GetFirehoseRS(strSQL, Conn)
										%>
										</tr>
										<tr>
										<td>
										
											<select class='mandatory medium' id='Lsttjenesteomrader' multiple name='lsttjenesteomrader' size='10' onchange="PopulateCategories(this);">
											<%
											strTomIDs = ""
											strSelection = ""
											' loop on result AND display in table
											while (not rsVikartjenesteomrade.EOF)
												If  rsVikartjenesteomrade("VikarID" ) > 0 Then
													strSelection = " SELECTED "
													strTomIDs =  strTomIDs & rsVikartjenesteomrade("tomID") & ","
												Else
													strSelection = " "
												End If
												Response.Write "<OPTION " & strSelection & "VALUE=" & rsVikartjenesteomrade("tomID" ) & "> " &  rsVikartjenesteomrade("navn" ) & "</OPTION>"
												rsVikartjenesteomrade.MoveNext
											wend
											
											if(strTomIDs <> "") then
												strTomIDs = Mid(strTomIDs,1,Len(strTomIDs) - 1)
											end if
											
											' Close AND release recordset
											rsVikartjenesteomrade.Close
											Set rsVikartjenesteomrade = Nothing									
											%>
											</select>
											
										</td>
									</tr>
								</table>
							</td>
							<td>
								<table ID="TableCategory">
									<tr>
										<td valign='top'>Kategori:</td>
										<%
											strSQL = "Select DISTINCT t2.categoryid, t2.name,t1.vikarid, t2.tomid " & _
													"FROM vikar_category t1, oppdrag_category t2 " & _
													"WHERE t1.categoryid =* t2.categoryid "
													
											if lngVikarID > 0 then
												strSQL = strSQL & "AND t1.vikarid =  " & lngVikarID & " "
											else
													strSQL = strSQL & "AND t1.vikarid IS NULL "
											end if
											
											if strTomIDs <> "" then
												strSQL = strSQL & "AND t2.tomid in(" & strTomIDs & ") "
											else
												strSQL = strSQL & "AND t2.tomid IS NULL "
												
											end if	
											
											strSQL = strSQL	& "order by t2.tomid"
											
											set rsVikarCategory = GetFirehoseRS(strSQL, Conn)
										%>
									</tr>
									<tr>
										<td>
											<div id="divVikarCategory">
											<select class='mandatory' id='VikarCategory' multiple name='VikarCategory' size='10' style='width:220px;'>
											<%											
											strSelection = ""
											strSelectedCategories = ","
											' loop on result AND display in table

											while (not rsVikarCategory.EOF)												
												If  rsVikarCategory("VikarID" ) > 0 Then
													strSelection = " SELECTED "	
													strSelectedCategories = strSelectedCategories & rsVikarCategory("CategoryID" ) & ","											
												Else
													strSelection = " "
												End If

												Response.Write "<OPTION " & strSelection & "VALUE=" & rsVikarCategory("categoryID" ) & "> " &  rsVikarCategory("name" ) & "</OPTION>"

												rsVikarCategory.MoveNext
											wend																
											
											' Close AND release recordset
											rsVikarCategory.Close
											Set rsVikarCategory = Nothing									
											%>
											</select>											
											</div>
											<input NAME="tbxSelectedCategories" 	TYPE="HIDDEN" VALUE="<%=strSelectedCategories%>" ID="tbxSelectedCategories">
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
				</div>
				<div class="contentHead">
					<h2>Notat</h2>
				</div>
				<div class="content">
					<table class="layout" cellpadding="10" cellspacing="0" ID="Table11">
						<col width="50%">
						<col width="50%">
						<tr>
							<td>
								<h3><br>Notat</h3>
								<TEXTAREA ID="Textarea1"  NAME="tbxNotat"><%=strNotat%></TEXTAREA>
								<h3><br>Stillings-interesse</h3>
								<TEXTAREA ID="Textarea3"  NAME="tbxJobs"><%=strInterestedJobs%></TEXTAREA>
								<h3><br>Kandidatpresentasjon</h3>
								<TEXTAREA ID="Textarea2"  NAME="tbxPresentasjon"><%=strPresentasjon%></TEXTAREA>
							</td>
							<td>
								<h3>Oppsummering etter intervju<br>(OBS! Husk dine initialer!)</h3>
								<TEXTAREA ID=lnk23  NAME="tbxOppsumInt"><%=strOppsumIntervju%></TEXTAREA>
								<h3><br>Oppsummering etter referansesjekk</h3>
								<TEXTAREA ID=lnk24  NAME="tbxOppsumRef"><%=strOppsumRef%></TEXTAREA>
							</td>
						</tr>
					</table>
				</div>
			</form>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>