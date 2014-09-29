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

	dim Conn
	dim sConnectionString
	dim strSQL
	dim rsConsultants
	dim oFileObject
	dim redirURL

	dim sUser
	dim iPos
	dim sStrippedDate
	dim strNavn
	dim strEPost
	dim strAdr1
	dim strAdr2
	dim strKjonn
	dim strTlf
	dim strMobil
	dim sFileName
	dim oOutStream
	dim oBkStream
	dim oBackupFile
	dim oTestFile
	dim RecStr
	dim nStartPos

	dim strResponse
	dim strConsultantsToUpdate

	Function OnlyDigits( strString )
		dim idx
		dim Digit
		dim strNewstring
		' Remove all non-nummeric signs from string
		If Not IsNull(strString) Then
		For idx = 1 To Len( strString) Step 1
			Digit = Asc( Mid( strString, idx, 1 ) )
			If (( Digit > 47 ) And ( Digit < 58 )) Then
				strNewstring = strNewString & Mid(strString, idx,1)
			End If
		Next
		End If
		Onlydigits = strNewString
	End Function

	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	if (request("IsPostback")="1") then
		'Check if transfer went ok..
		strResponse = lcase(request("btnTransferOK"))
		if strResponse = "ja" then
			'Comma separated list of consultants to export (from checkboxes)
			strConsultantsToUpdate = request("hdnConsultants")
			strSQL = "exec [dbo].[UpdateNotTransferedConsultants] '" & strConsultantsToUpdate & "'"
			if (ExecuteCRUDSQL(strSQL, Conn) = false) then
				CloseConnection(Conn)
				set Conn = nothing			
				AddErrorMessage("En feil oppstod under oppdatering av eksporterte!")
				call RenderErrorMessage()
			end if
			redirURL = "Vikar_timeliste_list3.asp"
		elseif strResponse = "nei" then
			redirURL = "LoennMeny.asp"
		end if
		CloseConnection(Conn)
		set Conn = nothing				
		response.Redirect(redirURL)
	end if
	'Retrieve consultants for export
	strSQL = "exec [dbo].[GetNotTransferedConsultants]"
	set rsConsultants = GetDynamicRS(strSQL, Conn)
%>
<html>
	<head>
		<title>Eksport av vikarer til H&amp;L</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Eksport av vikarer til H&amp;L</h1>
			</div>
			<div class="content">
				<form name="frmTransferConsultant" id="frmTransferConsultant" action="eksp_konsulent.asp" method="post">
					<input type="hidden" name="IsPostback" id="IsPostback" value="1">
					<%
					if HasRows(rsConsultants) then
						%>
						<table class="listing">
							<tr>
								<th>Ansattnummer</th>
								<th>Navn</th>
							</tr>
							<%
							Do Until (rsConsultants.EOF)
								%>
								<tr>
									<td><%=rsConsultants("ansattnummer").Value%><input type="hidden" name="hdnConsultants" value="<%=rsConsultants("VikarID").Value%>"></td>
									<td><%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & rsConsultants( "vikarid" ), rsConsultants("fornavn").Value & " " & rsConsultants("etternavn").Value, "Vis vikar " & rsConsultants("fornavn").Value & " " & rsConsultants("etternavn").Value )%></td>
								</tr>
								<%
								rsConsultants.MoveNext
							Loop
							%>
						</table>
						<%

						dim util
						set  util = Server.CreateObject("XisSystem.Util")
						
						call util.Logon()

						Set oFileObject = Server.CreateObject("Scripting.FileSystemObject")

						sUser			= Request.ServerVariables("LOGON_USER")
						iPos			= Instr(sUser,"\")
						sUser			= Mid(sUser, iPos + 1)
						sStrippedDate	= OnlyDigits(Now)
						sStrippedDate	= Left(sStrippedDate,10)
						sFileName		= "Kon_" & sStrippedDate & "_" & sUser & ".txt"

						'"rubicon_test" is the test area for HL / rubicon files, "rubicon" is the production area
						oBackupFile = Application("RubiconFileRoot") & "backupFiles\" & sFileName
						oTestFile = Application("RubiconFileRoot") & "Kon.txt"

						Response.write "HovedFil:<a target='new' href='" & oTestFile & "'>" & oTestFile & "</a><br>"
						Response.write "Sikkerhetskopi:<a target='new' href='" & oBackupFile & "'>" &  oBackupFile & "</a><br>"

						Set oOutStream	= oFileObject.CreateTextFile (oTestFile, True, False)
						Set oBkStream	= oFileObject.CreateTextFile (oBackupFile, True, False)

						set oFileObject = nothing

						rsConsultants.MoveFirst
						'---Write untransfered consultants to file

						'First line is header with tab separated field names:
						RecStr = "NR" & vbTab & "NAVN" & vbTab & "ADR1" & vbTab & "ADR2" & vbTab & "PNR" & vbTab & "KJONN" & vbTab & "BANK1" & vbTab & "EMAILLONN" & vbTab & "P_EMAILARB" & vbTab & "P_TLFPRIV" & vbTab & "P_TLFMOBIL" & vbTab
						oOutStream.WriteLine RecStr
						oBkStream.WriteLine RecStr


						while not (rsConsultants.EOF)
						'Write the consultant information to the file
							'The name field cannot be more than 30 characters long
							strNavn	= trim(rsConsultants("etternavn").Value) & " " & trim(rsConsultants("fornavn").Value)
							if len(strNavn)> 30 then
									strNavn	= trim(rsConsultants("etternavn").Value) & " " & trim(mid(rsConsultants("fornavn").Value,1,1)) & "."
									if len(strNavn)> 30 then
										nStartPos = InStrRev(rsConsultants("etternavn").Value," ")
										strNavn	= mid(trim(rsConsultants("etternavn").Value),1, nStartPos)
									end if
							end if

							'The adress fields cannot be more than 30 characters long
							strAdr1	=  trim(rsConsultants("adresse").Value)
							if len(strAdr1)> 30 then
									strAdr1	= mid(trim(rsConsultants("adresse").Value),1,30)
									strAdr2	= mid(trim(rsConsultants("adresse").Value),31)
							end if

							'Valid values for gender in H & L file is M/K/*
							if isnull(rsConsultants("kjoenn").Value) then
								strKjonn = "*"
							else
								if len(trim(rsConsultants("kjoenn").Value))=0 then
									strKjonn = "*"
								else
									strKjonn = rsConsultants("kjoenn").Value
								end if
							end if

							if isnull(rsConsultants("epost").Value) then
								strEPost = ""
							else
								if len(trim(rsConsultants("epost").Value))=0 then
									strEPost = ""
								elseif len(trim(rsConsultants("epost").Value))>49 then
									strEPost = ""
								else
									strEPost = trim(rsConsultants("epost").Value)
								end if
							end if

							strTlf	= trim(rsConsultants("telefon").Value)
							strMobil = trim(rsConsultants("MobilTlf").Value)

							RecStr = ""
							RecStr = trim(rsConsultants("Ansattnummer").Value) & vbTab & _
							strNavn & vbTab &_
							strAdr1  & vbTab & _
							strAdr2  & vbTab & _
							trim(rsConsultants("postnr").Value) & vbTab & _
							strKjonn & vbTab & _
							trim(rsConsultants("bankkontonr").Value) & vbTab & _
							strEPost & vbTab & _
							strEPost & vbTab & _
							strTlf & vbTab & _
							strMobil & vbTab

							oOutStream.WriteLine RecStr
							oBkStream.WriteLine RecStr

							rsConsultants.MoveNext
						wend

						Set oOutStream = Nothing
						Set oBkStream = Nothing
						
						call util.Logoff()		
						set util = nothing						
						%>
						<p>
							<h2>Var overføringen vellykket?</h2>
							<input name="btnTransferOK" TYPE="SUBMIT" VALUE="Ja">
							<INPUT name="btnTransferOK" TYPE="SUBMIT" VALUE="Nei">
						</p>
						<%
					else
						%>
						Ingen vikarer til overføring!
						<%
					end if
				%>
				</form>
			</div>
		</div>
	</body>
</html>
<%
	CloseConnection(Conn)
	set Conn = nothing
	set rsConsultants = nothing
%>