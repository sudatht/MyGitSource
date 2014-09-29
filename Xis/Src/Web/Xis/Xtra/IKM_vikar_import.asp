<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<%
Server.ScriptTimeout = 2000
%>
<!--#INCLUDE FILE="includes\Library.inc"-->
<!--#INCLUDE FILE="includes\MailLib.inc"-->
<%

	dim sql
	dim fornavn
	dim etternavn
	dim gateAdresse
	dim postnummer
	dim sted
	dim tjenesteomrade
	dim personnr
	dim kjonn
	dim mobile
	dim privatePhone
	dim phoneWork
	dim email
	dim NewVikarID
	dim tomid
	dim fagid
	dim cvid
	dim nummer
	dim dob
	dim docName
	dim docnameTmp
	dim cons
	dim cv
	dim DAFilePath

	dim rsConsultants
	dim rsConsultantID
	dim rsDuplicates
	dim Command

	dim fso

	CONST STATUSID = "1"
	CONST TYPEID = "1"
	CONST TIMELONN = "0"
	CONST MEDID = "270"
	CONST AVDELINGSKONTORID = "2"
	CONST KURSKODE = "0"
	CONST DOCPATH = "d:\w3server\_dummy web folders\Div"
	CONST CVPATH = "d:\w3server\_dummy web folders\Vikardok"

	' Check parameters..
	' ------------------
	'etternavn, Mandatory
	'Fornavn, Mandatory
	'Adresse, Mandatory
	'	- Postnr må være nummerisk
	'Tjenesteområder, Mandatory
	'Foedselsdato, Mandatory
	'Avdelingskontor, Mandatory
	'Personnummer, Mandatory
	'Timelønn = 0
	'Kjønn, Mandatory
	'Status = 1 (Søker)

	' Open database connection
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.ConnectionTimeOut = Session("Xtra_ConnectionTimeout")
	Conn.CommandTimeOut = Session("Xtra_CommandTimeout")
	Conn.Open Session("Xtra_ConnectionString"), Session("Xtra_RuntimeUserName"), Session("Xtra_RuntimePassword")

	set Command = server.CreateObject("adodb.command")
	set Command.ActiveConnection =  Conn

	'Retrieve import data
	sql = "SELECT " & _
	"[Address], [PostalCode], " & _
	" CASE [Country] " & _
	      " WHEN 'NORWAY' THEN [City] " & _
	      " WHEN 'NO' THEN [City] " & _
	      " ELSE [City] + ',' + [Country] " & _
	" END AS Place, " & _
	"[FirstName], [LastName], [TelephoneMobile], [PersonNumber],  " & _
	"[TelephonePrivate], [TelephoneWork], [Email],  " & _
	"[Date_of_birth], [Fagid_x-is] as fagId, [Tom_id] " & _
	"FROM  " & _
	"[importFromIKM]  " & _
	"ORDER BY " & _
	"[LastName], [FirstName] "

	'Response.Write sql & "<br>"

	set rsConsultants = Conn.Execute(sql)

	nummer	= 1

	Set fso = Server.CreateObject("Scripting.FileSystemObject")

	'Run through all consultants..
	while (not rsConsultants.EOF)
		'Make sure consultant is not a duplicate
		personnummer = rsConsultants.fields("PersonNumber").value
		fornavn = rsConsultants.fields("Firstname").value
		etternavn = rsConsultants.fields("Lastname").value
		gateAdresse = rsConsultants.fields("address").value
		postnummer = rsConsultants.fields("postalcode").value
		sted = rsConsultants.fields("place").value
		tjenesteomrade = rsConsultants.fields("tom_id").value
		mobile = rsConsultants.fields("TelephoneMobile").value
		privatePhone = rsConsultants.fields("TelephonePrivate").value
		phoneWork = rsConsultants.fields("TelephoneWork").value
		email = rsConsultants.fields("Email").value
		tomid = rsConsultants.fields("Tom_id").value
		fagid = rsConsultants.fields("Fagid").value

		dob = rsConsultants.fields("Date_of_birth").value
		dob = mid(dob, 9, 2) & "." & mid(dob, 6, 2) & "." & mid(dob, 1, 4)

		'Retrieve possible duplicates
		Command.CommandText = "exec sp_intra_getPossibleDuplicates '"  & fornavn & "', '" & etternavn & "', " & dbdate(dob)
		Command.CommandType = 1
		'Response.Write Command.CommandText & "<br><br>"
		set rsDuplicates = Command.Execute

		if (rsDuplicates.EOF) then
			'Save consultant
			sql = "INSERT INTO Vikar( [Fornavn], [Etternavn], [Foedselsdato]," & _
					"[StatusID], [TypeID], [AnsMedID], [loenn1]," & _
					"[Telefon], [MobilTlf], [Epost], "&_
					"[kurskode]" & _
					" )" &_
					"Values('" & fixLongText(fornavn, 50, false) & "','" &_
					fixLongText(etternavn, 50, false) & "'," & _
					dbdate(dob) & "," & _
					STATUSID & "," & _
					TYPEID & "," &_
					MEDID & "," & _
					TIMELONN & "," & _
					"'" & fixLongText(privatePhone, 20, true) & "'," & _
					"'" & fixLongText(mobile, 20, true) & "'," & _
					"'" & fixLongText(email, 255, false) & "'," & _
					KURSKODE & ")"

			'Response.Write sql & "<br><br>"

			Conn.Execute( sql )
			If Conn.Errors.Count > 0 then
				Call SqlError()
			End if

			' Get VikarID
			Set rsConsultantID  = Conn.Execute("Select NewVikarID = max( VikarID ) from Vikar")
			NewVikarID = rsConsultantID("NewVikarID")
			' Close and release recordset
			rsConsultantID.Close
			Set rsConsultantID = Nothing

			'Save address on consultant
			' AdresseRelasjon: 1 = Kunde / 2 => Vikar
			sql = "INSERT INTO ADRESSE([AdresseRelasjon], [AdresseRelID], [Adresse], [Postnr], [Poststed], [AdresseType] ) " & _
					"Values( 2," & NewVikarID & "," & _
					"'" & gateAdresse & "','" & _
					postnummer & "','" & _
					sted & "', 1)"

			'Response.Write sql & "<br><br>"

			Conn.Execute( sql )
			If Conn.Errors.Count > 0 Then
				Call SqlError()
			End if

			'Save avdelingskontor
			sql = "INSERT INTO VIKAR_ARBEIDSSTED( VikarID, AvdelingskontorID ) Values(" & NewVikarID & "," & AVDELINGSKONTORID & ")"

			'Response.Write sql & "<br><br>"

			Conn.Execute( sql )
			If Conn.Errors.Count > 0 then
				Call SqlError()
			End if

			'Save tjenesteområde
			sql = "INSERT INTO VIKAR_TJENESTEOMRADE( VikarID, tomID ) Values(" & NewVikarID & "," & tomid & ")"

			'Response.Write sql & "<br><br>"

			Conn.Execute( sql )
			If Conn.Errors.Count > 0 then
				Call SqlError()
			End if

			'Create CV
			set cons = Server.CreateObject("XtraWeb.Consultant")
			cons.XtraConString = Application("XtraWebConnection")
			cons.GetConsultant(NewVikarID)

			'Determine CV information
			set cv	= cons.CV
			cv.XtraConString = Application("Xtra_intern_ConnectionString")
			cv.XtraDataShapeConString = Application("ConXtraShape")
			cv.Refresh

			if cons.CV.DataValues.Count = 0 then
				cons.CV.Save
			end if

			cvid = cons.CV.DataValues("cvid")

			'Create document archive folder
			DAFilePath = CVPATH & NewVikarID & "\"

			if (not fso.FolderExists(DAFilePath))then
				fso.createFolder(DAFilePath)
			end if

			'Add Cv-file
			docnameTmp = fornavn & " " & etternavn & " " & personnummer & ".doc"
			docnameTmp = lcase(docnameTmp)
			docnameTmp = replace(docnameTmp,"å", "a")
			docnameTmp = replace(docnameTmp,"ø", "o")
			docnameTmp = replace(docnameTmp,"æ", "ae")
			docnameTmp = replace(docnameTmp,"-", "")
			docName = replace(docnameTmp," ", "_")

			If (fso.FileExists(DOCPATH & docName)) then
				Response.Write "Docpath 1:" & DOCPATH & docName & "<br>"
				'Move doc to consultant doc area
				fso.MoveFile DOCPATH & docName, DAFilePath & "CV_" & NewVikarID & ".doc"
			else
				Response.Write "Document not found:" & DOCPATH & docName & "<br>"
			end if

			if (cint(fagid) > 0) then
				sql = "INSERT INTO [Vikar_Kompetanse]( [VikarID], [CVid], [K_TypeID], [K_TittelID]) Values(" &_
				NewVikarID & "," & _
				cvid & "," & _
				"4," & _
				fagid & ")"

				Conn.Execute( sql )
				If Conn.Errors.Count > 0 then
					Call SqlError()
				End if

			end if

			set cv = nothing
			cons.CV.cleanup
			cons.cleanup
			set cons = nothing

			'Add consultant to report with status
			Response.Write "Vikar nr." & nummer & " - " & fornavn & " " & etternavn & "(" & NewVikarID & ") ble lagt til.<br>"
		else
			'There are possible duplicates
			Response.Write "Mulig duplikat for Vikar nr." & nummer & " " & fornavn & " " & etternavn & ".<br>"
			while (not rsDuplicates.EOF)
				Response.Write "Duplikat <a href='vikarvis.asp?vikarid=" & rsDuplicates.fields("vikarid").value & "' target='new'>" & rsDuplicates.fields("fornavn").value & " " & rsDuplicates.fields("etternavn").value & " " & rsDuplicates.fields("foedselsDato").value & "</a><br>"
				rsDuplicates.movenext
			wend
		end if
		rsDuplicates.close
		set rsDuplicates = nothing
		Response.Write "<br>"

		rsConsultants.movenext
		nummer = nummer + 1
	wend
	set fso = nothing
	set Command = nothing
	' Close and release recordset
	rsConsultants.Close
	Set rsConsultants = Nothing

	function fixLongText(text, maxlength, isNummeric)
		if (len(text) <= clng(maxlength)) then
			fixLongText = text
			exit function
		end if

		'if (isNummeric = true) then
		'	dim regEx
		'	Set regEx = New RegExp
		'	regEx.Pattern = "^\D"
		'	regEx.IgnoreCase = True
		'	fixLongText = regEx.Replace(text, "|")
		'	fixLongText = replace(fixLongText, "|", "")
		'else
		fixLongText = Mid(text, 1, maxlength)
		'end if
	end function
 %>
