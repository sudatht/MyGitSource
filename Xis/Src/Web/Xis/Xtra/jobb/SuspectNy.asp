<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	'File purpose:		Transfers the specified suspect to the consultant table (vikar).

	dim LsuspectID			'as long
	dim ioppsigelse			'as integer
	dim iworkType			'as integer
	dim strHasLicense
	dim strHasCarAccess
	dim lngAnsMedID
	dim strVikarID			'as string
	dim strSuspectID		'as string
	dim strEtternavn		'as string
	dim strFornavn			'as string
	dim strTelefon			'as string
	dim strMobilTlf			'as string
	dim strEPost			'as string
	dim strFax				'as string
	dim strHjemmeside		'as string
	dim strFoedselsdato		'as string
	dim strAnsMed			'as string
	dim strKjonn  			'as string
	dim strNotat			'as string
	dim strInterestedJobs   'as string
	dim strAdress			'as string
	dim strPostnr			'as string
	dim strPostSted			'as string
	dim strSQL				'as string
	dim strLink1URL			'as string
	dim strLink2URL			'as string
	dim strLink3URL			'as string
	dim hasCar			'as string
	dim hasCarValue			'as string

	dim ObjSuspect			'as xtraweb.suspect
	dim ObjApplication		'as xtraweb.JobApplication
	dim ColJobPlaces		'as XtraWeb.JobPlaces
	dim ObjJobPlace			'as XtraWeb.JobPlace
	dim ObjCV				'as XtraWeb.CV
	dim ObjConnection		'as ADODB.connection
	dim rsAdress			'as ADODB.recordset
	dim rsVikarAvdeling		'as ADODB.recordset
	dim rsTjenesteomrade	'as ADODB.recordset
	dim sDAFilePath
	dim sCVFilePath
	dim objFSO
	dim StrCvfile

	const c_soeker_status = "1"

	'Initialize objects / variables
	set ObjSuspect						= Server.CreateObject("XtraWeb.Suspect")
	ObjSuspect.XtraConString			= Application("Xtra_intern_ConnectionString")
	ObjSuspect.XtraDataShapeConString	= Application("ConXtraShape")

	LsuspectID = Request.Querystring("suspectID")

	if (not ISNULL(Session("medarbID")) ) AND (Not ISEMPTY(Session("medarbID")) ) then
		lngAnsMedID = Session("medarbID")
	end if

	'Retrieve specified suspect
	if not ObjSuspect.GetSuspect(LsuspectID) then
		ObjSuspect.CV.cleanup
		ObjSuspect.cleanup
		set ObjSuspect = nothing
		AddErrorMessage("feil: Suspect ikke funnet!")
		call RenderErrorMessage()
	end if

	'-- Transfer values to variables
	'First mandatory fields..
	strFornavn			= ObjSuspect.DataValues("Fornavn").value
	strEtternavn		= ObjSuspect.DataValues("Etternavn").value
	
	'Optional fields..
	if (isnull(ObjSuspect.DataValues("Foedselsdato").value)) then
		strFoedselsdato		= "NULL"
	else
		strFoedselsdato		= DbDate(ObjSuspect.DataValues("Foedselsdato").value)
	end if
	
	strTelefon		= ObjSuspect.DataValues("Telefon").value
	strMobilTlf		= ObjSuspect.DataValues("MobilTlf").value
	strFax			= ObjSuspect.DataValues("Fax").value
	strEPost		= ObjSuspect.DataValues("EPost").value
	strHjemmeside	= ObjSuspect.DataValues("Hjemmeside").value
	strKjonn		= ucase(ObjSuspect.DataValues("Kjoenn").value)
	strLink1URL   	= ObjSuspect.DataValues("Link1URL").value
	strLink2URL   	= ObjSuspect.DataValues("Link2URL").value
	strLink3URL   	= ObjSuspect.DataValues("Link3URL").value

	if (ObjSuspect.DataValues("foererkort").value="1") then
		strHasLicense = "1"
	else
		strHasLicense = "0"
	end if

	if (cbool(ObjSuspect.DataValues("bil"))) then
		strHasCarAccess = 1
	else
		strHasCarAccess = 0
	end if
	
	hasCar = cint(ObjSuspect.DataValues("bil"))
	if(hasCar = 0) then
	  hasCarValue = 0	
	else
	  hasCarValue = 1
	end if
	
	ioppsigelse	= ObjSuspect.DataValues("oppsigelsestid").value
	iworkType	= ObjSuspect.DataValues("WorkType").value

	set ObjApplication				= ObjSuspect.JobApplication
	ObjApplication.Refresh()
	set ColJobPlaces				= ObjApplication.JobPlaces
	set ObjCV						= ObjSuspect.CV
	ObjCV.XtraConString				= Application("Xtra_intern_ConnectionString")
	ObjCV.XtraDataShapeConString	= Application("ConXtraShape")
	ObjCV.refresh
    
    'SKA
    set uplFiles = ObjCV.UplaodFiles

	if ObjApplication.DataValues.Count > 1 then
		strNotat = ObjApplication.DataValues("Description").value
		strInterestedJobs = ObjApplication.DataValues("InterestedJobs").value
	end if
		

	' Open database connection
	
	Set objCon = GetConnection(GetConnectionstring(XIS, ""))	
	
	strSQL = "select A.Adresse, A.Postnr, A.PostSted from V_SUSPECT_ADRESSE A " & _
				" WHERE A.adresseRelID =" & LsuspectID & " and A.AdresseType = 2"

	set rsAdress = GetFirehoseRS(strSQL, objCon)	

	if (HasRows(rsAdress)) then
		strAdress	= rsAdress.fields("Adresse").value
		strPostnr	= rsAdress.fields("Postnr").value
		strPostSted = rsAdress.fields("PostSted").value
	end if
	rsAdress.close
	set rsAdress = nothing

	strSQL = "INSERT INTO Vikar(" & _
		"Fornavn, Etternavn, Foedselsdato," & _
		"StatusID, AnsMedID, Notat,InterestedJobs," & _
		"Telefon, MobilTlf, fax, Epost, Hjemmeside, "&_
		"kjoenn, Foererkort, Bil, " & _
		"Oppsigelsestid,WorkType, Link1URL,Link2URL,Link3URL,hasCar,Regdato " & _
		")" &_
		"VALUES('" & _
		strFornavn & "','" &_
		strEtternavn & "'," & _
		strFoedselsdato & "," & _
		c_soeker_status & "," & _
		lngAnsMedID & "," & _
		FixString(strNotat) & "," & _
		FixString(strInterestedJobs) & "," & _
		FixString(strTelefon) & "," & _
		FixString(strMobilTlf) & "," & _
		FixString(strFax) & "," & _
		FixString(strEPost) & "," & _
		FixString(strHjemmeside) & "," & _
		FixString(strKjonn) & "," & _
		FixString(strHasLicense) & "," & _
		cstr(cint(ObjSuspect.DataValues("bil"))) & "," & _
		FixString(ioppsigelse) & "," & _
		FixString(iworkType) & "," & _
		FixString(strLink1URL) & "," & _
		FixString(strLink2URL) & "," & _
		FixString(strLink3URL) & "," & _
		hasCarValue & "," & _
		dbDate(date) & " )"

	' Start transaction
	objCon.Begintrans

	If ExecuteCRUDSQL(strSQL, objCon) = false then
		CloseConnection(objCon)
		set objCon = nothing
		AddErrorMessage("Feil oppstod under overføring suspect.")
		call RenderErrorMessage()
	End if	

	' Get new VikarID
	set rsVikar = GetFirehoseRS("Select NewVikarID=max( VikarID ) from Vikar", objCon)

	NewVikarID = rsVikar("NewVikarID")

	' Close and release recordset
	rsVikar.Close
	Set rsVikar = Nothing
	
	'Log activity for suspect transfer as vikar
	Dim strActivity
	Dim rsActivityType
	Dim strComment
	Dim rsVikarStatus
	strActivity = "Søker reg. som vikar"
	'Get vikar status
	set rsVikarStatus = GetFirehoseRS("SELECT Vikarstatus FROM H_VIKAR_STATUS WHERE VikarstatusID = " & c_soeker_status, objCon)
	strComment = "Registrert som vikar med status " & rsVikarStatus("Vikarstatus")
	' Close and release recordset
      	rsVikarStatus.Close
      	Set rsVikarStatus = Nothing
	 
	'Get activity id
	set rsActivityType = GetFirehoseRS("SELECT AktivitetTypeID FROM H_AKTIVITET_TYPE WHERE AktivitetType = '" & strActivity & "'", objCon)
	nActivityTypeID = CInt(rsActivityType("AktivitetTypeID"))
	' Close and release recordset
      	rsActivityType.Close
      	Set rsActivityType = Nothing
	sDate = GetDateNowString()
	strSql = "INSERT INTO AKTIVITET( AktivitetTypeID, AktivitetDato, VikarID, Notat, Registrertav, RegistrertAvID, AutoRegistered ) Values(" & nActivityTypeID & ",'" & sDate & "'," & NewVikarID & ",'" & strComment & "','" & Session("Brukernavn") & "'," & lngAnsMedID & ", 1)"
	
	If ExecuteCRUDSQL(strSql, objCon) = false then
		'CloseConnection(objCon)
		set objCon = nothing
		AddErrorMessage("Aktivitetsregistrering for ny vikar i xis feilet.")
		call RenderErrorMessage()
	End if
	
	'Same time log the activity of suspect register on web
	strActivity = "Søker fra web reg."
	'Get activity id
	set rsActivityType = GetFirehoseRS("SELECT AktivitetTypeID FROM H_AKTIVITET_TYPE WHERE AktivitetType = '" & strActivity & "'", objCon)
	nActivityTypeID = CInt(rsActivityType("AktivitetTypeID"))
	' Close and release recordset
      	rsActivityType.Close
      	Set rsActivityType = Nothing
	sDate = DbDate(ObjSuspect.DataValues("RegDato").value)
	Dim dt
	dt = cdate(ObjSuspect.DataValues("RegDato").value)
	
	'dt.Hour = 0
	sDate = Year(dt) & "-" & Month(dt) & "-" & Day(dt) & " 00:00:00"
	strComment = strFornavn & " " & strEtternavn & " registrert via web"
	strSql = "INSERT INTO AKTIVITET( AktivitetTypeID, AktivitetDato, VikarID, Notat, Registrertav, RegistrertAvID, AutoRegistered ) Values(" & nActivityTypeID & ",'" & sDate & "'," & NewVikarID & ",'" & strComment & "','" & Session("Brukernavn") & "'," & lngAnsMedID & ", 1)"
	'response.write(strSql)
	If ExecuteCRUDSQL(strSql, objCon) = false then
		'CloseConnection(objCon)
		set objCon = nothing
		AddErrorMessage("Aktivitet ikke lagret for websøker.")
		call RenderErrorMessage()
	End if

	' Create new main adress
	' AdresseRelasjon: 1 = Kunde / 2 => Vikar
	strSQL = "Insert into ADRESSE(AdresseRelasjon, AdresseRelID, Adresse, Postnr, Poststed, AdresseType ) " & _
	"Values( 2," & NewVikarID & "," & _
	FixString(strAdress) & "," & _
	FixString(strPostnr) & "," & _
	FixString(strPostSted) & "," & _
	"1)"

	If ExecuteCRUDSQL(strSQL, objCon) = false then
		objCon.RollBackTrans
		CloseConnection(objCon)
		set objCon = nothing
		AddErrorMessage("Feil oppstod under overføring addresse.")
		call RenderErrorMessage()
	End if	

	for each ObjJobPlace in ColJobPlaces
		strSQL = "insert into VIKAR_ARBEIDSSTED( VikarID, AvdelingskontorID ) " & _
			"SELECT " & NewVikarID & ", [AVDELINGSKONTOR].[ID] " & _
			"FROM [AVDELINGSKONTOR] " & _
			"INNER JOIN [LOKASJON] ON [AVDELINGSKONTOR].[LokasjonID] = [LOKASJON].[LokasjonID] " & _
			"WHERE [LOKASJON].[Navn] LIKE '" & trim(ObjJobPlace.DataValues("PlaceName")) & "' "

		If ExecuteCRUDSQL(strSQL, objCon) = false then
			objCon.RollBackTrans
			CloseConnection(objCon)
			set objCon = nothing
			AddErrorMessage("Feil oppstod under overføring til avdelingskontor.")
			call RenderErrorMessage()
		End if
	next

	objCon.CommitTrans

	'Create document archive folder
	sDAFilePath = Application("ConsultantFileRoot") & NewVikarID & "\"
	sCVFilePath = Application("CVFileRoot")
	
	'Response.Write("file path of sCVFilePath :" & sCVFilePath )
	'Response.Write("file path of sDAFilePath :" & sDAFilePath )

	dim util
	set  util = Server.CreateObject("XisSystem.Util")

	call util.Logon()
		
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	if (not objFSO.FolderExists(sDAFilePath))then
		objFSO.createFolder(sDAFilePath)
	end if

	'Transfer CV file if it exists
	'Response.Write("file path of sCVFilePath :" & sCVFilePath )
	'Response.Write("file path of sDAFilePath :" & sDAFilePath )

	if not isNull(ObjCv.DataValues("Filename")) then
		StrCvfile = ObjCv.DataValues("Filename")
		'Response.Write(ObjCv.DataValues("Filename") )
		if(objFSO.FileExists(sCVFilePath & "\" & StrCvfile)) then
			objFSO.CopyFile sCVFilePath & "\" & StrCvfile, sDAFilePath, false
		end if
	end if
	
	'sCVFilePath = sCVFilePath & "\" & LsuspectID
	sCVFilePath = sCVFilePath  & LsuspectID	
	
	if (uplFiles.Count > 0) then
	for i=1 to uplFiles.Count
	    strFileName = uplFiles.Item(i).DataValues("FileName")
	    'Response.write(sCVFilePath)
	    if(objFSO.FileExists(sCVFilePath & "\" & strFileName)) then
			'Response.write(sCVFilePath)
			'Response.write(strFileName)
			'Response.write(sDAFilePath)  
			objFSO.CopyFile sCVFilePath & "\" & strFileName, sDAFilePath, false
	    end if	
	
	next
	
	end if
	
	'MDE CV transfer 
		Set objCon2 = GetConnection(GetConnectionstring(XIS, ""))
		strSQL = "SELECT  [CV].[Filename] FROM [CV]  WHERE  [Filename] IS NOT NULL AND [TYPE] = 'S' AND [ConsultantID]=" & LsuspectID
		set rsCV =  GetFirehoseRS(strSQL,objCon2)
			if (HasRows(rsCV)) then
			  StrCvfile = rsCV.fields("Filename")
			  set objCon2 = nothing
		      if(objFSO.FileExists(sCVFilePath & "\" & StrCvfile)) then
		  	  objFSO.CopyFile sCVFilePath & "\" & StrCvfile, sDAFilePath, false
		     end if
	end if
 
		
	set rsCV = Nothing
	
	
	call util.Logoff()		
	set util = nothing
	%>
	
	<% 'This file transfers all the cv data (edu, courses, prof, ref) from suspect to vikar %>   
	
	<!--#INCLUDE FILE="Suspect_overfor_cv.asp"-->


	<%
	

			'DNN user updateion goes here
	'response.Write("suspet is "& LsuspectID)
	'response.Write("vikar id is"& NewVikarID)
	
	Call UserMapUpdate(LsuspectID,NewVikarID )
	
	'Update suspect status
	strSQL = "UPDATE v_suspect SET overfort = 1 WHERE suspectID = "& LsuspectID
	
	call ExecuteCRUDSQL(strSQL, objCon)

	CloseConnection(objCon)
	set objCon = nothing

	set ObjCV	= nothing
	set ObjJobPlace	= nothing
	set ColJobPlaces = nothing
	ObjApplication.Cleanup()
	set ObjApplication	= nothing
	ObjSuspect.CV.Cleanup()
	ObjSuspect.cleanup
	set ObjSuspect	= nothing

	Response.redirect "../VikarNy.asp?vikarID=" & NewVikarID
	
	
sub UserMapUpdate(suspectid,vikarid)

	
	dim iSuspectID 	'string
	dim IVikarID 		'Consultant's id
	Dim objUserProxy  'Web service proxy for the DNN user service
	Dim sUserServiceURL 'Url of the user web service
	Dim iApp
	
	iSuspectID  = Cstr(suspectid)
     IVikarID   = Cstr(vikarid)
	iApp = Cstr(Application("Application"))
	sUserServiceURL = Application("DNNUserServiceURL")
	Set objUserProxy = Server.CreateObject("ECDNNUserProxy.UserServiceProxy")
	objUserProxy.Url = sUserServiceURL

		If Not objUserProxy.ConvertSuspectUserToVikar(IVikarID,iSuspectID ,iApp) Then
            AddErrorMessage("Failed to save web user") // 1
			Call RenderErrorMessage()
	   End If
		
end sub
%>
