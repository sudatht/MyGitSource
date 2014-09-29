<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.FSO.inc"-->
<!--#INCLUDE FILE="includes\RStoJSON.inc"-->
<%
dim funcName
dim intSelcID
dim fieldName
dim sortOrder
dim vikarId
dim tomID
dim tomIDs
dim selectedCategories

' Move parameters to local variables
funcName = Request.Querystring("FuncName")
intSelcID = Request.Querystring("SelectedID")
vikarId = Request.Querystring("VikarID")
fieldName = Request.Querystring("FieldName")
sortOrder = Request.Querystring("SortOrder")	
tomID = Request.QueryString("TomID")	
selectedCategories =  Request.QueryString("SelectedCategories")	
tomIDs = Request.QueryString("TomIDs")

if funcName <> "" then
	Select case funcName
		Case "GetAllCoWorkersAsOptionList"		
			call GetAllCoWorkersAsOptionList(intSelcID)
		Case "GetActiveCoWorkersAsOptionList"
			call GetActiveCoWorkersAsOptionList(intSelcID)
		Case "GetVikarDoks"
			call GetVikarDoks(vikarId,fieldName,sortOrder)
		Case "GetOppdragCategories"
			call GetOppdragCategories(tomID)
		Case "GetVikarCategories"
			call GetVikarCategories(tomIDs,selectedCategories)
	End Select
end if

function GetAllCoWorkersAsOptionList(intSelectedID)
	dim strSel		'Contains "SELECTED" if current coworkers is the selected one
	dim rsMedarbeider	'Recordset containing coworkers
	dim strSQL		'SQL statement to execute, retriving coworkers
	dim strName		'Name of coworker
	dim fldMedID	'field ref 
	dim con			'Connection

	Set con = GetConnection(GetConnectionstring(XIS, ""))		
	if isnull(intSelectedID) then
		intSelectedID = 0
	end if	
	
	if (clng(intSelectedID) > 0) then	
		strSQL = "SELECT MED.Etternavn, MED.Fornavn, MED.MedID FROM  MEDARBEIDER MED INNER JOIN VIKAR VIK ON MED.VIKARID = VIK.VIKARID WHERE VIK.STATUSID = 4  UNION SELECT MED.Etternavn, MED.Fornavn, MED.MedID FROM  MEDARBEIDER AS MED WHERE (MedID = " & intSelectedID & ") ORDER BY MED.Etternavn"
	else	
		strSQL = "SELECT MED.Etternavn, MED.Fornavn, MED.MedID FROM  MEDARBEIDER MED INNER JOIN VIKAR VIK ON MED.VIKARID = VIK.VIKARID WHERE VIK.STATUSID = 4  ORDER BY MED.Etternavn"
	end if
	
	'Retrieve medarbeidere
	set rsMedarbeider = GetFirehoseRS(strSQL, con)
	
	set fldMedID = rsMedarbeider.fields("MedID")
	response.ContentType="text/html; Charset=ISO-8859-1" 
	Do Until rsMedarbeider.EOF 
		If (clng(intSelectedID) = clng(fldMedID.value))  Then
			strSel = " selected"
		Else
			strSel = ""
		End If 
		strName = cstr(rsMedarbeider.fields("Etternavn").value) & "  " & cstr(rsMedarbeider.fields("Fornavn").value)
		response.write "<option value=""" & fldMedID.value & """ " & strSel &" >" & strName & "</option>"
		rsMedarbeider.MoveNext
	Loop

	response.flush
	rsMedarbeider.Close
	Set rsMedarbeider = Nothing
	CloseConnection(con)
	set con = nothing			
end function

function GetActiveCoWorkersAsOptionList(intSelectedID)
	dim strSel		'Contains "SELECTED" if current coworkers is the selected one
	dim rsMedarbeider	'Recordset containing coworkers
	dim strSQL		'SQL statement to execute, retriving coworkers
	dim strName		'Name of coworker
	dim fldMedID	'field ref 
	dim con			'Connection

	Set con = GetConnection(GetConnectionstring(XIS, ""))		
	if isnull(intSelectedID) then
		intSelectedID = 0
	end if
	
	if (clng(intSelectedID) > 0) then
		strSQL = "SELECT MED.Etternavn, MED.Fornavn, MED.MedID FROM  MEDARBEIDER MED INNER JOIN VIKAR VIK ON MED.VIKARID = VIK.VIKARID WHERE MED.Active = 1 AND VIK.STATUSID = 4 UNION SELECT MED.Etternavn, MED.Fornavn, MED.MedID FROM  MEDARBEIDER AS MED WHERE (MedID = " & intSelectedID & ") ORDER BY MED.Etternavn"
	else
		strSQL = "SELECT MED.Etternavn, MED.Fornavn, MED.MedID FROM  MEDARBEIDER MED INNER JOIN VIKAR VIK ON MED.VIKARID = VIK.VIKARID WHERE MED.Active = 1 AND VIK.STATUSID = 4 ORDER BY MED.Etternavn"
	end if
	'Retrieve medarbeidere
	set rsMedarbeider = GetFirehoseRS(strSQL, con)

	set fldMedID = rsMedarbeider.fields("MedID")
	response.ContentType="text/html; Charset=ISO-8859-1"
	Do Until rsMedarbeider.EOF 
		If (clng(intSelectedID) = clng(fldMedID.value))  Then
			strSel = " selected"
		Else
			strSel = ""
		End If 
		strName = cstr(rsMedarbeider.fields("Etternavn").value) & "  " & cstr(rsMedarbeider.fields("Fornavn").value)
		response.write "<option value=""" & fldMedID.value & """ " & strSel &" >" & strName & "</option>"
		rsMedarbeider.MoveNext
	Loop

	response.flush
	rsMedarbeider.Close
	Set rsMedarbeider = Nothing
	CloseConnection(con)
	set con = nothing			
end function

function GetVikarDoks(vID,fName,sOrder)
	dim util
	dim filePath
	dim errStr
	dim rsFSO
	
	set  util = Server.CreateObject("XisSystem.Util")

	'Get location of all files
	filePath = Application("ConsultantFileRoot") & vID & "\"

	if (util.EnsurePathExists(filePath) = false) then		
		errStr = "Feil under aksessering av nettverksressurs (" & filePath & ")"
		AddErrorMessage(errStr)
		call RenderErrorMessage()
	end if
	call util.Logon()
	Set rsFSO = GetFSOFiles(filePath,fName,sOrder)
	call util.Logoff()
	set util = nothing
	response.ContentType="text/html; Charset=ISO-8859-1"
	response.write RStoJSON(rsFSO)
	response.flush
	'finally, close out the recordset
  	rsFSO.close()
  	Set rsFSO = Nothing
  	
end function

function GetOppdragCategories(tomID)
	dim strSQL
	dim con
	dim rsCategory
	dim strName
	dim catID
	
	Set con = GetConnection(GetConnectionstring(XIS, ""))
	
	strSQL = "SELECT OPPDRAG_CATEGORY.CategoryID, OPPDRAG_CATEGORY.Name FROM OPPDRAG_CATEGORY LEFT OUTER JOIN " & _
	"TJENESTEOMRADE ON OPPDRAG_CATEGORY.TomID = TJENESTEOMRADE.TomID " & _
	"WHERE     (TJENESTEOMRADE.TomID =" &  tomID & ")"
	set rsCategory = GetFirehoseRS(strSQL, con)
	
	response.ContentType="text/html; Charset=ISO-8859-1"
	Response.Write "<option VALUE='0' SELECTED >Choose -></option>"
	Do Until rsCategory.EOF 
			
		strName = cstr(rsCategory.fields("Name").value)
		catID = cstr(rsCategory.fields("CategoryID").value)
		response.write "<option value=""" & catID & """ >" & strName & "</option>"
		rsCategory.MoveNext
	Loop

	response.flush
	rsCategory.Close
	Set rsCategory = Nothing
	CloseConnection(con)
	set con = nothing	
	
end function

function GetVikarCategories(tomIDs,selectedCategories)
	dim strSQL
	dim con
	dim rsCategory
	dim strName
	dim catID
	dim strSelected
	
	Set con = GetConnection(GetConnectionstring(XIS, ""))
	
	strSQL = "SELECT OPPDRAG_CATEGORY.CategoryID, OPPDRAG_CATEGORY.Name, TJENESTEOMRADE.TomID FROM OPPDRAG_CATEGORY LEFT OUTER JOIN " & _
	"TJENESTEOMRADE ON OPPDRAG_CATEGORY.TomID = TJENESTEOMRADE.TomID " & _
	"WHERE     (TJENESTEOMRADE.TomID IN (" &  tomIDs & ")) ORDER BY TJENESTEOMRADE.TomID "
	set rsCategory = GetFirehoseRS(strSQL, con)
	
	response.ContentType="text/html; Charset=ISO-8859-1"
	Do Until rsCategory.EOF 
			
		strName = cstr(rsCategory.fields("Name").value)
		catID = cstr(rsCategory.fields("CategoryID").value)
		
		if(InStr(selectedCategories, "," & catID & ",") <> 0) then
			strSelected = " SELECTED "
		else 
			strSelected = ""
		end if
		
		response.write "<option " & strSelected & " value=""" & catID & """ >" & strName & "</option>"
		rsCategory.MoveNext
	Loop

	response.flush
	rsCategory.Close
	Set rsCategory = Nothing
	CloseConnection(con)
	set con = nothing	
	
end function
%>