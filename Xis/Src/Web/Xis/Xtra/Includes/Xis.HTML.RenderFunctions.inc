<%
function GetCoWorkersAsOptionList(intSelectedID)
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
	rsMedarbeider.Close
	Set rsMedarbeider = Nothing
	CloseConnection(con)
	set con = nothing			
end function

function GetActiveAgreementList(intSelectedID)
	dim strSel		'Contains "SELECTED" if current FAgreement is the selected one
	dim rsFAgreements	'Recordset containing FAgreement
	dim strSQL		'SQL statement to execute, retriving FAgreement
	dim strName		'Name of FAgreement
	dim fldFaID	'field ref 
	dim con			'Connection

	Set con = GetConnection(GetConnectionstring(XIS, ""))		
	if isnull(intSelectedID) then
		intSelectedID = 0
	end if
	
	if (clng(intSelectedID) > 0) then
		strSQL = "select * from FrameworkAgreement  where ValidToDate > GETDATE() Order By FaCode"
	else
		strSQL = "select * from FrameworkAgreement  where ValidToDate > GETDATE() Order By FaCode"
	end if
	'Retrieve medarbeidere
	set rsFAgreements = GetFirehoseRS(strSQL, con)

	set fldFaID = rsFAgreements.fields("FaID")
	Do Until rsFAgreements.EOF 
		If (clng(intSelectedID) = clng(fldFaID.value))  Then
			strSel = " selected"
		Else
			strSel = ""
		End If 
		strName = cstr(rsFAgreements.fields("FaCode").value) & " - " 
		If Not isNull(rsFAgreements.fields("FaName").value) Then
			strName = strName & CStr(rsFAgreements.fields("FaName").value)
		End If 
		response.write "<option value=""" & fldFaID.value & """ " & strSel &" >" & strName & "</option>"
		rsFAgreements.MoveNext
	Loop
	rsFAgreements.Close
	Set rsFAgreements = Nothing
	CloseConnection(con)
	set con = nothing			
end function

function GetActiveCategoryList(intFAgreementId)
	dim strSel		
	dim rsFAgreements	'Recordset containing Categories
	dim strSQL		'SQL statement to execute, retriving Categories
	dim strName		'Name of Category
	dim fldFaCID	'field ref 
	dim con			'Connection

	Set con = GetConnection(GetConnectionstring(XIS, ""))		
	if isnull(intFAgreementId) then
		intFAgreementId = 0
	end if
	
	if (clng(intFAgreementId) > 0) then
		strSQL = "SELECT * FROM FrameworkCategory  WHERE faId = " & intFAgreementId & " ORDER By CategoryCode"
	else
		strSQL = "SELECT * FROM FrameworkCategory  WHERE faId = 0 ORDER By CategoryCode"
	end if
	'Retrieve medarbeidere
	set rsFAgreements = GetFirehoseRS(strSQL, con)

	set fldFaCID = rsFAgreements.fields("CategoryID")
	Do Until rsFAgreements.EOF 
		'If (clng(intSelectedID) = clng(fldFaCID.value))  Then
		'	strSel = " selected"
		'Else
		strSel = ""
		'End If 
		strName = cstr(rsFAgreements.fields("CategoryCode").value) & " - " 
		If Not isNull(rsFAgreements.fields("CategoryName").value) Then
			strName = strName & cstr(rsFAgreements.fields("CategoryName").value)
		End If
		response.write "<option value=""" & fldFaCID.value & """ " & strSel &" >" & strName & "</option>"
		rsFAgreements.MoveNext
	Loop
	rsFAgreements.Close
	Set rsFAgreements = Nothing
	CloseConnection(con)
	set con = nothing			

end function

function SetEmptyCategoryVikar(intOppdragId)

	dim strSQL

	Set ConnTrans = GetClientConnection(GetConnectionstring(XIS, ""))
	ConnTrans.BeginTrans
	
	
	strSQL = "EXECUTE [spECSetEmptyVikarCategory] " & intOppdragId

	If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
		ConnTrans.RollbackTrans
		CloseConnection(ConnTrans)
		set ConnTrans = nothing
		AddErrorMessage("Feil Empty Oppdrag Vikar.")
		call RenderErrorMessage()
	Else
	
		ConnTrans.CommitTrans
		CloseConnection(ConnTrans)
		set ConnTrans = nothing
	
	End If

end function

%>