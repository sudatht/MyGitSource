<%
function GetCoWorkersAsOptionList(intSelectedID)
	dim strSel		'Contains "SELECTED" if current coworkers is the selected one
	dim rsMedarbeider	'Recordset containing coworkers
	dim strSQL		'SQL statement to execute, retriving coworkers
	dim strName		'Name of coworker
	dim fldMedID		'field ref 
	dim conn		'Connection

	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
	Conn.CommandTimeOut = Session("xtra_CommandTimeout")
	Conn.Open Session("xtra_ConnectionString"),Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")
		
	if isnull(intSelectedID) then
		intSelectedID = 0
	end if
	
	
	if (clng(intSelectedID)>0) then
		strSQL = "select med.etternavn, med.fornavn, med.medID from medarbeider as med inner join vikar on med.vikarid = vikar.vikarid where vikar.statusID = 4 union select med.etternavn, med.fornavn, med.medID from medarbeider as med where med.medid = " & intSelectedID & " order by med.etternavn"
	else
		strSQL = "select med.etternavn, med.fornavn, med.medID from medarbeider as med inner join vikar on med.vikarid = vikar.vikarid where vikar.statusID = 4 order by med.etternavn"
	end if
	'Retrieve medarbeidere
	Set rsMedarbeider = Conn.Execute(strSQL)

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
	Set rsMedarbeider=Nothing
end function

%>