<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.utils.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<% 
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	Dim IntTypeID			'Holds either 3 (qualification) or 4 (jobwish)
	Dim LngVikarId			'Holds consultant id
	Dim LngCvid				'Holds consultants CV id
	Dim intExperience
	Dim intEducation
	Dim lntRangering
	Dim RsCV
	Dim strSQL

	' Check VikarID
	If Request("tbxVikarID") = "" Then
	   AddErrorMessage("Feil: VikarID mangler!")
	   call RenderErrorMessage()
	else
		LngVikarId = Request("tbxVikarID")
	End If

	'Check rangering
	If Request("tbxRangering") = "0" Then
	   lntRangering = "NULL"
	Else
	   lntRangering = Request("tbxRangering")
	End If

	'Check experience
	If Request("tbxExperience") = "0" Then
	   intExperience = "NULL"
	Else
	   intExperience = Request("tbxExperience")
	End IF

	'Check Education
	If Request("tbxEducation") = "0" Then
	   intEducation = "NULL"
	Else
	   intEducation = Request("tbxEducation")
	End IF

	IntTypeID = clng(Request("tbxTypeID"))

	' Initializes AND  opens database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	
	
	'Check for CV!
	strSQL = "SELECT cvid FROM cv WHERE type='C' AND consultantid=" & LngVikarId
	set RsCV = GetFirehoseRS(strSQL, Conn)
	'Cv does not exist, create new
	if (HasRows(RsCV) = false) then
		strSQL = "INSERT INTO CV( consultantid, type) VALUES(" & LngVikarId & ", 'C')"
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			AddErrorMessage("En eller flere feil oppstod under oppretting av CV.")
			call RenderErrorMessage()
		End if
		StrSQL = "SELECT cvid FROM cv WHERE type='C' AND consultantid =" & LngVikarId
		set RsCV = GetFirehoseRS(strSQL, Conn)
	end if
	'Store CVid for later use..
	LngCvid = clng(RsCV.fields("Cvid").value)
	rsCV.Close
	set rsCV = nothing

	' Action against database depending on Button pressed
	If Request("tbxAction") = "lagre" AND LenB(Request("tbxKompetanseID")) <> 0 Then
		if IntTypeID = 3 then
			strSQL = "EXEC UpdateQualificationByID " & Request("tbxKompetanseID") & _
			", 0" & _
			", " & lntRangering & _
			", " & FixString(Request("tbxKommentar"))
 		elseif IntTypeID = 4 then
			' Create SQL-statement
				strSQL = "Update VIKAR_KOMPETANSE set " & _
				"Relevant_WorkExperience = " & intExperience & ", " & _
				"Relevant_Education = " & intEducation & ", " & _												
	           "kommentar = " & "'" & Request("tbxKommentar") & "'" & _  
			   " WHERE KompetanseID = " & Request("tbxKompetanseID")
		end if		   
	   ' Update kompetanse in database
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			AddErrorMessage("En eller flere feil oppstod under oppdatering av kompetanse.")
			call RenderErrorMessage()
		End if
	ElseIf lcase(Request("tbxAction")) = "slette" Then
		' Create SQL-statement
		strSQL = "DELETE FROM VIKAR_KOMPETANSE WHERE KompetanseID = " & Request("tbxKompetanseID")
		' Delete adress in database
		if (ExecuteCRUDSQL(strSQL, Conn) = false) then
			' Database error ?
			If Conn.Errors.Count > 0 Then
				AddErrorMessage("Feil oppstod under sletting av kompetanse.")
				call RenderErrorMessage()
			End If
		end if
	End If
	CloseConnection(Conn)
	set Conn = nothing	
	' Redirect to new page
	Response.Redirect "VikarVis.asp?VikarID=" & LngVikarId & "&Nop=nop#competance"
%>