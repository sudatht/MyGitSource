<%@ LANGUAGE="VBSCRIPT" %>
<%Option explicit%>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<%
	'File purpose:		Save the jobwishes recieved from vikarjobbonskeliste.asp,
	'					for the specified consultant.

	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	Dim Lvikarid				'Holds consultant id
	Dim RsCV					'holds recordset with consultants CV information
	Dim lAreaID					'Holds selected areaID 
	Dim LCvid					'Holds id to consultants CV
	Dim ObjConn					'as adodb.Connection
	Dim strSQL					'string to hold SQL
	Dim frmIntput				'holds reference form fields in loops
	dim intIdPos				'Position of id in form value
	dim StrCommentary			'Holds commentary to current Qualification in loop
	dim intTittelID				'Holds id to current Qualification in loop
	dim StrWhereAreaSQL			'Holds where SQL when a fagområde is choosen.
	dim IntExperienceRank
	dim IntEducationRank


	' Transfer Vikarid to var
	If (len(trim(Request("VikarID"))) > 0) Then
	   Lvikarid = CLng( Request("VikarID") )
	Else
	   Lvikarid = 0
	   Response.write "Error in Parameter. tbxVikarID has no value!"
	   Response.end
	End If

	lAreaID = Request("dbxOldArea")

	If Request("dbxEndret") <> "0" Then
		' Open database Connection
		Set ObjConn = GetConnection(GetConnectionstring(XIS, ""))
		'Check for CV!
		strSQL = "select cvid from cv where type='C' and consultantid=" & Lvikarid
		set RsCV = GetFirehoseRS(strSQL, ObjConn)

		'Cv does not exist, create new
		if RsCV.EOF then
			strsql = "INSERT INTO CV( consultantid, type) values(" & Lvikarid & ", 'C')"
			ObjConn.Execute( strSQL )
			If ObjConn.Errors.Count > 0 then
			Call SqlError()
			End if
			strSQL = "SELECT cvid FROM cv where type='C' AND consultantid=" & Lvikarid
			set RsCV = GetFirehoseRS(strSQL, ObjConn)
		end if
		LCvid = clng(RsCV.fields("Cvid").value)
		set rsCV = nothing

		StrSQL = "exec DropProfessionsForArea " & Lvikarid & "," & lAreaID 
		if (ExecuteCRUDSQL(StrSQL, ObjConn) = false) then
			If ObjConn.Errors.Count > 0 then
				Call SqlError()
			End if
		end if

		for each frmIntput in Request.Form
			if instr(1, frmIntput,"TittelID") then
				intTittelID = mid(frmIntput,1, instr(1, frmIntput,"TittelID")-1)
				IntExperienceRank	=	Request.Form.Item(intTittelID & "rdoErfaringniva")
				IntEducationRank	=	Request.Form.Item(intTittelID & "rdoUtdannelseNiva")				
				StrCommentary		=	Request.Form.Item(intTittelID & "tbxKommentar")
			
				if (isnull(IntExperienceRank)) OR (IntExperienceRank = 0) then
					IntExperienceRank = "NULL"
				end if	
				if (isnull(IntEducationRank)) OR (IntEducationRank = 0) then
					IntEducationRank = "NULL"
				end if	
				' Create SQL-statement
		        strSQL = "INSERT INTO [Vikar_Kompetanse]( [VikarID], [CVid], [K_TypeID], [K_TittelID], [kommentar],[Relevant_WorkExperience],[Relevant_Education]) " & _
						"Values(" &_
						Lvikarid & "," & _
						LCvid & "," & _
						"4," & _
						intTittelID & "," & _
						Quote(PadQuotes(StrCommentary)) & "," & _
						IntExperienceRank & "," & _				
						IntEducationRank  & ")"
				if (ExecuteCRUDSQL(strSQL, ObjConn) = false) then
					If ObjConn.Errors.Count > 0 then
						Call SqlError()
					End if
				end if
			end if
		next
		'Fred 111202, updated consultant's qualification lastupdated date
        strSQL = "UPDATE [VIKAR] SET kompetansedato= getdate() WHERE vikarID =" & Lvikarid
		call ExecuteCRUDSQL(strSQL, ObjConn)
		
		CloseConnection(ObjConn)
		set ObjConn = nothing
	end if

	Response.redirect request("dbxJobwishSource") & "?VikarID=" & Lvikarid & "&dbxArea=" & request("dbxarea") & "&dbxShowAll=" & Request("dbxShowAll")
%>