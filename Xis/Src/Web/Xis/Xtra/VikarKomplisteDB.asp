<%@ LANGUAGE="VBSCRIPT" %>
<%Option explicit%>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<%
	'File purpose:		Save the qualification data recieved from vikarkompliste.asp,
	'					for the specified consultant.

	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	Dim Lvikarid			'Holds consultant id
	Dim RsCV				'holds recordset with consultants CV information
	Dim LCvid				'Holds id to consultants CV
	Dim ObjConn				'as adodb.Connection
	Dim strSQL				'string to hold SQL
	Dim frmIntput			'holds reference form fields in loops
	dim intIdPos			'Position of id in form value
	dim IntRank				'Rank of current Qualification in loop
	dim StrCommentary		'Holds commentary to current Qualification in loop
	dim IntLevel			'Level of current Qualification
	dim intTittelID			'Holds id to current Qualification in loop
	dim StrWhereAreaSQL		'Holds where SQL when a productgroup is choosen.

	' Transfer Vikarid to var
	If Request.Form("VikarID") <> "" Then
		Lvikarid = CLng( Request.Form("VikarID") )
	Else
		Lvikarid = 0
		Response.write "Error in Parameter. tbxVikarID has no value!"
		Response.end
	End If

	If clng(Request.Form("dbxOldArea")) > 0 Then
		StrWhereAreaSQL = " and K_TittelID in (select K_TittelID from h_komp_tittel where ProdOmradeID =" & clng(Request.Form("dbxOldArea")) & ")"
	Else
		StrWhereAreaSQL = ""
	End If

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

		StrSQL = "DELETE FROM vikar_kompetanse WHERE K_TypeID = 3 AND vikarid=" & Lvikarid & StrWhereAreaSQL
		if (ExecuteCRUDSQL(StrSQL, ObjConn) = false) then
			If ObjConn.Errors.Count > 0 then
				Call SqlError()
			End if
		end if

		for each frmIntput in Request.Form
			if instr(1, frmIntput, "TittelID") then
				intTittelID = mid(frmIntput,1, instr(1, frmIntput,"TittelID")-1)
				IntRank			=	Request.Form.Item(intTittelID & "rdorangering")
				StrCommentary	=	Request.Form.Item(intTittelID & "tbxKommentar")
				IntLevel		=	Request.Form.Item(intTittelID & "dbxLevel")

				if (isnull(IntRank)) OR (IntRank = 0) then
					IntRank = "NULL"
				end if
				' Create SQL-statement
				strSQL = "Insert into Vikar_Kompetanse( VikarID, CVid, K_LevelID, Rangering, K_TypeID, K_TittelID, kommentar) " & _
						"Values(" &_
						Lvikarid & "," & _
						LCvid & "," & _
						IntLevel & "," & _
						IntRank & "," & _
						"3," & _
						intTittelID & "," & _
						Quote(PadQuotes(StrCommentary)) & ")"
				if (ExecuteCRUDSQL(strSQL, ObjConn) = false) then
					If ObjConn.Errors.Count > 0 then
						Call SqlError()
					End if
				end if
			end if
		next
		strSQL = "update VIKAR set " &_
			"Kompetansedato = " & dbDate(Date) &_
			" where VikarID = " &Lvikarid

		call ExecuteCRUDSQL(strSQL, ObjConn)
		
		CloseConnection(ObjConn)
		set ObjConn = nothing
	End If

	Response.Redirect request("dbxQualificationSource") & "?VikarID=" & Lvikarid & "&dbxArea=" & Request("dbxArea")& "&dbxShowAll=" & Request("dbxShowAll")
%>