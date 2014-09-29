<%
	dim ObjSusp			'as xtraweb.suspect
	dim ObjCons			'as xtraweb.consultant
	dim ObjEdu			'as xtraweb.education
	dim ObjExp			'as xtraweb.experience
	dim ObjRef			'as xtraweb.reference
	dim allEdu			'as xtraweb.educations
	dim allExp			'as xtraweb.experiences
	dim allRef			'as xtraweb.references
	dim	ObjJobApp		'as xtraweb.jobapplication
	dim LConsCVid		'as long
	dim LSuspCVid		'as long
	dim allFiles
	dim ObjFile
	

	set ObjRef = Server.CreateObject("XtraWeb.Reference")
	set ObjExp = Server.CreateObject("XtraWeb.Experience")
	set ObjEdu = Server.CreateObject("XtraWeb.Education")
	set ObjCV = Server.CreateObject("XtraWeb.CV")

	set ObjCons = Server.CreateObject("XtraWeb.Consultant")
	ObjCons.XtraConString = Application("Xtra_intern_ConnectionString")
	ObjCons.XtraDataShapeConString = Application("ConXtraShape")

	set ObjSusp = Server.CreateObject("XtraWeb.Suspect")
	ObjSusp.XtraConString = ObjCons.XtraConString
	ObjSusp.XtraDataShapeConString = ObjCons.XtraDataShapeConString

	if (not ObjSusp.GetSuspect(LsuspectID)) then
		ObjSusp.CV.cleanup
		ObjSusp.cleanup
		set ObjSusp = nothing
		AddErrorMessage("Fant ikke suspect. Cv ikke overført!")
		CloseConnection(objCon)
		set objCon = nothing		
		call RenderErrorMessage()
	end if

	if (not objCons.GetConsultant(NewVikarID)) then
		ObjSusp.CV.cleanup
		ObjSusp.cleanup
		set ObjSusp = nothing

		Objcons.CV.cleanup
		Objcons.cleanup
		set ObjCons = nothing
		
		AddErrorMessage("Fant ikke vikar. Cv ikke overført!")
		CloseConnection(objCon)
		set objCon = nothing		
		call RenderErrorMessage()
	end if

	ObjCons.DataValues("foererkort") = ObjSusp.DataValues("foererkort")
	ObjCons.DataValues("bil") = ObjSusp.DataValues("bil")
	ObjCons.save

	set ObjCV  = ObjSusp.CV

	ObjCV.XtraConString = ObjSusp.XtraConString
	ObjCV.XtraDataShapeConString = ObjSusp.XtraDataShapeConString

	ObjCV.Refresh
	LSuspCVid = ObjCV.DataValues("cvid").value

	set allEdu = ObjCV.Educations
	set allExp = ObjCV.Experiences
	set allRef = ObjCV.References
	set allFiles = ObjCV.UplaodFiles 'ska

	'Saving CV as new consultant CV
	set ObjCV.owner = Objcons
	ObjCV.Save

	LConsCVid = ObjCV.DataValues("cvid").value
	Lvikarid = ObjCons.DataValues("vikarid").value

	'Tranfer all Qualifications & jobwishes
	strSQL = "UPDATE [vikar_kompetanse] SET [cvid] = "& LConsCVid & ", [vikarid] = " & Lvikarid & " WHERE [cvid] = " & LSuspCVid

	If ExecuteCRUDSQL(strSQL, objCon) = false then
		CloseConnection(objCon)
		set objCon = nothing
		AddErrorMessage("Feil oppstod under overføring til ansvarlig.")
		call RenderErrorMessage()
	End if

	'Tranfer all educations
	set allEdu.owner = ObjCV
	for each ObjEdu in allEdu
		ObjEdu.save
	next
	
	
	'Transfer all courses
	strSQL = "UPDATE [cv_data] SET [cvid] = "& LConsCVid & " WHERE [cvid] = " & LSuspCVid & " AND [FieldType] = 'COU' "

	If ExecuteCRUDSQL(strSQL, objCon) = false then
		CloseConnection(objCon)
		set objCon = nothing
		AddErrorMessage("Feil oppstod under overføring til ansvarlig.")
		call RenderErrorMessage()
	End if
	

	'Tranfer all experiences
	set allExp.owner = ObjCV
	for each ObjExp in allExp
		ObjExp.save
	next

	set allRef.owner = ObjCV
	'Tranfer all references
	for each ObjRef in allRef
		ObjRef.datavalues.remove "referenceid"
		ObjRef.datavalues("cvid").value = ObjCV.datavalues("cvid").value
		ObjRef.save
	next

    'ska Tranfer all upload files
    set allFiles.owner = ObjCV
    
    for i=1 to allFiles.Count
        allFiles.Item(i).DataValues.remove "UploadID"
        allFiles.Item(i).datavalues("cvid").value = ObjCV.datavalues("cvid").value
        allFiles.Item(i).save
    next
    
    'for each ObjFile in allFiles
    '    ObjFile.datavalues.remove "Filename"
	'	 ObjFile.datavalues("cvid").value = ObjCV.datavalues("cvid").value
    '    ObjFile.save    
    'next
    
	'Slette suspect jobb referanser
	strSQL = "DELETE FROM CV_References WHERE cvid = "& LSuspCVid
	if ExecuteCRUDSQL(strSQL, objCon) = false then
		CloseConnection(objCon)
		set objCon = nothing
		AddErrorMessage("Feil under sletting av referanser.")
		call RenderErrorMessage()
	End if
	
	'Slette suspect CV_Upload
	strSQL = "DELETE FROM CV_Upload WHERE cvid = "& LSuspCVid
	if ExecuteCRUDSQL(strSQL, objCon) = false then
		CloseConnection(objCon)
		set objCon = nothing
		AddErrorMessage("Feil under sletting av CV_Upload.")
		call RenderErrorMessage()
	End if

	'Slette suspect cv
	strSQL = "DELETE FROM CV WHERE cvid = " & LSuspCVid

	if ExecuteCRUDSQL(strSQL, objCon) = false then
		CloseConnection(objCon)
		set objCon = nothing
		AddErrorMessage("Feil under sletting av suspect CV.")
		call RenderErrorMessage()
	End if

	set ObjRef	= nothing
	set ObjExp	= nothing
	set ObjEdu	= nothing
	set ObjFile = nothing
	set allEdu	= nothing
	set allExp	= nothing
	set allRef	= nothing
	set allFiles = nothing
	set ColJobgroups = nothing
	set Objcv	= nothing

	ObjSusp.CV.cleanup
	ObjSusp.cleanup
	set ObjSusp = nothing

	Objcons.CV.cleanup
	Objcons.cleanup
	set ObjCons = nothing
%>