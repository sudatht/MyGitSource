<%@  language="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\MailLib.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.Settings.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="includes\Xis.Security.Utils.inc"-->
<!--#INCLUDE FILE="includes\DNN.Users.inc"-->
<%
'Delete Personal Information
'PRO@EC
'16/03/2007

If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
	call Response.Redirect("/xtra/IngenTilgang.asp")
end if

dim lngVikarID
dim brukerID
dim objCon
dim strVikarSelect
dim rsVikarDel

If  HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
     
    brukerID = Session("BrukerID")

	' Move parameters to local variables
	lngVikarID = Request.Querystring("VikarID")

	' Check VikarID
	If lngVikarID = "" Then
		AddErrorMessage("Feil: Mangler parameter vikarID. Noter navn på vikar og kontakt systemansvarlig.")
		call RenderErrorMessage()
	End If
	
	' Open database connection
    Set objCon = GetConnection(GetConnectionstring(XIS, ""))
        
	If lngVikarID <> "" Then
	    strVikarSelect = "SELECT " & _
	            "VIKAR."
	    '####################################################
	    set rsVikarDel = GetFirehoseRS(strVikarSelect, objCon)
	    
	    ' Close and release recordset
		rsVikar.close
		Set rsVikar = Nothing
	End IF
	
End IF
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
<head>
    <title>Personal Information Delete</title>
</head>
</html>
