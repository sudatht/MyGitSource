<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_REPORT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim resportSessionVar

	if(isnull(Request.QueryString("report"))) then
		AddErrorMessage("Ingen rapport data funnet!")
		call RenderErrorMessage()	
	else
		resportSessionVar = session(Request.QueryString("report"))
	end if

	Response.Clear
	Response.ContentType = "text/csv"
	Response.AddHeader "Content-Disposition", "filename=report.csv;"

	Response.Write resportSessionVar

%>