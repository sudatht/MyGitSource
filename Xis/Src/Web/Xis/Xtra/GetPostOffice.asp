<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.Settings.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
dim conn
dim rsPostOffice
dim PostSted
dim strSQL

Set Conn = GetConnection(GetConnectionstring(XIS, ""))	
strSQL= "SELECT sted FROM H_POSTNUMMER WHERE PostNr='"  & Request.Form("PostNr") & "'"

set rsContact = GetFirehoseRS(strSQL, conn)
if(HasRows(rsContact) = true) then
   PostSted = rsContact("sted")
end if
rsContact.close
set rsContact = nothing

Response.CharSet = "iso-8859-9"
Response.Write(PostSted)

%>