<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes/Library.inc"--> 
<% 

' Do we have KontaktID ?
If Request.Querystring( "KontaktID" ) = "" Then
    Response.Write "Systemfeil: KontaktID mangler"
    Response.End
Else
   KontaktID = Request.Querystring( "KontaktID" )
End If

' Do we have FirmaID ?
If Request.Querystring( "FirmaID" ) = "" Then
   Response.Write "Systemfeil: FirmaID mangler"
   Response.End
Else
   FirmaID = Request.Querystring( "FirmaID" )
End If

' Do we have Dato ?
If Request.Querystring( "Dato" ) = "" Then
    Response.Write "Systemfeil: Dato mangler"
    Response.End
Else
   Dato = Request.Querystring( "Dato" )
End If
 
' Open database connection
' -------------------------
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("Xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("Xtra_CommandTimeout")
Conn.Open Session("Xtra_ConnectionString"), Session("Xtra_RuntimeUserName"), Session("Xtra_RuntimePassword")

' update navndag dato on kontaktperson
' ------------------------------------------

' Build SQl-statement
strSQL = "Update KONTAKT set navndag= " &  DbDate( Dato )  & " where kontaktid = " & KontaktID

' Update in database 
Conn.Execute( strSQL )

' Error from database ?
If Conn.Errors.Count > 0 then
   Call SqlError()
End if

' Redirect page
Response.redirect "kundevis.asp?FirmaID="& FirmaID

%>