<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\Library.inc"-->

<% 

' Check input values
If Trim( Request.form("tbxadress") ) = "" Then
   Response.Write "Adresse mangler"
   Response.End
End If

' Do not stop when error occurs
On Error Resume Next

' Open database connection
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("Xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("Xtra_CommandTimeout")
Conn.Open Session("Xtra_ConnectionString"), Session("Xtra_RuntimeUserName"), Session("Xtra_RuntimePassword")

' 
' Action against database depending on Button pressed
' 
If Request("pbnDataAction") = "Nullstill" Then 

   ' Restart adress with no AdrID to clear all fields 
   Response.redirect "suspectadresse.asp?Relasjon=" & Request.form("tbxRelasjon") & "&ID=" & Request.form("tbxID")

   ' Abort script 
   Response.End

ElseIf Request("pbnDataAction") = "Lagre" and Request.form("tbxAdrID") = "" Then 

  ' Create SQL-statement
  strSQL = "Insert into SUSPECT_ADRESSE(AdresseRelID, AdresseRelasjon, Adresse, Postnr, Poststed, AdresseType ) " & _
           "Values( " & Request.form("tbxID") & "," & _
		   Request.form("tbxRelasjon") & "," & _
		   "'" & Request.form("tbxadress") & "'," & _
		   Quote( Request.form("tbxPostnr") ) & "," & _
		   Quote( Request.form("tbxPoststed") ) & "," & _
		   Request.form("dbxAdrType") & ")" 

   ' Create new adress in database
   Conn.Execute( strSQL )
  
   ' Database error ?
   If Conn.Errors.Count > 0 Then

      ' Write error message
      Call SqlError()
   End If

ElseIf Request("pbnDataAction") = "Lagre" and Request.form("tbxAdrID") <> "" Then

  ' Create SQL-statement
  strSQL = "Update SUSPECT_ADRESSE set Adresse = " & "'" & Request.form("tbxAdress") & "'" & _
                                      ", Postnr =" & Quote(Request.form("tbxPostNr")) & _
                                     ", Poststed =" & Quote( Request.form("tbxPoststed") ) & _
		  ", AdresseType =" & Request.form("dbxAdrType") & _
		  " where AdrID = " & Request.form("tbxAdrID")

   ' Update adress in database
   Conn.Execute( strSQL )

   ' Database error ?
   If Conn.Errors.Count > 0 Then

      ' Write error message
      Call SqlError()
   End If

ElseIf Request("pbnDataAction") = "Slette" Then

  ' Create SQL-statement
  strSQL = "Delete SUSPECT_ADRESSE where Adrid = " & Request.form("tbxAdrID")

   ' Delete adress in database
   Conn.Execute( strSQL )

   ' Database error ?
   If Conn.Errors.Count > 0 Then

      ' Write error message
      Call SqlError()
   End If

Else
   ' Error in parameter 
   Response.Write "Error in Parameter"  

   ' Abort script 
   Response.End 
End If

' Redirect to new page
Response.Redirect "suspectvis.asp?FirmaID=" & Request.form("tbxID")
%>
