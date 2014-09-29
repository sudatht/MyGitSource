<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes/Library.inc"--> 

<% 
' Do not stop when error occurs
On Error Resume Next

' Check oppdragID
If Request.Form("tbxOppdragID") = "" Then
   Response.Write "Error: Parameter mangler (oppdragID)"
   Response.End
End If

' Open database connection
' ---------------------------
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("Xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("Xtra_CommandTimeout")
Conn.Open Session("Xtra_ConnectionString"), Session("Xtra_RuntimeUserName"), Session("Xtra_RuntimePassword")

' Action against database depending on Button pressed
' -----------------------------------------------------
If TRim( Request("pbnDataAction") ) = "Nullstill" Then 

   ' Restart oppdragkomp with no ID to clear all fields 
   Response.redirect "oppdragkomp.asp?oppdragID=" & Request.form("tbxoppdragID")

   ' Abort script 
   Response.End

ElseIf Trim( Request("pbnDataAction") ) = "Lagre" and Request.form("tbxKompetanseID") = "" Then 

  ' Create SQL-statement
  strSQL = "Insert into OPPDRAG_KOMPETANSE(oppdragID, K_TypeID, K_TittelID, K_LevelID, Beskrivelse ) " & _
           "Values( " &_
		   Request.form("tbxoppdragID") & "," & _
		   Request.form("tbxKTypeID") & "," & _
		   Request.form("dbxKompetanseTittel") & "," & _
		   Request.form("dbxLevel") & "," & _
		   Quote( Request.form("tbxBeskrivelse") ) & ")"

   ' Create new adress in database
   Conn.Execute( strSQL )
  
   ' Database error ?
   If Conn.Errors.Count > 0 Then

      ' Write error message
      Call SqlError()
   End If

ElseIf Trim( Request("pbnDataAction") ) = "Lagre" and Request.form("tbxKompetanseID") <> "" Then

  ' Create SQL-statement
  strSQL = "Update OPPDRAG_KOMPETANSE set K_TittelID = " & Request.form("dbxKompetanseTittel") & _
           ", K_LevelID = " & Request.form("dbxLevel") & _  
           ", beskrivelse = " & Quote( Request.form("tbxBeskrivelse") ) & _  
		   " where K_OppdrID = " & Request.form("tbxKompetanseID")
   
   ' Update adress in database
   Conn.Execute( strSQL )

   ' Database error ?
   If Conn.Errors.Count > 0 Then
      Call SqlError()
   End If

ElseIf Trim( Request("pbnDataAction") ) = "Slette" Then

  ' Create SQL-statement
  strSQL = "Delete OPPDRAG_KOMPETANSE where K_OppdrID = " & Request.form("tbxKompetanseID")

   ' Delete adress in database
   Conn.Execute( strSQL )

   ' Database error ?
   If Conn.Errors.Count > 0 Then
      Call SqlError()
   End If

Else
   ' Error in parameter 
   Response.Write "Error in Parameter"  

   ' Abort script 
   Response.End 
End If

' Redirect to new page
Response.Redirect "WebUI/OppdragView.aspx?OppdragID=" & Request.form("tbxoppdragID")
%>
