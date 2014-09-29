<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.Settings.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	' Check input values
	If Trim( Request.form("tbxadress") ) = "" Then
	Response.Write "Adresse mangler"
	Response.End
	End If

	' Open database connection
	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))	

	' Start transaction
	Conn.Begintrans

	' Action against database depending on Button pressed
	If Request("pbnDataAction") = "Nullstill" Then 

		' Restart adress with no AdrID to clear all fields 
		Response.redirect "adresse.asp?Relasjon=" & Request.form("tbxRelasjon") & "&ID=" & Request.form("tbxID")

	ElseIf Request("pbnDataAction") = "Lagre" and Request.form("tbxAdrID") = "" Then 

		' Create SQL-statement
		strSQL = "Insert into ADRESSE(AdresseRelID, AdresseRelasjon, Adresse, Postnr, Poststed, AdresseType ) " & _
			"Values( " & Request.form("tbxID") & "," & _
			Request.form("tbxRelasjon") & "," & _
			"'" & Request.form("tbxadress") & "'," & _
			Quote( Request.form("tbxPostnr") ) & "," & _
			Quote( Request.form("tbxPoststed") ) & "," & _
			Request.form("dbxAdrType") & ")" 

		' Create new adress in database
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			' Rollback transaction
				Conn.RollBackTrans
				AddErrorMessage("En eller flere feil oppstod under oppretting av adresse for vikar.")
				call RenderErrorMessage()
		End if

	ElseIf Request("pbnDataAction") = "Lagre" and Request.form("tbxAdrID") <> "" Then

		' Create SQL-statement
		strSQL = "Update ADRESSE set Adresse = " & "'" & Request.form("tbxAdress") & "'" & _
			", Postnr =" & Quote(Request.form("tbxPostNr")) & _
			", Poststed =" & Quote( Request.form("tbxPoststed") ) & _
			", AdresseType =" & Request.form("dbxAdrType") & _
			" where AdrID = " & Request.form("tbxAdrID")

		' Update adress in database
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			' Rollback transaction
				Conn.RollBackTrans
				AddErrorMessage("En eller flere feil oppstod under endring av adresse.")
				call RenderErrorMessage()
		End if

	ElseIf Request("pbnDataAction") = "Slette" Then

		' Create SQL-statement
		strSQL = "Delete Adresse where Adrid = " & Request.form("tbxAdrID")

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			' Rollback transaction
			Conn.RollBackTrans
			AddErrorMessage("En eller flere feil oppstod under sletting av adresse.")
			call RenderErrorMessage()
		End if
	Else
		AddErrorMessage("Feil i adresse parameter.")
		call RenderErrorMessage()	
	End If

	If Request.form("tbxRelasjon") = 2 Then
		' Adress connect to Vikar
		strRedirectURL = "vikarvis.asp?VikarID="
	Else
		' Error in adress relasjon
		AddErrorMessage("Error on Relasjon.")
		call RenderErrorMessage()	
	End If

	Conn.CommitTrans
	CloseConnection(Conn)
	set Conn = nothing

	' Redirect to new page
	Response.Redirect strRedirectURL & Request.form("tbxID")
%>
