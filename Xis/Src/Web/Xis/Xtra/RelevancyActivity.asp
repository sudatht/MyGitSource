<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>

<!--#INCLUDE FILE="includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="includes/xis.rights.inc"-->
<!--#INCLUDE FILE="includes/Xis.Settings.inc"-->
<%
	dim strActivity
	dim nActivityTypeID
	dim strVikarID
	dim strComment
	
	strActivity = "Oppfriskningsdato"
	strComment = "Oppfriskningsdato oppdatert"
	strVikarID = Request.QueryString("VikarID")
	
	' Get a database connection
	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))
	
	set rsActivityType = GetFirehoseRS("SELECT AktivitetTypeID FROM H_AKTIVITET_TYPE WHERE AktivitetType = '" & strActivity & "'", Conn)
		nActivityTypeID = CInt(rsActivityType("AktivitetTypeID"))
		' Close and release recordset
	      	rsActivityType.Close
	      	Set rsActivityType = Nothing
	      	
	      	sDate = GetDateNowString()
		strSql = "INSERT INTO AKTIVITET( AktivitetTypeID, AktivitetDato, VikarID, Notat, Registrertav, RegistrertAvID, AutoRegistered ) Values(" & nActivityTypeID & ",'" & sDate & "'," & strVikarID & ",'" & strComment & "','" & Session("Brukernavn") & "'," & CInt(Session("medarbID")) & ", 1)"
		
		If ExecuteCRUDSQL(strSql, Conn) = false then
			' Rollback transaction
			Conn.RollBackTrans
			AddErrorMessage("Aktivitetsregistrering på vikar feilet.")
			call RenderErrorMessage()
		End if
		
CloseConnection(Conn)
set Conn = nothing
		
Response.redirect "vikarvis.asp?VikarID=" & strVikarID
%>