<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	 
	' Move parameters to local variables
	lOppdragID    = CLng( Request.Querystring("OppdragID") )
	lKompetanseID = CLng( Request.Querystring("KompetanseID") )
	lTypeID       = CLng( Request.Querystring("TypeID") )

	' Check OppdragID
	If lOppdragID = "" Then
	Response.Write "<p class'warning'>Error: Parameter mangler!</p>"
	Response.End
	End If

' Open database connection
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

' Get oppdrag info
Set rsoppdrag = Conn.Execute("select OppdragID, Firma from oppdrag, Firma where oppdragID = " & lOppdragID & " and Oppdrag.FIRMAID = FIRMA.FIRMAID")

' No records found ?
If rsoppdrag.BOF = True And rsoppdrag.EOF = True Then 
   ' Write message
   Response.Write "<p class'warning'>Error: Feil i Parameter</p>"

   ' Abort script 
   Response.End
End If

' Move from recordset to variables
stroppdragName = rsoppdrag("OppdragID") & "  " & rsoppdrag("Firma")

' Close and release recordset
rsOppdrag.Close
Set rsOppdrag = Nothing

' Build heading
strHeading = "Kompetanse for oppdragnr " & stroppdragName

' Existing Kompetanse ?
If lKompetanseID <> "" Then
   Set rsKompetanse = Conn.Execute("Select K_TypeID, K_LevelID, K_TittelID, Beskrivelse  from OPPDRAG_KOMPETANSE where K_OppdrID = " & lKompetanseID )

   ' No records found ?
   If rsKompetanse.BOF = True And rsKompetanse.EOF = True Then 

      ' Write message
      Response.Write "<p class'warning'>Error: Feil i Parameter</p>"

      ' Abort script 
      Response.End
   End If

   ' move Kompetanse to variables
    lKTypeID    = rsKompetanse("K_TypeID") 
   lTittelID    = rsKompetanse("K_TittelID")
   lLevelID    = rsKompetanse("K_LevelID")
   strBeskrivelse = rsKompetanse("Beskrivelse")


   ' Close and release recordset
   rsKompetanse.Close
   Set rsKompetanse = Nothing

End If
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"
    "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <meta http-equiv="Content-Style-Type" content="text/css">
    <meta http-equiv="Content-Script-Type" content="text/javascript">
    <meta name="Developer" content="Electric Farm ASA">
	<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
	<title>Kompetanse</title>
</head>
<body>
	<div class="pageContainer" id="pageContainer">
		<div class="contentHead1">
			<h1><% =strHeading %></h1>
		</div>
		<div class="content">
			<form name="formEn" ACTION="oppdragkompdb.asp" METHOD="POST">
				<input TYPE="HIDDEN" NAME="tbxoppdragID" VALUE="<% =lOppdragID  %>">
				<input TYPE="HIDDEN" NAME="tbxKompetanseID" VALUE="<% =lKompetanseID %>">
				<div class="listing">
				<table>
					<tr>
						<th>Type:</th>
						<th>Kursnivå:</th>
						<th>Kommentar:</th>
					</tr>
					<tr>
						<td>
							<select NAME="dbxKompetanseTittel">
								<option VALUE="0">
								<% 
								' Get KompetanseTittel
								Set rsKompetanseTittel = Conn.Execute("select K_TittelID, KTittel from H_KOMP_TITTEL where K_TypeID = " & lTypeID )
								
								Do Until rsKompetanseTittel.EOF
								   If rsKompetanseTittel("K_TittelID") = lTittelID Then
								      strSelected = rsKompetanseTittel("K_TittelID") & " " & "SELECTED"
								   Else
								      strSelected = rsKompetanseTittel("K_TittelID")
								   End If
								%>
												<option VALUE="<% =strSelected %>"><% =rsKompetanseTittel("KTittel") %>
								<% 
								       rsKompetanseTittel.MoveNext
								   Loop
								
								   ' Close and release recordset
								   rsKompetanseTittel.Close
								   Set rsKompetanseTittel = Nothing
								
								%>
							</select>						
						</td>
						<td>
							<select NAME="dbxLevel">
								<option VALUE="0">
								<% 
								' Get Kompetanse level
								Set rsLevel = Conn.Execute("select K_LevelID, KLevel from H_KOMP_LEVEL" )
								
								Do Until rsLevel.EOF
								   If rsLevel("K_LevelID") = lLevelID Then
								      strSelected = rsLevel("K_LevelID") & " " & "SELECTED"
								   Else
								      strSelected = rsLevel("K_LevelID")
								   End If
								%>
												<option VALUE="<% =strSelected %>"><% =rsLevel("KLevel") %>
								<% 
								   rsLevel.MoveNext
								Loop
								
								   ' Close and release recordset
								   rsLevel.Close
								   Set rsLevel = Nothing
								
								%>
							</select>						
						</td>
						<td><input NAME="tbxBeskrivelse" TYPE="TEXT" SIZE="30" MAXLENGTH="255" VALUE="<% =strBeskrivelse %>"></td>
					</tr>
				</table>
				</div>


<% If lKompetanseID ="" Then
   Response.write "<INPUT NAME=pbnDataAction TYPE=SUBMIT  VALUE=Nullstill>"
   Response.write "<INPUT NAME=pbnDataAction TYPE=SUBMIT  VALUE=Lagre>"
Else
   Response.write "<INPUT NAME=pbnDataAction TYPE=SUBMIT  VALUE=Nullstill>"
   Response.write "<INPUT NAME=pbnDataAction TYPE=SUBMIT  VALUE=Lagre>"
   Response.write "<INPUT NAME=pbnDataAction TYPE=SUBMIT  VALUE=Slette>"
End If
%>

			</form>
	    </div>
	</div>
</body>
</html>

