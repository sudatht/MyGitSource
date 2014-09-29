<!--#INCLUDE FILE="includes\Library.inc"-->
<%

' Open database connection
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

'Parametere
avdkontor = Request("avdkontor")
frakode = Request("frakode")
backURL = Request("backURL")

'	lagertimeliste i oslo
'----------------------------------------------------------------------------------


      strSQL = "Execute Lag_Timeliste " & Request("OppdragVikarID") &_
			"," & Request("OppdragID") &_
			"," & Request("VikarID")	&_
			"," & Request("FirmaID") &_
			"," & DbDate(Request("Fradato")) &_
			"," & DbDate(Request("TilDato")) &_
			"," & DbTime(Request("Frakl")) &_
			"," & DbTime(Request("Tilkl")) & _
			"," & Request("TimerPrDag")  &_
			"," & DbTime(Request("Lunch")) &_
			"," & Request("Timeloenn") &_
			"," & Request("Timepris") &_
			"," & Request("VikarType") &_
			"," & Request("Bestilltav") &_
			",1" 

'se strSQL & " oslo"

      conn.Execute( strSQL )

 ' Update status on OPPDRAG to FULLSTENDIG
    strSQL = "Update oppdrag set StatusID = 5 where oppdragid =" & Request("OppdragID")
 ' Update in database 
    Conn.Execute( strSQL )


	frakode = frakode + 1

'----------------------------------------------------------------------------------
'	tilbake
'----------------------------------------------------------------------------------

	redir = backURL & "Timeliste_mellom.asp" &_
	"?frakode=" & frakode 
'se redir & " oslo"
	Response.redirect redir

'----------------------------------------------------------------------------------

%>