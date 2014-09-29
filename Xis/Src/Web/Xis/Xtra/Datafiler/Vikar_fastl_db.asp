<%@ LANGUAGE="VBSCRIPT" %>
<%
'Response.write Request.Form("Navn") & "<br>"

'<H3>Lagring av faste lønnsdata.</H3>

Sub fjernKomma(strString )
  pos = Instr(strString, ",")
  If pos <> 0 Then
	 mellom = Left(strString, pos-1) & "." & Mid(strString, pos+1)
 	strString = mellom 
  End If
'Response.Write strString
End Sub

'--------------------------------------------------------------------------------------------------
' Connect to database
'--------------------------------------------------------------------------------------------------

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

'--------------------------------------------------------------------------------------------------
' Check and put into variables
'--------------------------------------------------------------------------------------------------
If Request.Form("VikarID") = "" Then
   	strVikarID = Request.QueryString("VikarID")
Else
   	strVikarID = Request.Form("VikarID")
End IF
If Request.Form("Avdeling") = "" Then
	strAvdeling = Request.QueryString("Avdeling")
Else
	strAvdeling = Request.Form("Avdeling")
End If
strLoennsart = Request.Form("Loennsartnr")
strAntall = Request.Form("Antall")
strSats = Request.Form("Sats")
strBeloep = Request.Form("Beloep")
strSaldo = Request.Form("Saldo")
strNavn = Request.Form("Navn")

If Request.QueryString("ID") = "" Then
	strID = Request.Form("ID")
Else
	strID = Request.QueryString("ID")
End If

strSlett = Request.QueryString("Slett")

strEndre = Request.Form("Endre")

'Response.write strVikarID & "<br>"
'Response.write Request.Form("Navn") & "<br>"
'Response.write strAvdeling & "<br>"
'Response.write strLoennsart & "<br>"
'Response.write strAntall & "<br>"
'Response.write strSats & "<br>"
'Response.write strBeloep & "<br>"
'Response.write strSaldo & "<br>"
'Response.write strSlett & "<br>"
'Response.write strEndre & "<br>"
'Response.write strID & "<br>"


'--------------------------------------------------------------------------------------------------
' Delete row in VIKAR_LOEN_FASTE
'--------------------------------------------------------------------------------------------------
If strSlett <> "" Then

strSQL = "Delete from VIKAR_LOENN_FASTE where ID = "  & strID

'Response.write strSQL

conn.Execute(strSQL)

End IF 'deleting



If strLoennsart <> "" And strAntall <> "" And strSats <> "" Then

strSaldo = strAntall * strSats
call fjernKomma(strSaldo)
strBeloep = 0
'--------------------------------------------------------------------------------------------------
' Update row in VIKAR_LOEN_FASTE
'--------------------------------------------------------------------------------------------------
If strEndre = "Ja" Then

strBeloep = strSats * strAntall
Call fjernKomma(strBeloep)

strSQL = "Update VIKAR_LOENN_FASTE set" &_
        " Loennsart = '" & strLoennsart &_
	"', Antall = " & strAntall &_
	", Sats = " & strSats &_
	", Beloep = " & strBeloep &_
	", Saldo = " & strSaldo &_
	" where ID = " & strID

'Response.write strSQL

conn.Execute(strSQL)



'--------------------------------------------------------------------------------------------------
' Insert into database VIKAR_LOEN_FASTE
'--------------------------------------------------------------------------------------------------
Else 'insert (not delete or update)

strBeloep = strSats * strAntall
Call fjernKomma(strBeloep)

strSQL = "Insert into VIKAR_LOENN_FASTE (VikarID, Avdeling, Loennsart, Antall, Sats, Beloep, Saldo) " &_
	"values (" &_
		strVikarID & "," &_
		"1,'" &_
		strLoennsart & "'," &_
		strAntall & "," &_
		strSats & "," &_
		strBeloep & "," &_
		strSaldo & ")"

'Response.write strSQL

conn.Execute(strSQl)



%>
<% End If 'update or insert %>
<% End If 'filds have content %>

<% 	Response.Redirect "Vikar_fastl_vis.asp?VikarID=" & strVikarID & "&Avdeling=" & strAvdeling %>



		