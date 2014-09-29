<html>
<head>
	<title>Overføring til Hult & Lillevik</title>
</head>
<%

If Request.QueryString("VikarID") = "" Then
	strVikarID = Request.Form("VikarID")
	tilgang = Request.Form("tilgang")
	kode = Request.Form("kode")
	strOppdragID = Request.Form("OppdragID")
	strFirmaID = Request.Form("FirmaID")
Else
	strVikarID = Request.QueryString("VikarID")
	tilgang = Request.QueryString("tilgang")
	kode = Request.QueryString("kode")
	strOppdragID = Request.QueryString("OppdragID")
	strFirmaID = Request.QueryString("FirmaID")
End If 

%>
<frameset cols="300" framespacing="0" frameborder="0">
	<frame src="Vikarvis2.asp?VikarID=" frameborder="0" scrolling="Auto" marginwidth="0" marginheight="0"><% =strVikarID %>&kode=<% =kode %>&tilgang=<% =tilgang %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %> >
<% if tilgang = 1 Or tilgang = 2 Then %>
	<frame src="Vikar_timeliste_vis.asp?VikarID=" name="RIGHT_WINDOW" id="RIGHT_WINDOW" frameborder="0" scrolling="Auto" marginwidth="0" marginheight="0"><% =strVikarID %>&kode=<% =kode %>&tilgang=<% =tilgang %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %> >
<% Elseif tilgang = 3 And kode = 5 Then %>
	<frame src="Vikar_timeliste_vis.asp?VikarID=" name="RIGHT_WINDOW" id="RIGHT_WINDOW" frameborder="0" scrolling="Auto" marginwidth="0" marginheight="0"><% =strVikarID %>&kode=<% =kode %>&tilgang=<% =tilgang %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %> >
<% Else %>
	<frame src="tomt.htm" name="RIGHT_WINDOW" id="RIGHT_WINDOW" frameborder="0" scrolling="No" noresize marginwidth="0" marginheight="0">
<% End If %>
</FRAMESET>
</html>