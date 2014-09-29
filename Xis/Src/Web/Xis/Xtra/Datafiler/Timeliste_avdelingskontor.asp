<%

'skriptet som brukes på avdelingskontor for å se på timelsiter.

p = 1
Select case p
	Case 1
		Response.Redirect "http://intranett.xtra.no/xtra/datafiler/Vikar_Timeliste_vis3.asp?avdkontor=1&oppdragID=" & oppdragID & "&vikarid=" & vikarid & "&frakode=-1"
	Case 2
		Response.Redirect "Vikar_Timeliste_vis3.asp?oppdragID=" & oppdragID & "&vikarid=" & vikarid & "&frakode=1"
	Case 3
	Case Else
End Select

%>
