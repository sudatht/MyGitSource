<%
	Server.ScriptTimeout = 6000
	dim SOPeID
	dim teller
	Dim cts 
	dim personRs
	dim flushInterval : flushInterval = 100
	dim numberofIterations : numberofIterations = 500
	dim persons(6)
	dim selected
	persons(1) = 2742
	persons(2) = 2743
	persons(3) = 2744
	persons(4) = 2745
	persons(5) = 2746
	persons(6) = 2747

	if (Request("Loop") <> "") then
		numberofIterations = Request("Loop")
	end if

	Randomize

	for teller = 1 to numberofIterations
		
		selected = clng((6 * Rnd)) 
		SOPeID = persons(selected)
		
		Set cts = server.CreateObject("Integration.SuperOffice")
		set personRs = cts.GetPersonSnapshotById(clng(SOPeID))
		if (not personRs.EOF) then
			if (isnull(personRs("middlename"))) then
				strKontaktperson = personRs("firstname") & " " & personRs("lastname")
			else
				strKontaktperson = personRs("firstname") & " " & personRs("middlename") & " " & personRs("lastname")
			end if
		end if
		set personRs = nothing
		set cts = nothing	

		Set cts = server.CreateObject("Integration.SuperOffice")
		personsHTML = cts.HTMLGetPersonsForContactAsDropDown(clng(720), 0, "dbxSOKontaktP" & teller, "(Ingen valgt)", "0", " class='mandatory' ", False)
		set cts = nothing	

		if (teller mod flushInterval = 0) then
			Response.Write "Current:" & teller & "<br>"
			Response.Write SOPeID & " : " & strKontaktperson & "<br>"			
			Response.Write personsHTML & "<br>"
			Response.Flush
		end if
	next
%>
