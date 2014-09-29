<%
	Function GetSOPersonName(SOpersonID)
		dim cts
		dim personRs
		
		Set cts = server.CreateObject("Integration.SuperOffice")		
		set personRs = cts.GetPersonSnapshotById(clng(SOpersonID))
		if (HasRows(personRs)) then
			if (isnull(personRs("middlename"))) then
				GetSOPersonName = personRs("firstname") & " " & personRs("lastname")
			else
				GetSOPersonName = personRs("firstname") & " " & personRs("middlename") & " " & personRs("lastname")
			end if
			personRs.close
		else
			GetSOPersonName = ""
		end if
		set personRs = nothing	
	End Function

	private function GenerateSOContactURL(lngFirmaID)
		dim returnValue
		returnValue =  SUPEROFFICE_PROTOCOL_URL & SUPEROFFICE_PANEL_CONTACT_URL & "?" & SUPEROFFICE_CONTACT_ID & "=" & "3"
		GenerateSOContactURL = returnValue
	end function

%>