<%

	function GetFirmId(SOcuID)
		dim con 
		dim rs
		set con = GetConnection(GetConnectionstring(XIS, ""))	
		set rs = GetDynamicRS("SELECT [FirmaID] FROM [Firma] WHERE [SOCuID]=" & SOcuID, con)
		GetFirmId = 0
		
		if HasRows(rs) then
			GetFirmId = rs("FirmaID").value
		end if
		
		CloseConnection(con)
		set rs = nothing
		set con = nothing
	end function

%>