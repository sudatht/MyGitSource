<%
	Function GetSetting(strSettingName)
		Dim rsSetting
		Dim Conn
		
		strSQL = "SELECT SettingValue FROM Intra_settings WHERE settingName = '" & trim(strSettingName) & "'"

		Set Conn = GetConnection(GetConnectionstring(XIS, ""))	
		set rsSetting = GetFirehoseRS(strSQL, Conn)		
		if (rsSetting.EOF = false) then
			GetSetting = rsSetting.fields("SettingValue").value
		else
			Response.Write "Følgende innstilling/tekst mangler i Intrasettings:" & strSettingName &  "<br>"
			Response.end
		end if
		rsSetting.close
		set rsSetting = nothing	
		CloseConnection(Conn)
		set Conn = nothing				
	End function
%>