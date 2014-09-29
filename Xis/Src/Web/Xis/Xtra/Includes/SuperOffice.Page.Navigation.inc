<%
	function CreateSubPageURL(defaultpage)
		dim first : first = true
		if (len(Request("subPage"))> 0) then
			dim returnVal : returnVal = Request("subPage") & "?"
			for each x in Request.QueryString
				if (x <> "subPage") then
					if (first) then
						returnVal = returnVal & x & "=" & Request.QueryString(x)
						first = false
					else
						returnVal = returnVal & "&" & x & "=" & Request.QueryString(x)										
					end if
				end if
			next
			CreateSubPageURL = returnVal
		else
			CreateSubPageURL = defaultpage
		end if
	end function

	function CreateSONavigationLink(targetWindow, objectType, objectID, linkTitle, toolTip)
		dim queryId : queryId = GetSOPrimaryKeyID(objectType)
	
		if(isnull(objectID) or len(objectID) = 0) then
			CreateSONavigationLink = linkTitle
			exit function
		end if
		if(len(toolTip) > 0) then
			toolTip = "title='"  & replace(toolTip, "'", "") & "' "
		end if
		CreateSONavigationLink = "<a " & toolTip & " href='" & SUPEROFFICE_PROTOCOL_URL & targetWindow & "?" & queryId & "=" & objectID & "'>" & linkTitle & "</a>"
	end function

	function CreateSOPanelLink(targetWindow, objectType, linkTitle, toolTip)
	
		if(len(toolTip) > 0) then
			toolTip = "title='"  & replace(toolTip, "'", "") & "' "
		end if

		CreateSOPanelLink = "<a " & toolTip & " href='" & SUPEROFFICE_PROTOCOL_URL & targetWindow & "'>" & linkTitle & "</a>"
	end function

	function CreateSOLink(targetWindow, subPanel, url, linkTitle, toolTip)
		CreateSOLink = CreateSOStartLink(targetWindow, subPanel, url, linkTitle, toolTip) & "</a>"
	end function

	function CreateSOStartLink(targetWindow, subPanel, url, linkTitle, toolTip)
		dim targetRootPage : targetRootPage = GetRootPage(targetWindow) & "?"
		dim targetPage : targetPage = ""
		dim queryString : queryString = ""
		if(inStr(1, url, "?") > 0) then
			targetPage = mid(url, 1, inStr(1, url, "?") - 1)
			queryString = mid(url, inStr(1, url, "?") + 1)
		else
			Response.Write "No redirect page found!"
			Response.End
		end if
		if(len(subPanel) > 0) then
			subPanel = "." & subPanel
		end if
		if(len(toolTip) > 0) then
			toolTip = "title='"  & toolTip & "' "
		end if
		if (targetWindow = SUPEROFFICE_XIS_REPORTS_URL) then 'no frameset for this function - url goes directly to target
			targetRootPage = targetPage & "?"
			targetPage = ""
		end if
		CreateSOStartLink = "<a " & toolTip & "target='" & targetWindow & subPanel & "' href='" & targetRootPage & queryString & "&subPage=" & targetPage & "'>" & linkTitle
	end function

	function GetRootPage(targetWindow)
		select case targetWindow
			case SUPEROFFICE_XIS_MYPAGE_URL
				GetRootPage = "/xtra/minStartSideSubSO.asp"
			case SUPEROFFICE_XIS_TASK_URL
				GetRootPage = "/xtra/OppdragSubSO.asp"
			case SUPEROFFICE_XIS_CONSULTANT_URL
				GetRootPage = "/xtra/VikarSubSO.asp"
			case SUPEROFFICE_XIS_ECONOMICS_URL
				GetRootPage = "/xtra/økonomiSubSO.asp"
			case SUPEROFFICE_XIS_REPORTS_URL
				GetRootPage = "/xtra/rapporter/rapporter.asp"				
		end select
	end function
	
	function GetSOPrimaryKeyID(targetWindow)
		select case targetWindow
			case SUPEROFFICE_PANEL_CONTACT_URL
				GetSOPrimaryKeyID = "contact_id"
			case SUPEROFFICE_PANEL_PERSON_URL
				GetSOPrimaryKeyID = "person_id" 'add more as needed
		end select
	end function	
%>