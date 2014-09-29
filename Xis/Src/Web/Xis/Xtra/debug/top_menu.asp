<%
dim strAMenyElements(4,4)	'Menu category, Filename, menu element name, file exists: 0 = no, 1 = yes
dim intNoElements			'Total number of elements in menu
dim intCurrentElement		'The currently selected element
dim strSelectedMenuCategory	'Selected category element
dim objFSO					'File system object, used to check if a menu file exists
dim strScriptPath			'Path to this script file
dim strThisScript			'Name of the page this file is includet in
dim strLoopScript			'Name of current file element and path
dim strPrevElementCategory	'The previous element category

const C_CATEGORY = 1
const C_PAGENAME = 2
const C_PAGETITLE = 3
const C_PAGEEXISTS = 4

'Initialize menu elements
strAMenyElements(1, C_CATEGORY) = "Database"
strAMenyElements(1, C_PAGENAME) = "dbConnection.asp"
strAMenyElements(1, C_PAGETITLE) = "DB Connector"
strAMenyElements(1, C_PAGEEXISTS) = false
strAMenyElements(2, C_CATEGORY) = "Client"
strAMenyElements(2, C_PAGENAME) = "cookies.asp"
strAMenyElements(2, C_PAGETITLE) = "Cookies"
strAMenyElements(2, C_PAGEEXISTS) = false
strAMenyElements(3, C_CATEGORY) = "Client"
strAMenyElements(3, C_PAGENAME) = "browserVariables.asp"
strAMenyElements(3, C_PAGETITLE) = "Browser variables"
strAMenyElements(3, C_PAGEEXISTS) = false
strAMenyElements(4, C_CATEGORY) = "Server"
strAMenyElements(4, C_PAGENAME) = "default.asp"
strAMenyElements(4, C_PAGETITLE) = "Application/session variables"
strAMenyElements(4, C_PAGEEXISTS) = false

'Determine number of elements
intNoElements = ubound(strAMenyElements)

if intNoElements > 0 then
	'Create object to verify the individual script elements..
	set objFSO = server.CreateObject("Scripting.FileSystemObject")
	
	'get the filename of the file in which this ASP-script is included..
	strThisScript = Request.ServerVariables("SCRIPT_NAME")
	strThisScript = lcase(trim(mid(strThisScript, instrrev(strThisScript, "/") + 1)))

	strScriptPath = Request.ServerVariables("PATH_TRANSLATED")
	strScriptPath = trim(mid(strScriptPath, 1, instrrev(strScriptPath, "\")))

	'Run through all menu-categories and render them if files exists.. 
	intCurrentElement = 1
	while (intCurrentElement <= intNoElements)
		'Create concatinated file path..
		strLoopScript = strScriptPath & strAMenyElements(intCurrentElement, C_PAGENAME)
		if (objFSO.FileExists(strLoopScript) = true) then
				'Mark file as exists!
				strAMenyElements(intCurrentElement, C_PAGEEXISTS) = true
				'Store menu category for later use..
				if ( lcase(strAMenyElements(intCurrentElement, C_PAGENAME)) = strThisScript ) then
					strSelectedMenuCategory = strAMenyElements(intCurrentElement, C_CATEGORY)
				end if				

				if ( strPrevElementCategory <> strAMenyElements(intCurrentElement, C_CATEGORY) ) then
				'if this is a new meny category, we haven't rendered..	
					if ( strAMenyElements(intCurrentElement, C_PAGENAME) = strThisScript ) then
					'If this is the selected menu category, render it without a link..
						%>
						<span class="MenuCategorySelected"><a Style="font-name:arial;font-size:12px;color:lightblue;" href="#"><%=strAMenyElements(intCurrentElement, C_CATEGORY)%></a></span>
						<%
					else
					'else, the selected menu category is not selected, so render it with a link..
						%>
						<span class="MenuCategory"><a Style="font-name:arial;font-size:12px;color:blue;" href="<%=strAMenyElements(intCurrentElement, C_PAGENAME)%>"><%=strAMenyElements(intCurrentElement, C_CATEGORY)%></a></span>
						<%				
					end if
				end if
				strPrevElementCategory = strAMenyElements(intCurrentElement, C_CATEGORY)
		end if
		intCurrentElement = intCurrentElement + 1		
	wend
	%>
	<br>
	<%
	intCurrentElement = 1
	while (intCurrentElement <= intNoElements)
		if (cstr(strAMenyElements(intCurrentElement, C_CATEGORY)) = strSelectedMenuCategory and strAMenyElements(intCurrentElement, 4) = true) then
		'If this is a menu element / file that exists and this menu element / file exists under the current menu category..
			if ( lcase(strAMenyElements(intCurrentElement, C_PAGENAME)) = strThisScript ) then
			'If This menu element / file is active,  don't render it with a link..
				%>
				<a Style="border-left:15px;border-right:15px;font-name:arial;font-size:12px;color:lightblue;" href="#"><%=strAMenyElements(intCurrentElement, C_PAGETITLE)%></a>
				<%
			else
			'else, this menu element / file is not selected, so render it with a link..
				%>
				<a Style="border-left:15px;border-right:15px;font-name:arial;font-size:12px;color:blue;" href="<%=strAMenyElements(intCurrentElement, C_PAGENAME)%>"><%=strAMenyElements(intCurrentElement, C_PAGETITLE)%></a>
				<%				
			end if
		end if
		intCurrentElement = intCurrentElement + 1		
	wend
end if
%>