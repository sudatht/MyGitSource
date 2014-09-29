<%
'Declare all rights 
const RIGHT_NONE = 0 
const RIGHT_READ = 1
const RIGHT_WRITE = 2
const RIGHT_ADMIN = 3
const RIGHT_SUPER = 4

'Declare all access types
const ACCESS_HOTLIST = 1
const ACCESS_CUSTOMER = 2
const ACCESS_CONSULTANT = 3
const ACCESS_REPORT = 4
const ACCESS_TASK = 5
const ACCESS_ADMIN = 6


'Get user profile
dim UserProfile : UserProfile = Session("Profil")

'Helperfunction for determing rights

private function HasUserRight(access, right)
	dim userAccessRight
	HasUserRight = false 'init to no rights
	if(len(UserProfile) >= access) then 
		userAccessRight = cint(mid(UserProfile, access, 1))
		if(userAccessRight >= right) then
			HasUserRight = true			
		end if
	end if
end function


private function WriteProfile()
	dim profileLen : profileLen = len(UserProfile)
	dim i : i = 1
	if (profileLen > 0) then
		call response.write("<table border='1'>")
		call response.write("	<tr>")
		call response.write("		<th>")
		call response.write("			HOTLIST")
		call response.write("		</th>")
		call response.write("		<th>")
		call response.write("			CUSTOMER")		
		call response.write("		</th>")
		call response.write("		<th>")
		call response.write("			CONSULTANT")
		call response.write("		</th>")
		call response.write("		<th>")
		call response.write("			REPORT")
		call response.write("		</th>")
		call response.write("		<th>")
		call response.write("			TASK")		
		call response.write("		</th>")
		call response.write("		<th>")
		call response.write("			ADMIN")
		call response.write("		</th>")
		call response.write("	</tr>")
		call response.write("	<tr>")		
		for i = 1 to profileLen
			call response.write("		<td>")
			call response.write(GetRightType(mid(UserProfile, i, 1)))
			call response.write("		</td>")
		next
		call response.write("	</tr>")
		call response.write("</table>")
	end if
end function

private function GetRightType(constIndex)
	GetRightType = ""
	
	select case constIndex
	case "0" 
		GetRightType = "NONE"
	case "1"
		GetRightType = 	"READ"
	case "2"
		GetRightType = 	"WRITE"
	case "3"
		GetRightType = 	"ADMIN"
	case "4"
		GetRightType = 	"SUPER"
	end select
end function
%>