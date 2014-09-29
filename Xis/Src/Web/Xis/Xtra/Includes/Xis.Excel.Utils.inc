<%
function ToCSVString(inputString, preserveWhiteSpace)
	dim localTmp : localTmp = inputString

	if (preserveWhiteSpace = false) then
		localTmp = trim(localTmp)
	end if

	if (InStr(1, localTmp, """")) then
		localTmp = replace(localTmp, """", """""")
		localTmp = """" & localTmp & """"
		ToCSVString = localTmp
		exit function
	end if

	if (preserveWhiteSpace = true) then
		if(left(localTmp, 1) = vbTab or left(localTmp, 1) = " " or Right(localTmp, 1) = vbTab or Right(localTmp, 1) = " ") then
			localTmp = """" & localTmp & """"
			ToCSVString = localTmp
			exit function
		end if
	end if

	'if inputstring contains a comma, enclose it in quotes
	if (InStr(1, localTmp, ",")) then
		localTmp = """" & localTmp & """"
		ToCSVString = localTmp
		exit function
	end if

	'if inputstring contains enter character (line feed), quotify
	if (InStr(1, localTmp, vbLf)) then
		localTmp = """" & localTmp & """"
		ToCSVString = localTmp
		exit function
	end if

	'if inputstring contains enter characters (carriage return + line feed), quotify
	if (InStr(1, localTmp, vbCrLf)) then
		localTmp = """" & localTmp & """"
		ToCSVString = localTmp
		exit function
	end if

	ToCSVString = localTmp
end function

function recordSetRowToCSVString(sourceSet, delimiter, preserveWhiteSpace)
	dim returValue : returValue = ""
	dim fieldCount : fieldCount = sourceSet.Fields.Count

	for teller = 1 to fieldCount
		if (teller = fieldCount) then
			returValue = returValue & ToCSVString(rsTest.Fields(teller-1).Value, preserveWhiteSpace)
		else
			returValue = returValue & ToCSVString(rsTest.Fields(teller-1).Value, preserveWhiteSpace) & delimiter
		end if
	next
	returValue = returValue & vbCrLf
	recordSetRowToCSVString = returValue
end function
%>