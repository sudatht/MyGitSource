<%

public function GeneratePassword(numberOfChars)
	'Array of legal letters for password
	dim legalLetters
	'Array of legal digits for password
	dim legalDigits
	dim returnValue
	dim letterLength, digitsLength
	dim currentPos
	dim charChance
	dim roll
	dim nofDigits, nofLetters
	dim exitLoop

	'Initialize
	charChance = 70 '% Chance of being a character
	currentPos = 1
	legalLetters = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
	legalDigits = Array("2", "3", "4", "5", "6", "7", "8", "9")
	letterLength = UBound(legalLetters)
	digitsLength = UBound(legalDigits)

	call randomize()
	while (currentPos <= numberOfChars)
		'Use letter or digit?
		roll  = Int((100 - 1) * Rnd())
		if (roll <= charChance) then
		'add letter
			index = Int((letterLength - 1) * Rnd())
			nofLetters = nofLetters + 1
			returnValue = returnValue + legalLetters(index)
		else
		'add digit
			index = Int((digitsLength - 1) * Rnd())
			nofDigits = nofDigits + 1
			returnValue = returnValue + legalDigits(index)
		end if
		currentPos = currentPos + 1
	wend

	GeneratePassword = returnValue
end function


%>