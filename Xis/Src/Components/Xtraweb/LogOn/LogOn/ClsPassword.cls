VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GeneratePassword(ByVal intLength As Integer) As String
    Dim intCurrentLetter As Integer
    Dim intPercntChance As Integer
    Dim strLetter As String
    Dim intCharbase As Integer
    Dim intDigitbase As Integer
    Dim intNofLetters As Integer
    Dim intNofDigits As Integer
    
    intCharbase = 64
    intDigitbase = 48
    intCurrentLetter = 0
    If ((intLength < 6) Or (intLength > 10)) Then
        GeneratePassword = ""
        Exit Function
    End If
    'Password must be intLength letters long
    While (intCurrentLetter < intLength)
        'Make sure there are at least 3 digit in the password
        If (((intLength - 3) = intCurrentLetter) And (intNofDigits < 3)) Then
            While (intNofDigits < 3)
                Call Randomize
                intPercntChance = Int((10 * Rnd))
                strLetter = Chr(intDigitbase + intPercntChance)
                intNofDigits = intNofDigits + 1
                intCurrentLetter = intCurrentLetter + 1
                GeneratePassword = GeneratePassword + strLetter
            Wend
        'If there is 3 digits, the rest should be letters
        ElseIf (intNofDigits = 3) Then
            While (intCurrentLetter < intLength)
                Randomize
                intPercntChance = Int((26 * Rnd) + 1)
                strLetter = Chr(intCharbase + intPercntChance)
                intNofLetters = intNofLetters + 1
                intCurrentLetter = intCurrentLetter + 1
                GeneratePassword = GeneratePassword + strLetter
            Wend
        Else
            Randomize
            intPercntChance = Int((100 * Rnd) + 1)
            '65% chance Letter
            If (intPercntChance <= 65) Then
                intPercntChance = Int((26 * Rnd) + 1)
                strLetter = Chr(intCharbase + intPercntChance)
                intNofLetters = intNofLetters + 1
            '35% chance digit
            ElseIf (intPercntChance > 35) Then
                intPercntChance = Int((10 * Rnd))
                strLetter = Chr(intDigitbase + intPercntChance)
                intNofDigits = intNofDigits + 1
            End If
            GeneratePassword = GeneratePassword + strLetter
            intCurrentLetter = intCurrentLetter + 1
        End If
    Wend
End Function

'Checks if the given password is at least 6 characters long, not more than 10 characters,
'consists of at least 3 digits, and that the letters are in [a-z] (case insensitive).
Public Function IsValidPassword(ByVal strPassword As String) As Boolean
    Dim intLength As Integer
    Dim strLetter As String
    Dim intNofDigits As Integer
          
    IsValidPassword = True
    
    intLength = Len(strPassword)
    
    'Password must be intLength letters long
    If ((intLength < 6) Or (intLength > 10)) Then
        IsValidPassword = False
        Exit Function
    End If
 
    'password is not case sensitive
    strPassword = UCase(strPassword)

    'Check for invalid chars
    While (strPassword <> "")
        strLetter = Mid(strPassword, 1, 1)
        If ((Asc(strLetter) >= 48) And (Asc(strLetter) <= 57)) Then
            intNofDigits = intNofDigits + 1
        ElseIf ((Asc(strLetter) >= 65) And (Asc(strLetter) <= 90)) Then
            'Do nuffing
        Else
            IsValidPassword = False
            Exit Function
        End If
        strPassword = Mid(strPassword, 2)
    Wend
    If intNofDigits < 3 Then
        IsValidPassword = False
    End If
End Function
