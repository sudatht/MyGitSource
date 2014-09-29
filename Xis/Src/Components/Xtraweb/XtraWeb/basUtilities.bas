Attribute VB_Name = "basUtilities"
Option Explicit

Public Enum EF_LogEventType
    EF_LogEventType_Warning = vbLogEventTypeWarning
    EF_LogEventType_Error = vbLogEventTypeError
    EF_LogEventType_Info = vbLogEventTypeInformation
End Enum

Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" _
              (ByVal lpOutputString As String)


' -----------------------------------------------------------------------------
' Description       : Saves a record to Cv_Upload table
' Author            : ska
' -----------------------------------------------------------------------------
Public Sub Trace(ByVal strMessage As String)
    'In IDE
   Debug.Print strMessage
   'In compiled exe
   OutputDebugString strMessage
End Sub


Public Function ExistsInCol(ByVal colDataValues As Collection, ByVal strCheckFor As Variant) As Boolean
    'Checks if an object exists in a collection
    Dim objValue As DataValue
    
    For Each objValue In colDataValues
        If LCase(objValue.ValueName) = LCase(strCheckFor) Then
            ExistsInCol = True
            Exit Function
        End If
    Next
    ExistsInCol = False
End Function

Public Function ExistsInDataValues(ByVal colDataValues As DataValues, ByVal strCheckFor As Variant) As Boolean
    'checks if a object exist in a DataValue collection
    Dim objValue As DataValue
    For Each objValue In colDataValues
        If LCase(objValue.ValueName) = LCase(strCheckFor) Then
            ExistsInDataValues = True
            Exit Function
        End If
    Next
    ExistsInDataValues = False
    
End Function

Public Function WriteLog(ByVal strMember As String, ByVal strError As String, ByVal lLogType As EF_LogEventType)
    'writes to the log
    App.LogEvent strMember & ": " & strError, lLogType
End Function

Public Function CheckValueInCol(ByVal col As Collection, ByVal strCheckFor As Variant) As Boolean

On Error GoTo err_CheckValueInCol
    
    Dim strTest As String
    
    strTest = col(strCheckFor)
    CheckValueInCol = True
    Exit Function
err_CheckValueInCol:
    
    CheckValueInCol = False
End Function


Public Function CheckValueInArray(ByVal arArray As Variant, ByVal vValue As Variant) As Boolean
    ' checks if a value exists in an array
    Dim i As Integer
    i = 1
    Do Until IsEmpty(arArray(i))
        If arArray(i) = vValue Then
            CheckValueInArray = True
            Exit Function
        End If
        i = i + 1
    Loop
    CheckValueInArray = False
End Function

Public Function BuildErrorMsg(objErr As ADODB.Error) As String
    Dim strError As String
    
    strError = "Description: " & objErr.Description & vbCrLf
    strError = strError & "NativeError: " & objErr.NativeError & vbCrLf
    strError = strError & "Number: " & objErr.Number & vbCrLf
    strError = strError & "Source: " & objErr.Source & vbCrLf
    strError = strError & "SQLState: " & objErr.SQLState
    BuildErrorMsg = strError
    
End Function

Public Sub WriteADOErrorsToLog(objErrors As ADODB.Errors)
    Dim objErr As ADODB.Error
    For Each objErr In objErrors
        WriteLog "Address_Save", BuildErrorMsg(objErr), EF_LogEventType_Error
    Next

End Sub

Public Function ReturnVal(ByRef ObjField As Field) As Variant
        If IsNull(ObjField.Value) Then
            ReturnVal = adEmpty
        Else
            ReturnVal = ObjField.Value
        End If
End Function

Public Function ReturnString(ByRef ObjField As Field) As Variant
        If (IsNull(ObjField.Value) Or Len(Trim(ObjField.Value)) = 0) Then
            ReturnString = " "
        Else
            ReturnString = ObjField.Value
        End If
End Function

Public Function leftpad(StrIn As String, LNoChar As Long, ChrPad As String)
    If Len(StrIn) < LNoChar Then
        While Len(StrIn) < LNoChar
            StrIn = ChrPad & StrIn
        Wend
        leftpad = StrIn
    Else
        leftpad = StrIn
    End If
End Function

