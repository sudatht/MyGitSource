VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "Suspect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mobjCV As CV
Private mobjEmployee As Employee
Private mobjOwner As Object
Private mobjDataValues As DataValues
Private mobjAddresses As Addresses
Private mobjJobApplication As JobApplication

Private mStrXtraCon As String
Private mStrXtraDataShapeCon As String
Private mStrImpCon As String

Public Property Get XtraConString() As String
    XtraConString = mStrXtraCon
End Property

Public Property Let XtraConString(strConString As String)
    mStrXtraCon = strConString
End Property

Public Property Get XtraDataShapeConString() As String
    XtraDataShapeConString = mStrXtraDataShapeCon
End Property

Public Property Let XtraDataShapeConString(strConString As String)
    mStrXtraDataShapeCon = strConString
End Property

Public Property Set JobApplication(ByVal vData As JobApplication)
    Set mobjJobApplication = vData
End Property

Public Property Get JobApplication() As JobApplication
    If Not mobjJobApplication Is Nothing Then
        mobjJobApplication.XtraConString = mStrXtraCon
    End If
    Set JobApplication = mobjJobApplication
End Property

Public Property Set Addresses(ByVal vData As Addresses)
    '##BD Dette er en kolleksjon av %Address% objekter.
    Set mobjAddresses = vData
End Property

Public Property Get Addresses() As Addresses
    Set Addresses = mobjAddresses
    mobjAddresses.Refresh
End Property

Public Property Set DataValues(ByVal vData As DataValues)
    '##BD Dette er en kolleksjon av %DataValue% objekter.
    Set mobjDataValues = vData
End Property

Public Property Get DataValues() As DataValues
Attribute DataValues.VB_UserMemId = 0
    Set DataValues = mobjDataValues
End Property

Public Property Get Employee() As Employee
    '##BD Referanse til ansatt objekt hos Xtra. Dvs. kontaktperson hos Xtra.
    If mobjEmployee Is Nothing Then
        Set mobjEmployee = objContext.CreateInstance("xtraweb.Employee")
    End If
    Set Employee = mobjEmployee
End Property

Public Property Set Employee(vData As Employee)
    Set mobjEmployee = vData
End Property

Public Property Get CV() As CV
    '##BD Denne egenskapen er en referanse til et %CV% objekt for en suspect.
    If mobjCV Is Nothing Then
        Set mobjCV = objContext.CreateInstance("xtraweb.CV")
    End If
    If mobjCV.XtraConString <> "" Then
        mobjCV.Refresh
    End If
    Set CV = mobjCV
End Property

Public Property Set CV(vData As CV)
    Set mobjCV = vData
End Property

Public Function Save() As Boolean
On Error GoTo err_Save
    '##BD Lagrer objektet. Lagrer både endringer og nye objekter.
    Dim fldData As ADODB.Field
    Dim ObjDAL As IdataAccess
    Dim objValue As XtraWeb.DataValue
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim lSuspectID As Long

    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No connection string (Xtra)!"
    End If
    
    Set ObjDAL = New ClsDataAccess

    Set rs = New ADODB.Recordset

    'check if new
    If ExistsInDataValues(mobjDataValues, "suspectid") Then
        'exists
        lSuspectID = mobjDataValues("suspectid").Value
    End If

    strSQL = "Select * from v_suspect where suspectid = " & lSuspectID
    Set rs = ObjDAL.FetchRC(mStrXtraCon, strSQL)
    If rs.RecordCount = 0 Then
        'new
        rs.AddNew
    End If
    For Each objValue In mobjDataValues
        If Not IsNull(objValue.Value) Then
            If LCase(objValue.ValueName) <> "suspectid" Then
                rs(objValue.ValueName) = objValue.Value
            End If
        End If
    Next
    
    rs("slettes") = 0
    rs("regdato") = Date
    ObjDAL.BatchUpdate mStrXtraCon, rs

    ' Refresh all the values from the database
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
    For Each fldData In rs.Fields
        mobjDataValues(fldData.Name) = fldData.Value
    Next
    rs.Close
    Set rs = Nothing
    Set ObjDAL = Nothing
    Save = True
    Exit Function
    
err_Save:
    WriteLog "Suspect_Save", Error$, EF_LogEventType_Error
    Save = False

End Function

Public Function GetSuspect(ByVal lSuspectID As Long) As Boolean
Trace "Enter function Suspect.GetSuspect "
Trace "Parameter suspectid :" & lSuspectID
On Error GoTo err_GetSuspect
    Dim rsSuspect As ADODB.Recordset
    Dim ObjDAL As IdataAccess
    Dim strSQL As String
    Dim fldData As ADODB.Field
    
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No connection string (Xtra)!"
    End If

    Set ObjDAL = New ClsDataAccess
    Set rsSuspect = New ADODB.Recordset

    strSQL = "Select * from v_suspect where suspectid = " & lSuspectID
    Set rsSuspect = ObjDAL.FetchRC(mStrXtraCon, strSQL)
    If Not rsSuspect.EOF Then
        For Each fldData In rsSuspect.Fields
            mobjDataValues(fldData.Name) = fldData.Value
        Next
        GetSuspect = True
    Else
        GetSuspect = False
    End If
    Set rsSuspect = Nothing
    Set ObjDAL = Nothing
    Trace "Exit function Suspect.GetSuspect "
    Exit Function
    
err_GetSuspect:
    Trace "Error function Suspect.GetSuspect "
    WriteLog "Suspect_GetSuspect", Error$, EF_LogEventType_Error
End Function

Private Sub ObjectControl_Activate()
    On Error GoTo err_Class_Initialize
        Set objContext = GetObjectContext()
        Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
        Set mobjCV = objContext.CreateInstance("xtraweb.CV")
        Set mobjCV.Owner = SafeRef(Me)
        Set mobjAddresses = objContext.CreateInstance("xtraweb.Addresses")
        Set mobjAddresses.Owner = SafeRef(Me)
        Set mobjJobApplication = objContext.CreateInstance("xtraweb.JobApplication")
        Set mobjJobApplication.Owner = SafeRef(Me)
        Exit Sub
        
err_Class_Initialize:
        WriteLog "Suspect_Class_Initialize", Error$, EF_LogEventType_Error
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set mobjJobApplication.Owner = Nothing
    Set mobjJobApplication = Nothing
    Set mobjDataValues = Nothing
    Set mobjCV.Owner = Nothing
    Set mobjCV = Nothing
    Set mobjAddresses.Owner = Nothing
    Set mobjAddresses = Nothing
    Set mobjEmployee = Nothing
    Set mobjOwner = Nothing
    Set objContext = Nothing
End Sub

Public Sub CleanUp()
    Set mobjJobApplication.Owner = Nothing
    Set mobjJobApplication = Nothing
    Set mobjDataValues = Nothing
    Set mobjCV.Owner = Nothing
    Set mobjCV = Nothing
    Set mobjAddresses.Owner = Nothing
    Set mobjAddresses = Nothing
    Set mobjEmployee = Nothing
    Set mobjOwner = Nothing
    Set objContext = Nothing
End Sub

