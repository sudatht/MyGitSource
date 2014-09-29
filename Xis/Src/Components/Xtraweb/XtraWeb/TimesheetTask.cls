VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "TimesheetTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mobjDataValues As DataValues
Private mObjWeek As TimeSheetWeek
Private mobjOwner As Object
Private mobjTasks As TimesheetTasks

Public Property Set Owner(ByVal vDataObj As Object)
    
    On Error GoTo Err_handler
    If Not vDataObj Is Nothing Then
        If ((Not TypeOf vDataObj Is Consultant) And (Not TypeOf vDataObj Is Customer)) Then
            Err.Raise vbObjectError + 751, "TimesheetTask_owner", "TimesheetTask owner is of wrong type."
        End If
        Set mobjOwner = vDataObj
    End If
    Exit Property

Err_handler:
    Set mobjOwner = Nothing
    WriteLog "timesheetTask_Owner", Error$, EF_LogEventType_Error
End Property

Public Property Get Owner() As Object
    Set Owner = mobjOwner
End Property

Public Property Set Tasks(ByVal vDataObj As TimesheetTasks)
    
    On Error GoTo Err_handler
    
    Set mobjTasks = vDataObj
    Exit Property
Err_handler:
    
    Set mobjTasks = Nothing
    WriteLog "timesheetTask_Tasks", Error$, EF_LogEventType_Error
End Property

Public Property Get Tasks() As TimesheetTasks
    Set Tasks = mobjTasks
End Property

Public Property Get TimeSheetWeek() As TimeSheetWeek
    Set TimeSheetWeek = mObjWeek
End Property

Public Property Set DataValues(ByVal vData As DataValues)
    Set mobjDataValues = vData
End Property

Public Property Get DataValues() As DataValues
    Set DataValues = mobjDataValues
End Property

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
    Set mobjTasks = objContext.CreateInstance("xtraweb.TimesheetTasks")
End Sub

Private Sub ObjectControl_Deactivate()
    Set mobjDataValues = Nothing
    Set mObjWeek = Nothing
    Set mobjOwner = Nothing
    Set mobjTasks = Nothing
    Set objContext = Nothing
End Sub

Public Sub CleanUp()
    Set mobjDataValues = Nothing
    Set mObjWeek = Nothing
    Set mobjOwner = Nothing
    Set mobjTasks = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Public Sub GetTimeSheetWeek(Optional ByVal Lvikarid As Long)
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim lID As Long
    Dim strOwnerType As String
    Dim StrProcedure As String
    Dim strUkenr As String
    Dim LOppdragid As Long
    Dim rsTimeSheetForWeek As ADODB.Recordset
    Dim ObjDay As TimeSheetDay
    Dim ObjFld As Field
    Dim ObjTasks As TimesheetTasks
    Dim dtTilDato As Date
    Dim dtWeekStart As Date
    Dim StrConnection As String
    Dim rsCommentary As Recordset
    Dim LkontaktID As Long

On Error GoTo Err_handler

    If ((mobjOwner Is Nothing) Or (mobjOwner.DataValues.Count = 0)) Then
       Err.Raise 751 + vbObjectError, "timesheetTask_GetTimeSheetWeek", " Missing owner (Lvikarid/ifirmaid) !"
    End If
                   
    If mobjDataValues Is Nothing Then
        Err.Raise 753 + vbObjectError, "timesheetTask_GetTimeSheetWeek", " No tasks for this week (mobjdatavalues is empty) !"
    End If
    
    LOppdragid = mobjDataValues("oppdragid")
    dtTilDato = CDate(mobjTasks.DataValues("tildato"))
    dtWeekStart = CDate(mobjTasks.DataValues("WeekStart"))
    
    Set ObjCmd = New ADODB.Command
                     
    If TypeOf mobjOwner Is Consultant Then
        lID = mobjOwner.DataValues("vikarid").Value
        StrProcedure = "GetConsultantTimesheetForWeek"
        If lID < 1 Then
            Err.Raise 754 + vbObjectError, "timesheetTask_GetTimeSheetWeek", " Invalid consultant id !"
        End If
    ElseIf TypeOf mobjOwner Is Customer Then
        lID = Lvikarid
        StrProcedure = "GetCustomerTimesheetForWeek"
        If lID < 1 Then
            Err.Raise 755 + vbObjectError, "timesheetTask_GetTimeSheetWeek", " Invalid firma id !"
        End If
    End If
             
    'Add input parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "vikarid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = lID
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
       
    If TypeOf mobjOwner Is Customer Then
        Set ObjParam = New Parameter
        ObjParam.Name = "KontaktID"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = mobjOwner.DataValues("kontaktID").Value
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
    End If
        
    Set ObjParam = New Parameter
    ObjParam.Name = "Oppdragid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LOppdragid
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing

    Set ObjParam = New Parameter
    ObjParam.Name = "startdato"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adDate
    ObjParam.Value = dtWeekStart
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjParam = New Parameter
    ObjParam.Name = "tildato"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adDate
    ObjParam.Value = dtTilDato
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjDAL = New ClsDataAccess
    StrConnection = mobjTasks.DataValues("strconnection")
    If Len(Trim(StrConnection)) <= 5 Then
        Err.Raise vbObjectError + 756, "TimesheetTasks_GetTasksForWeek", "Missing or invalid strConnection parameter."
    End If
    
    Set rsTimeSheetForWeek = ObjDAL.FetchRCSP(StrConnection, StrProcedure, ObjCmd.Parameters)
    If Not IsObject(rsTimeSheetForWeek) Then
        Err.Raise 757 + vbObjectError, "timesheetTask_GetTimeSheetWeek", "Error in GetTimeSheetForWeek!"
    End If

    'Dersom timesheets..
    Set mObjWeek = objContext.CreateInstance("xtraweb.TimeSheetWeek")
    Set mObjWeek.Task = SafeRef(Me)
    Set mObjWeek.Owner = Me.Owner
    If Not rsTimeSheetForWeek.EOF Then
        While Not rsTimeSheetForWeek.EOF
            'Legg til taskene til Tasks col'en
            Set ObjDay = objContext.CreateInstance("xtraweb.TimeSheetDay")
            For Each ObjFld In rsTimeSheetForWeek.Fields
                ObjDay.DataValues.Add ObjFld.Name, ObjFld.Value
            Next
            mObjWeek.Add ObjDay
            rsTimeSheetForWeek.MoveNext
        Wend
        
        If TypeOf mobjOwner Is Consultant Then
             'get Commentary for week.
             Set ObjCmd = New ADODB.Command
                        
            'Add input parameters to command object
            Set ObjParam = New Parameter
            ObjParam.Name = "Lvikarid"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = lID
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
                      
            Set ObjParam = New Parameter
            ObjParam.Name = "LOppdragid"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = LOppdragid
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
             
            strUkenr = DatePart("yyyy", dtWeekStart, vbMonday, vbFirstFourDays) & DatePart("ww", dtWeekStart, vbMonday, vbFirstFourDays)
            Set ObjParam = New Parameter
            ObjParam.Name = "intUkenr"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = CLng(strUkenr)
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
                               
            Set rsCommentary = ObjDAL.FetchRCSP(StrConnection, "GetCommentary", ObjCmd.Parameters)
            Set ObjCmd = Nothing
            If rsCommentary.EOF Then
                mObjWeek.DataValues.Add "Kommentar", ""
            Else
                mObjWeek.DataValues.Add "Kommentar", rsCommentary.Fields("Kommentar").Value
            End If
            Set rsCommentary = Nothing
        End If
    'Dersom ingen tasks..
    Else
        Set mObjWeek = Nothing
    End If
    Set rsTimeSheetForWeek = Nothing
    Set ObjDAL = Nothing
    Exit Sub

Err_handler:
    Set mObjWeek = Nothing
    WriteLog "timesheetTask_GetTimeSheetWeek", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Sub
