VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "TimesheetWeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mCol As Collection 'local variable to hold collection
Private mobjOwner As Object
Private mObjTask As TimesheetTask
Private mobjDataValues As DataValues

Public Property Set DataValues(ByVal vData As DataValues)
    Set mobjDataValues = vData
End Property

Public Property Get DataValues() As DataValues
    Set DataValues = mobjDataValues
End Property

Public Property Set Owner(ByVal vDataObj As Object)
    On Error GoTo Err_handler
    
    If Not vDataObj Is Nothing Then
        If ((Not TypeOf vDataObj Is Consultant) And (Not TypeOf vDataObj Is Customer)) Then
            Err.Raise vbObjectError + 751, "TimesheetWeek_owner", "TimesheetWeek owner is of wrong type."
        End If
        Set mobjOwner = vDataObj
    End If
    Exit Property

Err_handler:
    Set mobjOwner = Nothing
    WriteLog "timesheetWeek_Owner", Error$, EF_LogEventType_Error
End Property

Public Property Get Owner() As Object
    Set Owner = mobjOwner
End Property

Public Property Set Task(ByVal vData As TimesheetTask)
    Set mObjTask = vData
End Property

Public Property Get Task() As TimesheetTask
    Set Task = mObjTask
End Property

Public Function Add(ByVal ObjDay As TimeSheetDay, Optional strKey As String) As TimeSheetDay
    mCol.Add ObjDay
    'return the object created
    Set Add = ObjDay
    Set ObjDay = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As TimeSheetDay
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Function Save(ByVal StrConnection As String) As Boolean
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim ObjDay As TimeSheetDay
    Dim dtDay As Date
    Dim Lvikarid As Long
    Dim LOppdragid As Long
    Dim strUkenr As String
    Dim strKommentar As String
    
    If ((mobjOwner Is Nothing) Or (mobjOwner.DataValues.Count = 0)) Then
       Err.Raise 751 + vbObjectError, "timesheetweek_save", " Missing owner (Lvikarid/ifirmaid) !"
    End If
                   
    If Len(StrConnection) <= 5 Then
       Err.Raise 752 + vbObjectError, "timesheetweek_save", " Invalid connection string!"
    End If
    
    'Check children objects (days)
    If mCol.Count > 0 Then
        'save each timesheetday object
        Set ObjDAL = New ClsDataAccess
        
        'Store first date of this week for future use
        dtDay = CDate(mCol(1).DataValues("dato").Value)
        'Store oppdragid for commentary save
        LOppdragid = CLng(mCol(1).DataValues("oppdragid").Value)

        For Each ObjDay In mCol
            Set ObjCmd = New ADODB.Command
            
            Lvikarid = mobjOwner.DataValues("vikarid").Value
               
            'Transfer values from day to parameters
            
            Set ObjParam = New Parameter
            ObjParam.Name = "Lvikarid"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = Lvikarid
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
            
            Set ObjParam = New Parameter
            ObjParam.Name = "LTimelistevikarid"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = ObjDay.DataValues("TimelisteVikarID").Value
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
            
            Set ObjParam = New Parameter
            ObjParam.Name = "dtstarttid"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adDate
            ObjParam.Value = CDate(ObjDay.DataValues("starttid").Value)
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
                       
            Set ObjParam = New Parameter
            ObjParam.Name = "dtsluttid"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adDate
            ObjParam.Value = ObjDay.DataValues("sluttid").Value
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing

            Set ObjParam = New Parameter
            ObjParam.Name = "dtlunsj"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adDate
            ObjParam.Value = ObjDay.DataValues("lunsj").Value
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing

            Set ObjParam = New Parameter
            ObjParam.Name = "dbltimerpaajobb"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adDouble
            ObjParam.Value = ObjDay.DataValues("timerpaajobb").Value
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing

            Set ObjParam = New Parameter
            ObjParam.Name = "strwebEndret"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adTinyInt
            ObjParam.Value = ObjDay.DataValues("web_endret").Value
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
        
            If Not ObjDAL.ExecuteCRUDSP(StrConnection, "UpdateConsultantTasks", ObjCmd.Parameters) Then
               Err.Raise 753 + vbObjectError, "timeSheetWeek_save", "Could not Execute stored procedure UpdateConsultantTasks! All tasks aborted!"
            End If
        Next ObjDay
    
        'Save commentary
        strKommentar = Trim(mobjDataValues("Kommentar"))
        
        Set ObjCmd = New ADODB.Command
        
        'Add input parameters to command object:
        'A combination of vikarid, oppdragid and year + weekno identifies the commentary.
        Set ObjParam = New Parameter
        ObjParam.Name = "Lvikarid"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = Lvikarid
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
                    
        Set ObjParam = New Parameter
        ObjParam.Name = "LOppdragid"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = LOppdragid
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
         
        If (DatePart("ww", dtDay, vbMonday, vbFirstFourDays) < 10) Then
            strUkenr = DatePart("yyyy", dtDay, vbMonday, vbFirstFourDays) & "0" & DatePart("ww", dtDay, vbMonday, vbFirstFourDays)
        Else
            strUkenr = DatePart("yyyy", dtDay, vbMonday, vbFirstFourDays) & DatePart("ww", dtDay, vbMonday, vbFirstFourDays)
        End If
        
        Set ObjParam = New Parameter
        ObjParam.Name = "intUkenr"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = CLng(strUkenr)
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
                                       
        Set ObjParam = New Parameter
        ObjParam.Name = "Kommentar"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adVarChar
        ObjParam.Value = strKommentar
        ObjParam.Size = Len(strKommentar)
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing

        If Not ObjDAL.ExecuteCRUDSP(StrConnection, "UpdateCommentary", ObjCmd.Parameters) Then
            Err.Raise 754 + vbObjectError, "timeSheetWeek_save", "Could not Execute stored procedure UpdateCommentary!"
        End If
        
        Set ObjDAL = Nothing
        Set ObjCmd = Nothing
        Save = True
    Else
        Save = False
    End If
    Exit Function

Err_handler:
    Save = False
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    WriteLog "timesheetWeek_save", Error$, EF_LogEventType_Error
End Function

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    'creates the collection when this class is created
    Set mCol = New Collection
    Set mObjTask = objContext.CreateInstance("xtraweb.TimesheetTask")
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
End Sub

Private Sub ObjectControl_Deactivate()
    'destroys objects when this class terminates
    Set mCol = Nothing
    Set mobjOwner = Nothing
    Set mObjTask = Nothing
    Set mobjDataValues = Nothing
End Sub

Public Sub CleanUp()
    'destroys objects when this class terminates
    Set mCol = Nothing
    Set mobjOwner = Nothing
    Set mObjTask = Nothing
    Set mobjDataValues = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Public Function Approve(ByVal StrConnection As String, ByVal dtWeekStart As Date, ByVal LOppdragid As Long, Optional Lvikarid As Long) As Boolean

    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim lID As Long
    Dim IFraMaaned As Integer
    Dim strOwnerType As String
    Dim StrProcedure As String
    Dim dtTilDato As Date

On Error GoTo Err_handler

    If ((mobjOwner Is Nothing) Or (mobjOwner.DataValues.Count = 0)) Then
       Err.Raise 751 + vbObjectError, "timesheetweek_Approve", " Missing owner (Lvikarid) !"
    End If
          
    If Not IsDate(dtWeekStart) Then
        Err.Raise 752 + vbObjectError, "timesheetweek_Approve", " Invalid fromdate!"
    End If
          
    If LOppdragid < 1 Then
       Err.Raise 753 + vbObjectError, "timesheetweek_Approve", " No Task has been choosen!"
    End If
    
    IFraMaaned = DatePart("m", dtWeekStart, vbMonday, vbFirstFourDays)
          
    If DatePart("w", dtWeekStart, vbMonday, vbFirstFourDays) <> 7 Then
            dtTilDato = DateAdd("d", 7 - DatePart("w", dtWeekStart, vbMonday, vbFirstFourDays), dtWeekStart)
            'Dersom ukeslutt befinner seg i annen måned, backtrack inntil slutten av fradato måned
            While DatePart("m", dtTilDato, vbMonday, vbFirstFourDays) <> IFraMaaned
                dtTilDato = DateAdd("d", -1, dtTilDato)
            Wend
    Else
        dtTilDato = dtWeekStart
    End If
                                      
    If TypeOf mobjOwner Is Consultant Then
        lID = mobjOwner.DataValues("vikarid").Value
        StrProcedure = "ApproveConsultantTasks"
        If lID < 1 Then
           Err.Raise 754 + vbObjectError, "timesheetweek_Approve", " Invalid consultant id !"
        End If
    ElseIf TypeOf mobjOwner Is Customer Then
        lID = Lvikarid
        StrProcedure = "ApproveCustomerTasks"
        If lID < 1 Then
           Err.Raise 754 + vbObjectError, "timesheetweek_Approve", " Invalid customer id !"
        End If
    End If
                                           
    Set ObjCmd = New ADODB.Command
    
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
    ObjParam.Name = "fradato"
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
    
    If Not ObjDAL.ExecuteCRUDSP(StrConnection, StrProcedure, ObjCmd.Parameters) Then
        Err.Raise 755 + vbObjectError, "timesheetweek_Approve", "Could not Execute stored procedure " & StrProcedure & "! All tasks aborted!"
    End If
    
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Approve = True
    Exit Function
    
Err_handler:
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Approve = False
    WriteLog "timesheetWeek_Approve", Error$, EF_LogEventType_Error
End Function

Public Function Reject(ByVal StrConnection As String, ByVal dtWeekStart As Date, ByVal LOppdragid As Long, Optional Lvikarid As Long) As Boolean

    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim strOwnerType As String
    Dim StrProcedure As String
    Dim dtTilDato As Date

On Error GoTo Err_handler

    If Lvikarid < 1 Then
       Err.Raise 750 + vbObjectError, "timesheetweek_reject", " Invalid consultant id !"
    End If
        
    If Not IsDate(dtWeekStart) Then
        Err.Raise 751 + vbObjectError, "timesheetweek_reject", " Invalid fromdate!"
    End If
          
    If LOppdragid < 1 Then
       Err.Raise 752 + vbObjectError, "timesheetweek_reject", " No Task has been choosen!"
    End If
    
    If ((mobjOwner Is Nothing) Or (mobjOwner.DataValues.Count = 0)) Then
       Err.Raise 753 + vbObjectError, "timesheetweek_reject", " Missing owner (KontaktID) !"
    End If
        
    'Find date end interval
    If DatePart("w", dtWeekStart, vbMonday, vbFirstFourDays) <> 7 Then
        dtTilDato = DateAdd("d", 7 - DatePart("w", dtWeekStart, vbMonday, vbFirstFourDays), dtWeekStart)
    Else
        dtTilDato = dtWeekStart
    End If
           
    Set ObjCmd = New ADODB.Command
                                                                          
    'Add input parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "vikarid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = Lvikarid
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjParam = New Parameter
    ObjParam.Name = "KontaktID"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = mobjOwner.DataValues("kontaktID").Value
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
        
    Set ObjParam = New Parameter
    ObjParam.Name = "Oppdragid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LOppdragid
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjParam = New Parameter
    ObjParam.Name = "fradato"
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
    
    If Not ObjDAL.ExecuteCRUDSP(StrConnection, "RejectCustomerTimesheet", ObjCmd.Parameters) Then
        Err.Raise 756 + vbObjectError, "timesheetweek_reject", "Could not Execute stored procedure RejectCustomerTimesheet! All tasks aborted!"
    End If
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Reject = True
    Exit Function
    
Err_handler:
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Reject = False
    WriteLog "timesheetweek_reject", Error$, EF_LogEventType_Error
End Function
