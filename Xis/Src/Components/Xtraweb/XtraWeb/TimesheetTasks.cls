VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "TimesheetTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mCol As Collection
Private mobjDataValues As DataValues
Private mobjOwner As Object

Public Property Set Owner(ByVal vDataObj As Object)
    
    On Error GoTo Err_handler
    
    If Not vDataObj Is Nothing Then
        If ((Not TypeOf vDataObj Is Consultant) And (Not TypeOf vDataObj Is Customer)) Then
            Err.Raise vbObjectError + 701, "TimesheetTask_owner", "TimesheetTask owner is of wrong type."
        End If
        Set mobjOwner = vDataObj
    End If
    Exit Property
Err_handler:
    
    Set mobjOwner = Nothing
    WriteLog "timesheetTasks_Owner", Error$, EF_LogEventType_Error
End Property

Public Property Get Owner() As Object
    Set Owner = mobjOwner
End Property

Public Property Set DataValues(ByVal vData As DataValues)
    Set mobjDataValues = vData
End Property

Public Property Get DataValues() As DataValues
    Set DataValues = mobjDataValues
End Property

Public Function Add(ByVal objNewMember As TimesheetTask, Optional strKey As String) As TimesheetTask
    
    If Len(strKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, strKey
    End If
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As TimesheetTask
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

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mCol = New Collection
End Sub

Public Sub CleanUp()
    'destroys collection when this class is terminated
    Set mCol = Nothing
    Set mobjDataValues = Nothing
    Set mobjOwner = Nothing
End Sub

Private Sub ObjectControl_Deactivate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
    Set mobjDataValues = Nothing
    Set mobjOwner = Nothing
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Public Function GetTasksForWeek(ByVal StrConnection As String, ByVal dtWeekStart As Date) As Boolean
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim ObjTask As TimesheetTask
    Dim ObjFld As Field
    Dim ObjTasks As TimesheetTasks
    Dim lID As Long
    Dim strOwnerType As String
    Dim StrProcedure As String
    Dim dtTilDato As Date
    Dim rsTasksforWeek As ADODB.Recordset
    Dim IFraMaaned As Integer
    Dim ITilMaaned As Integer
    Dim LCurrentOppdragid As Long
    Dim LPrevOppdragid As Long

On Error GoTo Err_handler

    If Len(Trim(StrConnection)) <= 5 Then
        Err.Raise vbObjectError + 700, "TimesheetTasks_GetTasksForWeek", "Missing strConnection parameter."
    End If

    If mobjOwner Is Nothing Then
       Err.Raise 701 + vbObjectError, "TimesheetTasks_GetTasksForWeek", " No consultant/customer has been choosen (Lvikarid/Lfirmaid) !"
    End If
          
    If Not IsDate(dtWeekStart) Then
        Err.Raise 704 + vbObjectError, "TimesheetTasks_GetTasksForWeek", " Invalid fromdate!"
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
        strOwnerType = "vikarid"
        StrProcedure = "GetConsultantTasksForWeek"
        If lID < 1 Then
           Err.Raise 702 + vbObjectError, "TimesheetTask_GetTasksForWeek", " Invalid consultant id !"
        End If
    ElseIf TypeOf mobjOwner Is Customer Then
        lID = mobjOwner.DataValues("firmaid").Value
        strOwnerType = "firmaid"
        StrProcedure = "GetCustomerTasksForWeek"
        If lID < 1 Then
           Err.Raise 703 + vbObjectError, "TimesheetTask_GetTasksForWeek", " Invalid customer id !"
        End If
    End If
    
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
    
    With mobjDataValues
        .Add "Lweekno", DatePart("ww", dtWeekStart, vbMonday, vbFirstFourDays)
        .Add "year", DatePart("y", dtWeekStart, vbMonday, vbFirstFourDays)
        .Add "Weekstart", dtWeekStart
        .Add "tildato", dtTilDato
        .Add strOwnerType, lID
        .Add "strconnection", StrConnection
    End With
    
    Set ObjCmd = New ADODB.Command
                      
    'Add input parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = strOwnerType
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = lID
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    If Not mobjOwner Is Nothing Then
        If TypeOf mobjOwner Is Customer Then
            Set ObjParam = New Parameter
            ObjParam.Name = "kontaktID"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = mobjOwner.DataValues("kontaktID").Value
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
        End If
    End If

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
    
    Set rsTasksforWeek = ObjDAL.FetchRCSP(StrConnection, StrProcedure, ObjCmd.Parameters)
    Set ObjCmd = Nothing
    If Not IsObject(rsTasksforWeek) Then
        Err.Raise 704 + vbObjectError, "TimesheetTasks_GetTasksForWeek", "Error in GetTasksForWeek!"
    End If

    'Dersom tasks..
    If Not rsTasksforWeek.EOF Then
        LPrevOppdragid = 0
        While Not rsTasksforWeek.EOF
            'Legg til taskene til Tasks col'en
            LCurrentOppdragid = rsTasksforWeek("Oppdragid").Value
            If (LCurrentOppdragid <> LPrevOppdragid) Then
                Set ObjTask = objContext.CreateInstance("xtraweb.TimesheetTask")
                For Each ObjFld In rsTasksforWeek.Fields
                    ObjTask.DataValues.Add ObjFld.Name, ObjFld.Value
                Next
                mCol.Add ObjTask
                LPrevOppdragid = LCurrentOppdragid
            End If
            rsTasksforWeek.MoveNext
        Wend
        Set rsTasksforWeek = Nothing
    'Dersom ingen tasks..
    Else
        GetTasksForWeek = False
        Exit Function
    End If

    GetTasksForWeek = True
    Set ObjDAL = Nothing
    Exit Function

Err_handler:
    GetTasksForWeek = False
    WriteLog "timesheets_Gettaskforweek", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Function
