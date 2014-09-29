VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Webright"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mobjDataValues As DataValues

Public Property Set DataValues(ByVal vData As DataValues)
    Set mobjDataValues = vData
End Property

Public Property Get DataValues() As DataValues
    Set DataValues = mobjDataValues
End Property

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set mobjDataValues = Nothing
    Set objContext = Nothing
End Sub

Public Function Save(ByVal strcon As String) As Boolean

    Dim ObjDAL As IdataAccess
    Dim strSQL As String
    Dim StrParentType As String
    Dim LParentID As Long
    Dim StrParentIDname As String
    Dim Objright As Webright

On Error GoTo err_Save
        
    Save = False
               
    If DataValues.Count = 0 Then
         Err.Raise vbObjectError + 701, "", "Missing datavalues!"
    End If
    
    If Len(Trim(strcon)) < 5 Then
         Err.Raise vbObjectError + 702, "", "Invalid Connection string: " & strcon
    End If
           
    StrParentType = DataValues.Item("TYPE").Value
    Set ObjDAL = New ClsDataAccess
    If StrParentType = "KUNDE" Then
        strSQL = "delete from web_rettigheter_kontakt" & _
        " where kontaktid = " & CStr(DataValues("kontaktID").Value) & " and oppdragid = " & CStr(DataValues("oppdragID").Value) & _
        " and rettighetid = " & CStr(DataValues("RettighetID").Value)
    ElseIf StrParentType = "VIKAR" Then
        strSQL = "delete from web_rettigheter_vikar" & _
        " where vikarid = " & CStr(DataValues("vikarID").Value) & " and oppdragid = " & CStr(DataValues("oppdragID").Value) & _
        " and rettighetid = " & CStr(DataValues("RettighetID").Value)
    End If
    If Not ObjDAL.ExecuteCRUD(strcon, strSQL) Then
        Err.Raise 703, "", "Could not delete rights! SQL:" & strSQL
    End If
               
    If UCase(DataValues("checked").Value) = "CHECKED" Then
        If StrParentType = "KUNDE" Then
            strSQL = "insert into web_rettigheter_kontakt(kontaktid,oppdragid,rettighetID)" & _
            "values(" & DataValues("kontaktID").Value & "," & DataValues("oppdragID").Value & "," & DataValues("RettighetID").Value & ")"
            If Not ObjDAL.ExecuteCRUD(strcon, strSQL) Then
                Set ObjDAL = Nothing
                objContext.SetAbort
                Exit Function
            End If
        ElseIf StrParentType = "VIKAR" Then
            strSQL = "insert into web_rettigheter_vikar(VikarID,oppdragID,rettighetID)" & _
            "values(" & DataValues("VikarID").Value & "," & DataValues("oppdragID").Value & "," & DataValues("RettighetID").Value & ")"
            If Not ObjDAL.ExecuteCRUD(strcon, strSQL) Then
                Set ObjDAL = Nothing
                objContext.SetAbort
                Exit Function
             End If
        End If
    End If
    Set ObjDAL = Nothing
    Save = True
    objContext.SetComplete
    Exit Function
    
err_Save:
    Save = False
    Set ObjDAL = Nothing
    objContext.SetAbort
    WriteLog "Webright_Save", Error$, EF_LogEventType_Error
End Function


