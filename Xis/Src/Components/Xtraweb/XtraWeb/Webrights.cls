VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Webrights"
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

Public Function Add(ByVal Objright As Webright) As Webright
    Dim strKey As String
    mCol.Add Objright
    'return the object created
    Set Add = Objright
    Set Objright = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As Webright
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mCol.[_NewEnum]
End Property
Public Property Set DataValues(ByVal vData As DataValues)
    Set mobjDataValues = vData
End Property

Public Property Get DataValues() As DataValues
    Set DataValues = mobjDataValues
End Property

'Purpose:
'This function retrives all rights a consultant / customer contact has in all-
'one occurance for each right. It is to be used on the external web site,
' in the bar on the right side of personell / rekruttering-pages.
Public Function GetMenuRights() As ADODB.Recordset
On Error GoTo err_GetMenuRights
    Dim ObjDAL As IdataAccess
    Dim strSQL As String
    Dim StrParentType As String
    Dim strcon As String
        
    If DataValues.Count = 0 Then
         Err.Raise vbObjectError + 701, "", "Missing creator datavalues!"
    End If
    
    Set ObjDAL = New ClsDataAccess
    StrParentType = DataValues.Item("TYPE").Value
    strcon = DataValues.Item("XtraCon").Value
    
    If StrParentType = "KUNDE" Then
        strSQL = " Select Distinct wr.rettighetid, wr.rettighet,wr.intraNavn, wr.lokasjon, wr.web_sort" & _
                " From web_rettigheter wr, web_rettigheter_kontakt rk " & _
                " Where wr.rettighetid = rk.rettighetid" & _
                " and kontaktid = " & DataValues("KontaktID").Value & _
                " order by wr.web_sort"
    ElseIf StrParentType = "VIKAR" Then
        strSQL = " Select Distinct wr.rettighetid, wr.rettighet, wr.intraNavn, wr.lokasjon, wr.web_sort" & _
                " From web_rettigheter wr, web_rettigheter_vikar rv " & _
                " Where wr.rettighetid = rv.rettighetid" & _
                " and VikarId = " & DataValues("VikarID").Value & _
                " order by wr.web_sort"
    End If
    Set GetMenuRights = ObjDAL.FetchRC(strcon, strSQL)
    Set ObjDAL = Nothing
    Exit Function
    
err_GetMenuRights:
    Set GetMenuRights = Nothing
    Set ObjDAL = Nothing
    WriteLog "GetMenuRights", Error$, EF_LogEventType_Error
End Function

Public Function HasTaskRight(ByVal LrightID As Long, ByVal LOppdragid As Long) As Boolean
On Error GoTo err_HasTaskRight
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim strSP As String
    Dim StrParentType As String
    Dim LParentID As Long
    Dim StrParentIDname As String
    Dim strcon As String
       
    If DataValues.Count = 0 Then
         Err.Raise vbObjectError + 701, "", "Missing creator datavalues!"
    End If
    
    If Len(Trim(strcon)) < 5 Then
         Err.Raise vbObjectError + 702, "", "Invalid Connection string: " & strcon
    End If
    
    If LrightID < 1 Then
         Err.Raise vbObjectError + 703, "", "Invalid right id: " & LrightID
    End If
    
    If LOppdragid < 1 Then
         Err.Raise vbObjectError + 704, "", "Invalid oppdrag id: " & LOppdragid
    End If
           
    StrParentType = DataValues.Item("TYPE").Value
    strcon = DataValues.Item("XtraCon").Value
    
    If StrParentType = "KUNDE" Then
        StrParentIDname = "kontaktID"
        LParentID = DataValues("KontaktID").Value
        strSP = "GetCustomerTaskRight"
    ElseIf StrParentType = "VIKAR" Then
        StrParentIDname = "VikarID"
        LParentID = DataValues("VikarID").Value
        strSP = "GetConsultantTaskRight"
    End If
        
    'strConnection = "Provider=SQLOLEDB.1;User ID=ebusiness;Password=ebusiness;Data Source=PRAHA;Initial Catalog=XtraImp"
    
    Set ObjDAL = New ClsDataAccess
    Set ObjCmd = New ADODB.Command
    
    'Add parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = StrParentIDname
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LParentID
    ObjCmd.Parameters.Append ObjParam
    
    Set ObjParam = New Parameter
    ObjParam.Name = "OppdragID"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LOppdragid
    ObjCmd.Parameters.Append ObjParam
    
    Set ObjParam = New Parameter
    ObjParam.Name = "RettighetID"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LrightID
    ObjCmd.Parameters.Append ObjParam
    
    Set ObjParam = New Parameter
    ObjParam.Name = "Right"
    ObjParam.Direction = adParamOutput
    ObjParam.Type = adInteger
    ObjCmd.Parameters.Append ObjParam
        
    ObjDAL.FetchReturnValues strcon, strSP, ObjCmd
    Set ObjDAL = Nothing
    If ObjCmd.Parameters("right").Value = 1 Then
        HasTaskRight = True
    ElseIf ObjCmd.Parameters("right").Value = 0 Then
        HasTaskRight = False
    End If
    Set ObjCmd = Nothing
    Exit Function
    
err_HasTaskRight:
    HasTaskRight = False
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    WriteLog "HasRight", Error$, EF_LogEventType_Error
End Function

Public Function GetTaskRights(LOppdragid As Long) As Boolean
On Error GoTo err_GetTaskRights

    Dim ObjDAL As IdataAccess
    Dim ObjRC As Recordset
    Dim ObjFld As ADODB.Field
    Dim strSQL As String
    Dim StrParentType As String
    Dim LParentID As Long
    Dim Objright As Webright
    Dim strcon As String
        
    If DataValues.Count = 0 Then
         Err.Raise vbObjectError + 701, "", "Missing creator datavalues!"
    End If
    
    If LOppdragid < 1 Then
         Err.Raise vbObjectError + 702, "", "Invalid oppdragID:" & LOppdragid
    End If
    
    Set ObjDAL = New ClsDataAccess
    
    StrParentType = DataValues.Item("TYPE").Value
    strcon = DataValues.Item("XtraCon").Value
    If StrParentType = "KUNDE" Then
        strSQL = " select wr.rettighetid, wr.rettighet, wr.intraNavn, rk.kontaktid, rk.Oppdragid, Checked =  CASE WHEN rk.rettighetID is null THEN '' ELSE 'CHECKED' END, wr.lokasjon " & _
        " From web_rettigheter wr, web_rettigheter_kontakt rk " & _
        " Where wr.rettighetid *= rk.rettighetid" & _
        " and wr.Type =  'KUNDE'" & _
        " and rk.oppdragid = " & LOppdragid & _
        " and rk.kontaktid = " & DataValues("KontaktID").Value & _
        " order by wr.web_sort"
    ElseIf StrParentType = "VIKAR" Then
        strSQL = " Select wr.rettighetid, wr.rettighet, wr.intraNavn, rv.VikarId, rv.Oppdragid, Checked =  CASE WHEN rv.rettighetID is null THEN '' ELSE 'CHECKED' END, wr.lokasjon " & _
        " From web_rettigheter wr, web_rettigheter_vikar rv " & _
        " Where wr.rettighetid *= rv.rettighetid" & _
        " and wr.Type =  'VIKAR'" & _
        " and rv.oppdragid = " & LOppdragid & _
        " and rv.VikarId = " & DataValues("VikarID").Value & _
        " order by wr.web_sort"
    End If
    Set ObjRC = ObjDAL.FetchRC(strcon, strSQL)
    If Not ObjRC Is Nothing Then
        If ObjRC.RecordCount > 0 Then
            While Not ObjRC.EOF
                Set Objright = objContext.CreateInstance("xtraweb.Webright")
                For Each ObjFld In ObjRC.Fields
                    With Objright.DataValues
                        .Add ObjFld.Name, ObjFld.Value
                    End With
                Next ObjFld
                If StrParentType = "KUNDE" Then
                    Objright.DataValues.Add "TYPE", "KUNDE"
                ElseIf StrParentType = "VIKAR" Then
                    Objright.DataValues.Add "TYPE", "VIKAR"
                End If
                mCol.Add Objright
                Set Objright = Nothing
                ObjRC.MoveNext
            Wend
            Set ObjRC = Nothing
        Else
            Set ObjRC = Nothing
            GetTaskRights = False
            Set ObjDAL = Nothing
            Exit Function
        End If
    End If
    GetTaskRights = True
    Set ObjDAL = Nothing
    Exit Function

err_GetTaskRights:
    GetTaskRights = False
    Set ObjDAL = Nothing
    WriteLog "GetTaskRights", Error$, EF_LogEventType_Error
End Function

Public Function GetAllWebRights() As Boolean
    Dim ObjDAL As IdataAccess
    Dim ObjRC As Recordset
    Dim ObjFld As ADODB.Field
    Dim strSQL As String
    Dim StrParentType As String
    Dim LParentID As Long
    Dim Objright As Webright
    Dim strcon As String

On Error GoTo err_GetAllWebRights
        
    If DataValues.Count = 0 Then
         Err.Raise vbObjectError + 701, "", "Missing creator datavalues!"
    End If
        
    Set ObjDAL = New ClsDataAccess
    
    StrParentType = DataValues.Item("TYPE").Value
    strcon = DataValues.Item("XtraCon").Value
    
    If StrParentType = "KUNDE" Then
        LParentID = DataValues("KontaktID").Value
        strSQL = " Select Distinct wr.rettighetid, wr.intraNavn, wr.rettighet, wr.lokasjon, Checked =  CASE WHEN rk.rettighetID is null THEN '' ELSE 'CHECKED' END, wr.web_sort  " & _
        " From web_rettigheter wr, web_rettigheter_kontakt rk " & _
        " Where wr.rettighetid *= rk.rettighetid " & _
        " and wr.Type =  'KUNDE'" & _
        " and rk.kontaktid = " & LParentID & _
        " order by wr.web_sort"
    ElseIf StrParentType = "VIKAR" Then
        LParentID = DataValues("Vikarid").Value
        strSQL = " Select Distinct wr.rettighetid, wr.intraNavn, wr.rettighet, wr.lokasjon, Checked =  CASE WHEN rv.rettighetID is null THEN '' ELSE 'CHECKED' END, wr.web_sort  " & _
        " From web_rettigheter wr, web_rettigheter_vikar rv " & _
        " Where wr.rettighetid *= rv.rettighetid " & _
        " and wr.Type =  'VIKAR'" & _
        " and rv.vikarid = " & LParentID & _
        " order by wr.web_sort"
    End If
    Set ObjRC = ObjDAL.FetchRC(strcon, strSQL)
    If Not ObjRC Is Nothing Then
        If ObjRC.RecordCount > 0 Then
            While Not ObjRC.EOF
                Set Objright = objContext.CreateInstance("xtraweb.Webright")
                For Each ObjFld In ObjRC.Fields
                    With Objright.DataValues
                        .Add ObjFld.Name, ObjFld.Value
                    End With
                Next ObjFld
                If StrParentType = "KUNDE" Then
                    Objright.DataValues.Add "TYPE", "KUNDE"
                ElseIf StrParentType = "VIKAR" Then
                    Objright.DataValues.Add "TYPE", "VIKAR"
                End If
                mCol.Add Objright
                Set Objright = Nothing
                ObjRC.MoveNext
            Wend
            Set ObjRC = Nothing
        Else
            Set ObjRC = Nothing
            GetAllWebRights = False
            Set ObjDAL = Nothing
            Exit Function
        End If
    End If
    GetAllWebRights = True
    Set ObjDAL = Nothing
    Exit Function

err_GetAllWebRights:
    GetAllWebRights = False
    Set ObjDAL = Nothing
    WriteLog "GetAllWebRights", Error$, EF_LogEventType_Error
End Function

Public Function GetAllRights() As ADODB.Recordset
    Dim ObjDAL As IdataAccess
    Dim strSQL As String
    Dim StrParentType As String
    Dim LParentID As Long
    Dim strcon As String

On Error GoTo err_GetAllRights
        
    If DataValues.Count = 0 Then
         Err.Raise vbObjectError + 701, "", "Missing creator datavalues!"
    End If
    
    Set ObjDAL = New ClsDataAccess
    
    StrParentType = DataValues.Item("TYPE").Value
    strcon = DataValues.Item("XtraCon").Value
    
    If StrParentType = "KUNDE" Then
        strSQL = " select wr.rettighetid, wr.rettighet, wr.intraNavn, rk.Oppdragid, wr.lokasjon " & _
        " From web_rettigheter wr, web_rettigheter_kontakt rk " & _
        " Where wr.rettighetid = rk.rettighetid" & _
        " and kontaktid = " & DataValues("KontaktID").Value & _
        " order by wr.web_sort"
    ElseIf StrParentType = "VIKAR" Then
        strSQL = " Select wr.rettighetid, wr.rettighet, wr.intraNavn, rv.Oppdragid, wr.lokasjon " & _
        " From web_rettigheter wr, web_rettigheter_vikar rv " & _
        " Where wr.rettighetid = rv.rettighetid" & _
        " and VikarId = " & DataValues("VikarID").Value & _
        " order by wr.web_sort"
    End If
    Set GetAllRights = ObjDAL.FetchRC(strcon, strSQL)
    Exit Function
    
err_GetAllRights:
    Set GetAllRights = Nothing
    Set ObjDAL = Nothing
    WriteLog "GetAllRights", Error$, EF_LogEventType_Error
End Function

Public Function Save(ByVal strcon As String) As Boolean

    Dim ObjDAL As IdataAccess
    Dim strSQL As String
    Dim StrParentType As String
    Dim LParentID As Long
    Dim StrParentIDname As String
    Dim Objright As Webright
    Dim Bldeleted As Boolean

On Error GoTo err_Save
        
    If DataValues.Count = 0 Then
         Err.Raise vbObjectError + 701, "", "Missing creator datavalues!"
    End If
    
    If Len(Trim(strcon)) < 5 Then
         Err.Raise vbObjectError + 702, "", "Invalid Connection string: " & strcon
    End If
    
    If mCol.Count = 0 Then
         Save = False
         Exit Function
    End If

    StrParentType = DataValues.Item("TYPE").Value
    Set ObjDAL = objContext.CreateInstance("DataAccess.ClsDataAccess")
    Bldeleted = False
    For Each Objright In mCol
        If Not Objright.Save(strcon) Then
            Err.Raise 703, "", "Could Save rights! One of the rights failed to save."
        End If
    Next
    Set ObjDAL = Nothing
    Save = True
    objContext.SetComplete
    Exit Function
    
err_Save:
    Set mCol = Nothing
    Save = False
    Set ObjDAL = Nothing
    objContext.SetAbort
    WriteLog "webrights_Save", Error$, EF_LogEventType_Error
End Function

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
    Set mCol = New Collection
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set mCol = Nothing
    Set mobjDataValues = Nothing
    Set objContext = Nothing
End Sub


