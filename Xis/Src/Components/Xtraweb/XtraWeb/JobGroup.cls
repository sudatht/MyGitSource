VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "JobGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mobjDataValues As DataValues
Private mobjJobTypes As JobTypes

Public Property Set JobTypes(ByVal vData As JobTypes)
    Set mobjJobTypes = vData
End Property

Public Property Get JobTypes() As JobTypes
    Set JobTypes = mobjJobTypes
End Property

Public Property Set DataValues(ByVal vData As DataValues)
    Set mobjDataValues = vData
End Property

Public Property Get DataValues() As DataValues
Attribute DataValues.VB_UserMemId = 0
    Set DataValues = mobjDataValues
End Property

Public Sub Refresh(strXtraCon As String, lCVId As Long)
On Error GoTo err_Refresh
    Dim ObjDAL          As IdataAccess
    Dim rsUnder         As ADODB.Recordset
    Dim fldData         As ADODB.Field
    Dim strSQL          As String
    Dim IJobGroupID     As Integer
    Dim objJobtype      As JobType
    Dim strID           As String
    Dim LjobGroup       As Long
    
    If lCVId < 1 Then
       Err.Raise 700 + vbObjectError, "", " Missing CVid parameter!"
    End If
    
    If Trim(Len(strXtraCon)) = 0 Then
       Err.Raise 701 + vbObjectError, "", " Missing connection string parameter!"
    End If

    IJobGroupID = mobjDataValues("FagID").Value

    strSQL = "SELECT DISTINCT [kt].[K_TittelID], [kt].[KTittel], [kt].[FagID], " & _
    "Checked = CASE WHEN [vk].[CVid] IS NULL THEN '' ELSE 'checked' END " & _
    "FROM [h_komp_tittel] AS [kt], [vikar_kompetanse] AS [vk] " & _
    "WHERE [kt].[k_typeid] = 4 " & _
    "AND [vk].[K_TittelID] =* [kt].[K_TittelID]" & _
    "AND [kt].[web_jobwish_visible] = 1 " & _
    "AND [vk].[cvid] = " & lCVId & " " & _
    "AND [kt].[FagID] = " & IJobGroupID & " " & _
    "ORDER BY [kt].[k_tittelID]"

    'strSQL = "select distinct kt.K_TittelID, kt.KTittel, kt.tomid, " & _
    '"Checked = CASE WHEN vk.CVid is null THEN '' ELSE 'checked' END  " & _
    '"From h_komp_tittel kt, vikar_kompetanse vk " & _
    '"Where kt.k_typeid = 4 " & _
    '"and vk.K_TittelID =* kt.K_TittelID " & _
    '"and kt.web_jobwish_visible = 1 " & _
    '"and vk.cvid = " & lCVId & " " & _
    '"and kt.tomid = " & IJobGroupID & " " & _
    '"order by kt.k_tittelID "
    
    Set ObjDAL = New ClsDataAccess
    
    Set rsUnder = ObjDAL.FetchRC(strXtraCon, strSQL)
    Set ObjDAL = Nothing

    While Not rsUnder.EOF
        Set objJobtype = objContext.CreateInstance("xtraweb.JobType")
        For Each fldData In rsUnder.Fields
            objJobtype.DataValues(fldData.Name) = fldData.Value
        Next
        If LCase(objJobtype.DataValues("Checked").Value) = "checked" Then
            objJobtype.Checked = True
        End If
        mobjJobTypes.Add objJobtype
        rsUnder.MoveNext
    Wend
    Set rsUnder = Nothing
    Set ObjDAL = Nothing
    Exit Sub
    
err_Refresh:
    WriteLog "JobGroup_Refresh", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set rsUnder = Nothing
End Sub

Public Sub RefreshChanged(strXtraCon As String, lCVId As Long)
On Error GoTo err_Refresh
    Dim ObjDAL          As IdataAccess
    Dim rsUnder         As ADODB.Recordset
    Dim fldData         As ADODB.Field
    Dim strSQL          As String
    Dim RsWishes        As ADODB.Recordset
    Dim IJobGroupID     As Integer
    Dim objJobtype      As JobType
    Dim strID           As String

    If lCVId < 1 Then
       Err.Raise 700 + vbObjectError, "", " Missing CVid parameter!"
    End If
    
    If Trim(Len(strXtraCon)) = 0 Then
       Err.Raise 701 + vbObjectError, "", " Missing connection string parameter!"
    End If

    IJobGroupID = mobjDataValues("TomID").Value


    strSQL = "Select count(*) from CV_vikar_jobwishes_temp " & _
            "Where cvid = " & lCVId

    Set ObjDAL = New ClsDataAccess
    Set RsWishes = ObjDAL.FetchRC(strXtraCon, strSQL)
    If RsWishes(0) = 0 Then
        strSQL = "Select k_tittelID from vikar_kompetanse where k_typeid = 4 and cvid = " & lCVId
        Set RsWishes = ObjDAL.FetchRC(strXtraCon, strSQL)
        If Not RsWishes.EOF Then
            With RsWishes
                While Not .EOF
                    strSQL = "Insert into CV_vikar_jobwishes_temp(CVId, WishID, Original_value, Type) values(" & lCVId & "," & .Fields("k_tittelID").Value & ", 1, 'ORIGINAL')"
                    If Not ObjDAL.ExecuteCRUD(strXtraCon, strSQL) Then
                        Err.Raise 701 + vbObjectError, "", "Error during insert! (Wishes) StrSQL:" & strSQL
                    End If
                    RsWishes.MoveNext
                Wend
            End With
        End If
    End If
    
    Set RsWishes = Nothing

    strSQL = "select jw.type, kt.k_tittelID, kt.KTittel, kt.tomid, " & _
    "Checked = CASE WHEN jw.CVid is null THEN '' ELSE 'checked' END " & _
    "From h_komp_tittel kt, CV_vikar_jobwishes_temp jw " & _
    "Where kt.k_typeid = 4 " & _
    "and jw.wishid =* kt.K_TittelID " & _
    "and kt.web_jobwish_visible = 1 " & _
    "and jw.cvid = " & lCVId & " " & _
    "and kt.tomid = " & IJobGroupID & " " & _
    "order by kt.k_tittelID "
       
    
    Set mobjJobTypes = Nothing
    Set mobjJobTypes = objContext.CreateInstance("xtraweb.jobtypes")
    Set RsWishes = ObjDAL.FetchRC(strXtraCon, strSQL)
    Set ObjDAL = Nothing
    While Not RsWishes.EOF
        Set objJobtype = objContext.CreateInstance("xtraweb.jobtype")
        For Each fldData In RsWishes.Fields
            objJobtype.DataValues(fldData.Name).Value = fldData.Value
        Next
        If LCase(objJobtype.DataValues("Checked").Value) = "checked" Then
            objJobtype.Checked = True
        End If
        mobjJobTypes.Add objJobtype
        RsWishes.MoveNext
    Wend
    Set RsWishes = Nothing
    Exit Sub

err_Refresh:
    WriteLog "JobGroup_RefreshChanged", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set RsWishes = Nothing
End Sub

Public Function GetAllNonApproved(strXtraCon As String, lCVId As Long) As ADODB.Recordset
On Error GoTo err_GetAllNonApproved
    Dim ObjDAL              As IdataAccess
    Dim strSQL              As String
    
    If lCVId < 1 Then
       Err.Raise 700 + vbObjectError, "", " Missing CVid parameter!"
    End If
    
    If Trim(Len(strXtraCon)) = 0 Then
       Err.Raise 701 + vbObjectError, "", " Missing connection string parameter!"
    End If
                
    strSQL = "select jt.type, tr.navn, kt.ktittel, kt.k_tittelid " & _
        "from CV_vikar_jobwishes_temp jt, h_komp_tittel kt, tjenesteomrade tr " & _
        "where web_jobwish_visible = 1 " & _
        "and kt.k_typeid = 4 " & _
        "and jt.type <> 'ORIGINAL' " & _
        "and jt.wishid = kt.k_tittelID " & _
        "and tr.tomid = kt.tomid " & _
        "and cvid = " & lCVId & " " & _
        "order by tr.navn, kt.ktittel "
   
    Set ObjDAL = New ClsDataAccess
    Set GetAllNonApproved = ObjDAL.FetchRC(strXtraCon, strSQL)
    Set ObjDAL = Nothing
    Exit Function

err_GetAllNonApproved:
    WriteLog "JobGroup_GetAllNonApproved", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set GetAllNonApproved = Nothing
End Function

Public Function GetAllApproved(strXtraCon As String, lCVId As Long) As ADODB.Recordset
On Error GoTo err_GetAllApproved
    Dim ObjDAL              As IdataAccess
    Dim strSQL              As String
    
    If lCVId < 1 Then
       Err.Raise 700 + vbObjectError, "", " Missing CVid parameter!"
    End If
    
    If Trim(Len(strXtraCon)) = 0 Then
       Err.Raise 701 + vbObjectError, "", " Missing connection string parameter!"
    End If
                
    'strSQL = "select tr.navn, kt.ktittel, vk.k_tittelID " & _
    '    "from vikar_kompetanse vk, h_komp_tittel kt, tjenesteomrade tr " & _
    '    "Where kt.web_jobwish_visible = 1 " & _
    '    "and cvid = " & lCVId & " " & _
    '    "and vk.k_typeid = 4 " & _
    '    "and vk.k_tittelID = kt.k_tittelID " & _
    '    "and tr.tomid = kt.tomid " & _
    '    "order by tr.navn, kt.ktittel "
        
    strSQL = " SELECT [TR].[FAGOMRADE], [KT].[KTITTEL], [VK].[K_TITTELID] " _
             & "FROM [VIKAR_KOMPETANSE] AS [VK], [H_KOMP_TITTEL] AS [KT], [H_KOMP_FAGOMRADE] AS [TR] " _
             & "WHERE [KT].[WEB_JOBWISH_VISIBLE] = 1 AND [CVID] = " & lCVId & " AND [VK].[K_TYPEID] = 4 " _
             & "AND [VK].[K_TITTELID] = [KT].[K_TITTELID] AND [TR].[FAGID] = [KT].[FAGID] ORDER BY [TR].[FAGOMRADE], [KT].[KTITTEL]"


       
    Set ObjDAL = New ClsDataAccess
    Set GetAllApproved = ObjDAL.FetchRC(strXtraCon, strSQL)
    Set ObjDAL = Nothing
    Exit Function

err_GetAllApproved:
    WriteLog "JobGroup_GetAllApproved", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set GetAllApproved = Nothing
End Function

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
    Set mobjJobTypes = objContext.CreateInstance("xtraweb.JobTypes")
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set mobjDataValues = Nothing
    Set mobjJobTypes = Nothing
    Set objContext = Nothing
End Sub

