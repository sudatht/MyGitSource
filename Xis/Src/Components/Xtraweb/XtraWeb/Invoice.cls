VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "Invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mstrconnection As String
Private mobjOwner As Object

Public Property Set Owner(ByVal vDataObj As Object)
    
    On Error GoTo Err_handler
    
    If (Not TypeOf vDataObj Is Consultant) And (Not TypeOf vDataObj Is Customer) Then
        Err.Raise vbObjectError + 701, "invoice_owner", "Invoiceowner is of wrong type."
    End If
    
    Set mobjOwner = vDataObj
    Exit Property
Err_handler:
    
    Set mobjOwner = Nothing
    WriteLog "invoice_Owner", Error$, EF_LogEventType_Error
End Property

Public Property Get Owner() As Object
    Set Owner = mobjOwner
End Property

Public Property Get StrConnection() As String
    StrConnection = mstrconnection
End Property

Public Function GetInvoice(ByVal StrConnection As String, ByRef strPeriod As String) As InvoiceMonth
    Dim IYear As Integer
    Dim IMonth As Integer
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim RcInvoiceData As Recordset
    Dim fldCol As Field
    Dim ObjDay As InvoiceDay
    Dim OBjWeek As InvoiceWeek
    Dim Objmonth As InvoiceMonth
    Dim strPrevWeek As String
    Dim strCurrentWeek As String
    
    'Validate the in-parameters
    On Error GoTo Err_handler
    
    If Len(Trim(StrConnection)) <= 5 Then
        Err.Raise vbObjectError + 702, "invoice_getConsultantInvoice", "Missing strConnection parameter."
    Else
        mstrconnection = StrConnection
    End If
    
    If Len(Trim(strPeriod)) = 0 Then
        Err.Raise vbObjectError + 703, "invoice_getConsultantInvoice", "Missing Period parameter."
    End If

    If ((Len(Trim(strPeriod)) < 5) Or (Len(Trim(strPeriod)) > 6)) Then
        Err.Raise vbObjectError + 704, "invoice_getConsultantInvoice", "strPeriod has a invalid parameter value. StrPeriod is " & strPeriod
    End If
       
    'Check & transfer the parameters to variables..
    IYear = val(Left(strPeriod, 4))
    If IYear > Year(Now()) Then
        Err.Raise vbObjectError + 705, "invoice_getConsultantInvoice", "strPeriod consists of a invalid year. StrPeriod is " & strPeriod
    End If
    
    IMonth = val(Right(strPeriod, 2))
    If IMonth > 12 Then
        Err.Raise vbObjectError + 706, "invoice_getConsultantInvoice", "strPeriod consists of a invalid month. StrPeriod is " & strPeriod
    End If
    
    If mobjOwner Is Nothing Then
        Err.Raise vbObjectError + 707, "invoice_getConsultantInvoice", "No owner specified."
    Else
        If TypeOf mobjOwner Is Customer Then
            Err.Raise vbObjectError + 708, "invoice_getConsultantInvoice", "Wrong owner type for this type of operation! Owner must be Consultant"
        End If
    End If
       
    'Retreive the invoice data for consultant
    Set ObjCmd = New ADODB.Command
       
    'Add input parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "vikarid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = CLng(mobjOwner.DataValues("vikarid").Value)
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjParam = New Parameter
    ObjParam.Name = "aar"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = IYear
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
       
    Set ObjParam = New Parameter
    ObjParam.Name = "maaned"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = IMonth
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
        
    Set ObjDAL = New ClsDataAccess
    ' Get the Consultants invoice basis
    Set RcInvoiceData = ObjDAL.FetchRCSP(mstrconnection, "GetConsultantInvoiceBasis", ObjCmd.Parameters)
    Set ObjDAL = Nothing
    strPrevWeek = "0"
    'Build the various collections and classes..
    Set Objmonth = objContext.CreateInstance("xtraweb.InvoiceMonth")
    Set OBjWeek = objContext.CreateInstance("xtraweb.InvoiceWeek")
    While Not RcInvoiceData.EOF
        Set ObjDay = objContext.CreateInstance("xtraweb.InvoiceDay")
         'For all values for each day..
        For Each fldCol In RcInvoiceData.Fields
            ObjDay.DataValues.Add fldCol.Name, fldCol.Value
        Next
         strCurrentWeek = CStr(Format(CDate(ObjDay.DataValues("dagdato").Value), "ww", vbMonday))
        'If first day..
        If strPrevWeek = "0" Then
            strPrevWeek = strCurrentWeek
        End If
        'Check to see if new week..
        If strCurrentWeek <> strPrevWeek Then
            Set OBjWeek.Owner = SafeRef(Me)
            OBjWeek.Refresh
            Objmonth.Add OBjWeek
            Set OBjWeek = objContext.CreateInstance("xtraweb.InvoiceWeek")
        End If
        strPrevWeek = strCurrentWeek
        OBjWeek.Add ObjDay
        RcInvoiceData.MoveNext
    Wend
    If OBjWeek.Count > 0 Then
        Set OBjWeek.Owner = SafeRef(Me)
        OBjWeek.Refresh
        Objmonth.Add OBjWeek
    End If
    'Clean up temp objects
    Set RcInvoiceData = Nothing
    Set ObjDAL = Nothing
    'return the collection..
    Set GetInvoice = Objmonth
    Set Objmonth = Nothing
    Exit Function

Err_handler:
    Set GetInvoice = Nothing
    WriteLog "invoice_getConsultantInvoice", Error$, EF_LogEventType_Error
End Function

Public Function GetCustomerInvoice(ByVal StrConnection As String, ByRef strPeriod As String) As InVoiceConsultants
    Dim IYear As Integer
    Dim IMonth As Integer
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim RcInvoiceData As Recordset
    Dim fldCol As Field
    Dim ObjDay As InvoiceDay
    Dim Objmonth As InvoiceCustomerMonth
    Dim Lvikarid As Long
    Dim LPrevVikarid As Long
    Dim objConsultant As Consultant
    Dim objConsultants As InVoiceConsultants

    'Validate the in-parameters
    On Error GoTo Err_handler
    
    If Len(Trim(StrConnection)) <= 5 Then
        Err.Raise vbObjectError + 702, "invoice_getCustomerInvoice", "Missing strConnection parameter."
    Else
        mstrconnection = StrConnection
    End If
    
    If Len(Trim(strPeriod)) = 0 Then
        Err.Raise vbObjectError + 703, "invoice_getCustomerInvoice", "Missing Period parameter."
    End If

    If ((Len(Trim(strPeriod)) < 5) Or (Len(Trim(strPeriod)) > 6)) Then
        Err.Raise vbObjectError + 704, "invoice_getCustomerInvoice", "strPeriod has a invalid parameter value. Strperiod is " & strPeriod
    End If
       
    IYear = val(Left(strPeriod, 4))
    IMonth = val(Right(strPeriod, 2))
    If IMonth > 12 Then
        Err.Raise vbObjectError + 705, "invoice_getCustomerInvoice", "strPeriod consists of a invalid month.Strperiod is " & strPeriod
    End If
    
    If mobjOwner Is Nothing Then
        Err.Raise vbObjectError + 706, "invoice_getCustomerInvoice", "No owner specified."
    Else
        If TypeOf mobjOwner Is Consultant Then
            Err.Raise vbObjectError + 707, "invoice_getCustomerInvoice", "Wrong owner type for this type of operation! Owner must be Customer!"
        End If
    End If
        
    'Retreive the invoice data for consultant
    Set ObjCmd = New ADODB.Command
    
    'Add input parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "ifirmaid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = mobjOwner.DataValues("firmaid").Value
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
                     
    Set ObjParam = New Parameter
    ObjParam.Name = "ifaktperiode"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = strPeriod
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing

    Set ObjParam = New Parameter
    ObjParam.Name = "iKontaktId"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = mobjOwner.DataValues("KontaktID").Value
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
        
    Set ObjDAL = New ClsDataAccess
    
    ' Get the Consultants invoice basis
    Set RcInvoiceData = ObjDAL.FetchRCSP(mstrconnection, "GetCustomerInvoiceBasis", ObjCmd.Parameters)
    If Not IsObject(RcInvoiceData) Then
        Err.Raise "708" + vbObjectError, "invoice_getCustomerInvoice", "Error in or during execution of stored procedure GetCustomerInvoiceBasis!"
    End If
    Set ObjDAL = Nothing
    'Build the various collections and classes..
    Set Objmonth = objContext.CreateInstance("xtraweb.InvoiceCustomerMonth")
              
    If Not RcInvoiceData.EOF Then
        Set objConsultants = objContext.CreateInstance("xtraweb.InVoiceConsultants")
        Set Objmonth = objContext.CreateInstance("xtraweb.InvoiceCustomerMonth")
        While Not RcInvoiceData.EOF
            Lvikarid = RcInvoiceData("vikarid").Value
            'for each consultant add month and refresh sums..
            If LPrevVikarid <> Lvikarid Then
                If Not objConsultant Is Nothing Then
                    Set Objmonth.Owner = objConsultant
                    Set objConsultant.InvoiceMonth = Objmonth
                    Objmonth.Refresh
                    objConsultants.Add objConsultant
                End If
                '(re-)initialize values..
                Set Objmonth = objContext.CreateInstance("xtraweb.InvoiceCustomerMonth")
                Set objConsultant = objContext.CreateInstance("xtraweb.Consultant")
                objConsultant.XtraConString = mstrconnection
                objConsultant.GetConsultant (Lvikarid)
            End If
            Set ObjDay = objContext.CreateInstance("xtraweb.InvoiceDay")
            'all values for each day..
            For Each fldCol In RcInvoiceData.Fields
                ObjDay.DataValues.Add fldCol.Name, fldCol.Value
            Next
            LPrevVikarid = Lvikarid
            Set ObjDay.Owner = SafeRef(Me)
            Objmonth.Add ObjDay
            RcInvoiceData.MoveNext
        Wend
        'Add last occurance to collection
        If Objmonth.Count > 0 Then
            Set Objmonth.Owner = objConsultant
            Set objConsultant.InvoiceMonth = Objmonth
            Objmonth.Refresh
            objConsultants.Add objConsultant
        End If
    End If
    'Clean up temp objects
    Set RcInvoiceData = Nothing
    Set ObjDAL = Nothing
    'return the collection..
    Set GetCustomerInvoice = objConsultants
    Exit Function
    
Err_handler:
    Set GetCustomerInvoice = Nothing
    WriteLog "invoice_getCustomerInvoice", Error$, EF_LogEventType_Error
End Function

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Public Sub CleanUp()
    Set mobjOwner = Nothing
End Sub

Private Sub ObjectControl_Deactivate()
    Set mobjOwner = Nothing
    Set objContext = Nothing
End Sub

