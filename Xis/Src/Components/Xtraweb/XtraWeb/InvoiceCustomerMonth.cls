VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "InvoiceCustomerMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mCol As Collection
Private mobjConsultant As Consultant
Private mobjMonthlySums As invoiceCustomerMonthSums

Public Property Get Owner() As Consultant
    Set Owner = mobjConsultant
End Property

Public Property Set Owner(ByVal objConsultant As Consultant)
    Set mobjConsultant = objConsultant
End Property

Public Property Get MonthlySums() As invoiceCustomerMonthSums
    Set MonthlySums = mobjMonthlySums
End Property

Public Function Add(ByRef ObjDay As InvoiceDay) As InvoiceDay
    mCol.Add ObjDay
    'return the object created
    Set Add = ObjDay
    Set ObjDay = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As InvoiceDay
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

Private Sub ObjectControl_Deactivate()
    Set mCol = Nothing
    Set mobjConsultant = Nothing
    Set mobjMonthlySums = Nothing
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Public Sub Refresh()
    On Error GoTo Err_handler
    If Not mobjConsultant Is Nothing Then
        Set mobjMonthlySums = objContext.CreateInstance("xtraweb.invoiceCustomerMonthSums")
        Call GetMonthSum
        Call GetMonthOverTime
    Else
        Err.Raise 701 + vbObjectError, "invoiceCustomer_refresh", "No owner!"
    End If
    Exit Sub

Err_handler:
    WriteLog "invoiceCustomer_refresh", Error$, EF_LogEventType_Error
End Sub

Private Sub GetMonthSum()
    'Data access object used to execute SQL, stored procedures and so forth.
    Dim ObjDAL As IdataAccess
    'The consultant id to retreive the weekly monthly for.
    Dim rsInvoiceSum As ADODB.Recordset
    'temporary field to store datavalues
    Dim fldCol As ADODB.Field
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim LfirmaID As Long
    Dim lConsultantID As Long
    Dim dtfradato As Date
    Dim dtTilDato As Date
    Dim Lvikarid As Long
    Dim LbestAv As Long
    Dim ObjMonthSum As InvoiceCustomerMonthSum
    Dim mstrconnection As String
     
    On Error GoTo Err_handler
    
    If mCol.Count < 1 Then
        Err.Raise 701 + vbObjectError, "invoiceCustomer_GetMonthSum", "Invalid amount of invoiceDays"
    End If
       
    LfirmaID = mCol(1).DataValues("firmaid")
    If LfirmaID < 1 Then
        Err.Raise 702 + vbObjectError, "invoiceCustomer_GetMonthSum", "Invalid firmaid! Firmaid =" & LfirmaID
    End If
    
    If mobjConsultant Is Nothing Then
        Err.Raise 703 + vbObjectError, "invoiceCustomer_GetMonthSum", "No owner is specified!"
    End If
      
    Lvikarid = mobjConsultant.DataValues("vikarid")
    If Lvikarid < 1 Then
        Err.Raise 704 + vbObjectError, "invoiceCustomer_GetMonthSum", "Invalid vikarid! vikarid =" & Lvikarid
    End If
   
    LbestAv = mCol(1).DataValues("bestilltav")
    If LbestAv < 1 Then
        Err.Raise 705 + vbObjectError, "invoiceCustomer_GetMonthSum", "Invalid kontaktid! Kontaktid =" & LbestAv
    End If

    dtfradato = mCol(1).DataValues("dato")
    dtTilDato = mCol(mCol.Count).DataValues("dato")
    If ((Not IsDate(dtfradato)) Or (Not IsDate(dtTilDato))) Then
        Err.Raise 706 + vbObjectError, "invoiceCustomer_GetMonthSum", "Invalid from/to dato! Dates =" & dtfradato & " - " & dtTilDato
    End If
    
    mstrconnection = mCol(1).Owner.StrConnection
    If Len(mstrconnection) < 5 Then
        Err.Raise 707 + vbObjectError, "invoiceCustomer_GetMonthSum", "Invalid connection string!"
    End If
    
    'Retreive the invoice month sum for customer
    Set ObjCmd = New ADODB.Command
        
    'Add input parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "dtStartDato"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adDate
    ObjParam.Value = dtfradato
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjParam = New Parameter
    ObjParam.Name = "dtSluttDato"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adDate
    ObjParam.Value = dtTilDato
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
        
    Set ObjParam = New Parameter
    ObjParam.Name = "ifirmaid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LfirmaID
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing

    Set ObjParam = New Parameter
    ObjParam.Name = "Lkontaktid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LbestAv
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjParam = New Parameter
    ObjParam.Name = "ivikarid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = Lvikarid
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjDAL = New ClsDataAccess
    
    ' Get the Consultants invoice basis
    Set rsInvoiceSum = ObjDAL.FetchRCSP(mstrconnection, "GetCustomerMonthInvoiceSum", ObjCmd.Parameters)
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    If rsInvoiceSum Is Nothing Then
        Err.Raise 709 + vbObjectError, "invoiceCustomer_GetMonthSum", "Error in or during execution of stored procedure GetCustomerInvoiceBasis!"
    End If

    If Not rsInvoiceSum.EOF Then
        Do Until rsInvoiceSum.EOF
            Set ObjMonthSum = objContext.CreateInstance("xtraweb.InvoiceCustomerMonthSum")
            For Each fldCol In rsInvoiceSum.Fields
                If IsNull(fldCol.Value) Then
                    ObjMonthSum.DataValues.Add fldCol.Name, "0"
                Else
                    ObjMonthSum.DataValues.Add fldCol.Name, fldCol.Value
                End If
            Next
            ObjMonthSum.DataValues.Add "SumType", "Ordinære timer"
            mobjMonthlySums.Add ObjMonthSum
            rsInvoiceSum.MoveNext
        Loop
    End If
    Set rsInvoiceSum = Nothing
    Exit Sub
    
Err_handler:
    WriteLog "invoiceCustomer_GetMonthSum", Error$, EF_LogEventType_Error
End Sub

Private Sub GetMonthOverTime()
    'Data access object used to execute SQL, stored procedures and so forth.
    Dim ObjDAL As IdataAccess
    'The consultant id to retreive the weekly monthly for.
    Dim rsInvoiceOvertime As ADODB.Recordset
    'temporary field to store datavalues
    Dim fldCol As ADODB.Field
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim LfirmaID As Long
    Dim lConsultantID As Long
    Dim LukeStart As Long
    Dim LukeSlutt As Long
    Dim Lvikarid As Long
    Dim LbestAv As Long
    Dim ObjMonthSum As InvoiceCustomerMonthSum
    Dim mstrconnection As String
    Dim AOverTimeType(1, 2)
    Dim LSumTeller As Long
     
    On Error GoTo Err_handler
    
    If mCol.Count > 0 Then
       
        LfirmaID = mCol(1).DataValues("firmaid")
        If LfirmaID < 1 Then
            Err.Raise 702 + vbObjectError, "invoiceCustomer_GetMonthOverTime", "Invalid firmaid! Firmaid =" & LfirmaID
        End If
        
        If mobjConsultant Is Nothing Then
            Err.Raise 703 + vbObjectError, "invoiceCustomer_GetMonthOverTime", "No owner is specified!"
        End If
          
        Lvikarid = mobjConsultant.DataValues("vikarid")
        If Lvikarid < 1 Then
            Err.Raise 704 + vbObjectError, "invoiceCustomer_GetMonthOverTime", "Invalid vikarid! vikarid =" & Lvikarid
        End If
       
        LbestAv = mCol(1).DataValues("bestilltav")
        If LbestAv < 1 Then
            Err.Raise 705 + vbObjectError, "invoiceCustomer_GetMonthOverTime", "Invalid kontaktid! Kontaktid =" & LbestAv
        End If
    
        If (DatePart("m", mCol(1).DataValues("dato"), vbMonday, vbFirstFourDays) = 1) And (DatePart("ww", mCol(1).DataValues("dato"), vbMonday, vbFirstFourDays) > 10) Then
            LukeStart = CLng(DatePart("yyyy", mCol(1).DataValues("dato"), vbMonday, vbFirstFourDays) - 1 & DatePart("ww", mCol(1).DataValues("dato"), vbMonday, vbFirstFourDays))
        Else
            LukeStart = CLng(DatePart("yyyy", mCol(1).DataValues("dato"), vbMonday, vbFirstFourDays) & leftpad(DatePart("ww", mCol(1).DataValues("dato"), vbMonday, vbFirstFourDays), 2, "0"))
        End If
        
        If (DatePart("m", mCol(1).DataValues("dato"), vbMonday, vbFirstFourDays) = 12) And (DatePart("ww", mCol(1).DataValues("dato"), vbMonday, vbFirstFourDays) < 52) Then
            LukeSlutt = CLng(DatePart("yyyy", mCol(mCol.Count).DataValues("dato"), vbMonday, vbFirstFourDays) + 1 & DatePart("ww", mCol(mCol.Count).DataValues("dato"), vbMonday, vbFirstFourDays))
        Else
            LukeSlutt = CLng(DatePart("yyyy", mCol(mCol.Count).DataValues("dato"), vbMonday, vbFirstFourDays) & leftpad(DatePart("ww", mCol(mCol.Count).DataValues("dato"), vbMonday, vbFirstFourDays), 2, "0"))
        End If
        
        If ((LukeStart < 1) Or (LukeSlutt < 1)) Then
            Err.Raise 706 + vbObjectError, "invoiceCustomer_GetMonthSum", "Invalid from/to Weeks! Weeks =" & LukeStart & " - " & LukeSlutt
        End If
        
        mstrconnection = mCol(1).Owner.StrConnection
        If Len(mstrconnection) < 5 Then
            Err.Raise 707 + vbObjectError, "invoiceCustomer_GetMonthOverTime", "Invalid connection string!"
        End If
        
        'Retreive the invoice month sum for customer
        Set ObjCmd = New ADODB.Command
            
        'Add input parameters to command object
        Set ObjParam = New Parameter
        ObjParam.Name = "LStartUke"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = LukeStart
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
               
        Set ObjParam = New Parameter
        ObjParam.Name = "LSluttUke"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = LukeSlutt
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
                
        Set ObjParam = New Parameter
        ObjParam.Name = "ifirmaid"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = LfirmaID
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
    
        Set ObjParam = New Parameter
        ObjParam.Name = "Lkontaktid"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = LbestAv
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
        
        Set ObjParam = New Parameter
        ObjParam.Name = "   ivikarid"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = Lvikarid
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
        
        Set ObjDAL = New ClsDataAccess
        
        AOverTimeType(0, 0) = "GetCustomerMonthInvoiceOT35"
        AOverTimeType(0, 1) = "Overtid 35%"
        AOverTimeType(1, 0) = "GetCustomerMonthInvoiceOT70"
        AOverTimeType(1, 1) = "Overtid 70%"
     
        For LSumTeller = 0 To 1
            ' Get the Consultants invoice overtime sums..
            Set rsInvoiceOvertime = ObjDAL.FetchRCSP(mstrconnection, AOverTimeType(LSumTeller, 0), ObjCmd.Parameters)
            If rsInvoiceOvertime Is Nothing Then
                Err.Raise 709 + vbObjectError, "invoiceCustomer_GetMonthOverTime", "Error in or during execution of stored procedure " & AOverTimeType(LSumTeller, 0) & "!"
            End If
        
            If Not rsInvoiceOvertime.EOF Then
                Do Until rsInvoiceOvertime.EOF
                    Set ObjMonthSum = objContext.CreateInstance("xtraweb.InvoiceCustomerMonthSum")
                    For Each fldCol In rsInvoiceOvertime.Fields
                        ObjMonthSum.DataValues.Add fldCol.Name, fldCol.Value
                    Next
                    ObjMonthSum.DataValues.Add "SumType", AOverTimeType(LSumTeller, 1)
                    mobjMonthlySums.Add ObjMonthSum
                    rsInvoiceOvertime.MoveNext
                Loop
            End If
            Set rsInvoiceOvertime = Nothing
        Next LSumTeller
        Set ObjDAL = Nothing
    End If
    Exit Sub
    
Err_handler:
    WriteLog "invoiceCustomer_GetMonhOverTime", Error$, EF_LogEventType_Error
End Sub
