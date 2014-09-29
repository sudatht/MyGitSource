VERSION 5.00
Begin VB.Form frmDriver 
   Caption         =   "Integration Driver"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   ScaleHeight     =   7365
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtXML2 
      Height          =   3255
      Index           =   1
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3960
      Width           =   6975
   End
   Begin VB.TextBox txtXML 
      Height          =   3135
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   720
      Width           =   6975
   End
   Begin VB.CommandButton btnGetXMLRootContacts 
      Caption         =   "Xml Root Contacts"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton btnTest 
      Caption         =   "Test"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGetXMLRootContacts_Click()
    Dim cts As Integration.SuperOffice
    Dim teller
    
    Set cts = New Integration.SuperOffice
    Dim startTime As Date
    
    startTime = Now()
    For teller = 1 To 100
        txtXML(0).Text = cts.XMLGetRootContacts()
    Next teller
    Debug.Print (DatePart("s", Now() - startTime, 2, 2))
    teller = 0
    txtXML(0).Text = vbNullString
    startTime = Now()
    
    For teller = 1 To 100
        txtXML(0).Text = cts.XMLGetRootContactsFast()
    Next teller
    Debug.Print (DatePart("s", Now() - startTime, 2, 2))
    
End Sub

Private Sub btnTest_Click()
    'Call test_Address
    Call test_Contact
    'Call test_RootContacts
    'Call test_GetChildren
End Sub

Private Sub test_Contact()
    Dim cts As Integration.SuperOffice
    Dim result As String
    Dim rsResult As ADODB.Recordset
    Set cts = New Integration.SuperOffice
        
    Set rsResult = cts.GetPersonSnapshotById(CLng(9808))
    result = ""
    If (Not rsResult.EOF) Then
        While Not rsResult.EOF
            result = result & rsResult.Fields("firstname").Value & ", " & rsResult.Fields("lastname").Value & vbCrLf
            rsResult.MoveNext
        Wend
    End If
    rsResult.Close
    Set rsResult = Nothing
    
    MsgBox (result)
    
End Sub

Private Sub test_Address()
    Dim cts As Integration.SuperOffice
    Dim result As String
    Dim rsResult As ADODB.Recordset
    Set cts = New Integration.SuperOffice
        
    result = cts.HTMLGetPersonsForContactAsDropDown(2, 2, "dbxKontakt", "Velg en", "0", "", False)
    MsgBox (result)
    
    Set rsResult = cts.GetAddressByContactId(3, PostAddress)
    result = ""
    If (Not rsResult.EOF) Then
        While Not rsResult.EOF
            result = result & rsResult.Fields("Address1").Value & ", " & rsResult.Fields("city").Value & ", " & rsResult.Fields("zipcode").Value & vbCrLf
            rsResult.MoveNext
        Wend
    End If
    rsResult.Close
    Set rsResult = Nothing
    MsgBox (result)
End Sub


Private Sub test_RootContacts()
    Dim cts As Integration.SuperOffice
    Dim result As String
    Dim rsResult As ADODB.Recordset
    Set cts = New Integration.SuperOffice
        
    Set rsResult = cts.GetAllRootContacts()
    result = ""
    If (Not rsResult.EOF) Then
        While Not rsResult.EOF
            result = result & rsResult.Fields("source_record").Value & ", " & rsResult.Fields("name").Value & vbCrLf
            rsResult.MoveNext
        Wend
    End If
    rsResult.Close
    Set rsResult = Nothing
    MsgBox (result)
End Sub

Private Sub test_GetChildren()
    Dim cts As Integration.SuperOffice
    Dim result As String
    Dim rsResult As ADODB.Recordset
    Set cts = New Integration.SuperOffice
        
    Set rsResult = cts.GetChildrenForContact(9)
    result = ""
    If (Not rsResult.EOF) Then
        While Not rsResult.EOF
            result = result & rsResult.Fields("destination_record").Value & ", " & rsResult.Fields("name").Value & vbCrLf
            rsResult.MoveNext
        Wend
    End If
    rsResult.Close
    Set rsResult = Nothing
    MsgBox (result)
End Sub


