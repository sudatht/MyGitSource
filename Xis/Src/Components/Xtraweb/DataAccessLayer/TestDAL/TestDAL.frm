VERSION 5.00
Begin VB.Form FrmDALTest 
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "FetchRCSP"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FetchData"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "FrmDALTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim ObjDA As IdataAccess
    Dim objRC As Recordset
    Dim StrCon As String
   
        
    On Error GoTo test
    Set ObjDA = New ClsDataAccess
    Set objRC = ObjDA.FetchRC("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=ebusiness;password=ebusiness;Initial Catalog=140518;Data Source=hurum;Connect Timeout=15", "select * from units ")
    Set ObjDA = Nothing
    List1.Clear
    While Not objRC.EOF
        List1.AddItem objRC.Fields(1).Name & " : " & objRC.Fields(1).Value
        List1.AddItem objRC.Fields(3).Name & " : " & objRC.Fields(3).Value
        List1.AddItem objRC.Fields(4).Name & " : " & objRC.Fields(4).Value
        objRC.MoveNext
    Wend
    Label1.Caption = objRC.RecordCount
    Set objRC = Nothing
    Exit Sub
test:
    MsgBox Err.Number & ": " & Err.Description
End Sub

 
Private Sub Command2_Click()
    Dim ObjDA As IdataAccess
    Dim objRC As Recordset
    Dim StrCon As String
    Dim colparams As ADODB.Command
    Dim Objparam As ADODB.Parameter
       
    On Error GoTo test
    Set colparams = New Command
    Set Objparam = New ADODB.Parameter
    With Objparam
        .Name = "UNITID"
        .Direction = adParamInput
        .Type = adInteger
        colparams.Parameters.Append Objparam
    End With
    
    Set ObjDA = New ClsDataAccess
    Set objRC = ObjDA.FetchRCSP("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=ebusiness;password=ebusiness;Initial Catalog=EF;Data Source=praha;Connect Timeout=15", "fetchuser", colparams.Parameters)
    Set ObjDA = Nothing
    List1.Clear
    While Not objRC.EOF
        List1.AddItem objRC.Fields(1).Name & " : " & objRC.Fields(1).Value
        List1.AddItem objRC.Fields(3).Name & " : " & objRC.Fields(3).Value
        List1.AddItem objRC.Fields(4).Name & " : " & objRC.Fields(4).Value
        objRC.MoveNext
    Wend
    Label1.Caption = objRC.RecordCount
    Set objRC = Nothing
    Exit Sub
test:
    MsgBox Err.Number & ": " & Err.Description
End Sub
 

