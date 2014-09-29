VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim ObjLogOn As LogOn.ClsLogOn
    Dim StrRetur As String
    
    Set ObjLogOn = New ClsLogOn
    StrRetur = ObjLogOn.LogOn("Provider=SQLOLEDB.1;User ID=ebusiness;Password=ebusiness;Data Source=PRAHA;Initial Catalog=XtraImp", "siris911", "siris911")
    MsgBox (StrRetur)
    
End Sub

Private Sub Command2_Click()
    Dim Objpwd As LogOn.ClsPassword
    Dim StrRetur As String
    
    Set Objpwd = New LogOn.ClsPassword
    StrRetur = Objpwd.GeneratePassword(6)
    MsgBox (StrRetur)
    
End Sub

Private Sub Command3_Click()
    Dim Objpwd As LogOn.ClsPassword
    
    Set Objpwd = New LogOn.ClsPassword
    MsgBox (Objpwd.IsValidPassword("copy5cat65"))
End Sub
