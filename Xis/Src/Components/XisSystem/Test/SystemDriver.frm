VERSION 5.00
Begin VB.Form SystemDriver 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Xis System Driver"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCreateDri 
      Caption         =   "CreateDir"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "SystemDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCreateDri_Click()
    Dim util As XisSystem.util
    
    Set util = New XisSystem.util
    'util.EnsurePathExists ("\\xposltest1\CVUpload\CVdok")
    util.EnsurePathExists (txtPath.Text)
    
End Sub
