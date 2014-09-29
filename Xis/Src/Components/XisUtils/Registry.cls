' --------------------------------------------------------------------------
' Copyright(c) 2000-2006 Eurocenter DDC.
' No 65, Walukarama Road, Colombo 3, Sri Lanka
' All rights reserved.
'
' This software is the confidential and proprietary information of 
' Eurocenter DDC (Confidential Information). You shall not disclose such
' Confidential Information and shall use it only in accordance with the
' terms of the license agreement you entered into with Eurocenter.
'
' Solution Name     : Xis
' Project Name      : XisUtils
' Coding Standard   : EC Coding Standards
' Author            : STH
' Created Timestamp : 11/06/2006 10:57  AM
' --------------------------------------------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get values from register
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <param name=strRegKeyPath>Reg key path</param>
        ''' <param name=strRegSubKey>reg sub key</param>
        ''' <history>
        ''' </history>
        ''' -----------------------------------------------------------------------------
Public Function GetKeyValue(ByVal strRegKeyPath As String, ByVal strRegSubKey As String) As String

    Dim lnRootKey As Long
    Dim strResult As String
    lnRootKey = HKEY_LOCAL_MACHINE
    
    GetKeyValue = regQueryAKey(lnRootKey, strRegKeyPath, strRegSubKey)
 

End Function


