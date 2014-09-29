VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Solution Name     : Xtra Web Integration
' Project Name      : Localizer
' Coding Standard   : EC Coding Standards
' Author            : STH
' Created Timestamp : 24/08/2006 10:57  AM
' --------------------------------------------------------------------------
Option Explicit


Private Const LANGUAGE_PARAM As String = "Language"
Private Const RESOURCE_DIR As String = "App_LocalResources"



Private m_objContext As COMSVCSLib.ObjectContext
Private m_objResponse As ASPTypeLibrary.Response
Private m_objSession As ASPTypeLibrary.Session
Private m_objServer As ASPTypeLibrary.Server
Private m_objRequest As ASPTypeLibrary.Request
Private m_objApplication As ASPTypeLibrary.Application
Private m_scriptName As String
Private m_scriptPath As String

Private m_objDOM As MSXML2.DOMDocument

Private Sub Class_Initialize()
    Set m_objContext = GetObjectContext()
    With m_objContext
        Set m_objResponse = .Item("Response")
        Set m_objSession = .Item("Session")
        Set m_objServer = .Item("Server")
        Set m_objRequest = .Item("Request")
        Set m_objApplication = .Item("Application")
    End With
    m_scriptName = m_objRequest.ServerVariables("SCRIPT_NAME")
    m_scriptPath = "./"
    LoadResources
End Sub

Private Sub Class_Terminate()
    Set m_objSession = Nothing
    Set m_objResponse = Nothing
    Set m_objContext = Nothing
    Set m_objRequest = Nothing
    Set m_objApplication = Nothing
    m_scriptName = 0
    m_scriptPath = 0
End Sub

Private Sub LoadResources()
    Dim strLang As String
    Dim strAppLang As String
    Dim strFile As String
    Dim blnExists As Boolean
    Dim check As Boolean
    Dim objSPMMgr As COMSVCSLib.SharedPropertyGroupManager
    Dim objSPMGrp As COMSVCSLib.SharedPropertyGroup
    Dim objProp As COMSVCSLib.SharedProperty
    
    strFile = m_scriptName
    strFile = Mid(strFile, InStrRev(strFile, "/") + 1)
    strFile = Mid(strFile, 1, InStr(strFile, ".") - 1)
    strLang = m_objSession(LANGUAGE_PARAM) & ""
    strAppLang = m_objApplication(LANGUAGE_PARAM) & ""

    If strLang = "" Then
        If strAppLang = "" Then
            strFile = m_objServer.MapPath(m_scriptPath & RESOURCE_DIR & "/" & strFile & ".resx")
        Else
            strFile = m_objServer.MapPath(m_scriptPath & RESOURCE_DIR & "/" & strFile & "." & strAppLang & ".resx")
        End If
    Else
        strFile = m_objServer.MapPath(m_scriptPath & RESOURCE_DIR & "/" & strFile & "." & strLang & ".resx")
    End If
    
    Set m_objDOM = New MSXML2.DOMDocument
    
    Set objSPMMgr = New COMSVCSLib.SharedPropertyGroupManager
    Set objSPMGrp = objSPMMgr.CreatePropertyGroup(strFile, LockSetGet, _
    Process, blnExists)
    
    Set objProp = objSPMGrp.CreateProperty("ResourceXml", blnExists)
    If Not blnExists Then
        Debug.Print "Property not found. loading from " & strFile
        check = m_objDOM.Load(strFile)

    Else
        'check = m_objDOM.loadXML(CStr(objProp.Value))
        check = m_objDOM.Load(strFile)
        
    End If
End Sub

Public Function GetText(ByVal Tag As String) As String

    
    If Not (m_objDOM Is Nothing) Then
        If Not m_objDOM.selectSingleNode("/root/data[@name='" & Tag & "']/value") Is Nothing Then
            GetText = m_objDOM.selectSingleNode("/root/data[@name='" & Tag & "']/value").Text
        Else
            GetText = ""
            Err.Raise vbObjectError, , "Load tag fail"
        End If
    Else
        GetText = ""
        Err.Raise vbObjectError, , "Language file not loaded."
    End If

End Function

Public Sub WriteText(ByVal Tag As String)

    m_objResponse.Write (GetText(Tag))

End Sub

Public Sub BeginInclude(ByVal InFile As String, ByVal InPath As String)

    m_scriptName = InFile
    m_scriptPath = InPath
    LoadResources

End Sub

Public Sub EndInclude()

   m_scriptName = m_objRequest.ServerVariables("SCRIPT_NAME")
   m_scriptPath = "./"
   LoadResources
   
End Sub


