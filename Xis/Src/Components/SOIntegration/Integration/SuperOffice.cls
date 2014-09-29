VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SuperOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private cString As String

Private Const SUPEROFFICE_WEBUSER_CONNECTIONSTRING_KEYPATH = "SOFTWARE\Electric Farm\Integration\SuperOffice\Connectionstrings"
Private Const SUPEROFFICE_WEBUSER_CONNECTIONSTRING_VALUENAME = "WebConnectionString_oledb"

Public Enum AddressTypes
    PostAddress = 1
    StreetAddress = 2
    InvoiceAddress = 3
End Enum

Public Function GetAddressByContactId(ContactId As Long, addressType As AddressTypes) As ADODB.Recordset
    Dim strSQL As String
    Dim rsResult As ADODB.Recordset
    
    If addressType = AddressTypes.PostAddress Or addressType = AddressTypes.StreetAddress Then
        strSQL = "SELECT * FROM crm5.address WHERE owner_id = " & ContactId & " AND atype_idx = " & addressType
    End If
    If addressType = AddressTypes.InvoiceAddress Then
        strSQL = "SELECT string45 AS address1, string09 AS zipcode, string46 AS city from crm5.udcontactlarge, crm5.udcontactsmall, crm5.contact WHERE crm5.contact.userdef2_id = crm5.udcontactlarge.udcontactLarge_id AND crm5.contact.userdef_id = crm5.udcontactsmall.udcontactSmall_id AND crm5.contact.contact_id = " & ContactId
    End If

    Set rsResult = GetDisconnectedRS(strSQL)

    Set GetAddressByContactId = rsResult
End Function

Public Function GetPersonsSnapshotForContact(ContactId As Long, retrieveRetired As Boolean) As ADODB.Recordset
    Dim strSQL As String
    Dim retiredSQL As String: retiredSQL = ""
    Dim rsResult As ADODB.Recordset
    
    If (retrieveRetired = False) Then
        retiredSQL = " AND retired = 0"
    End If
    
    strSQL = "SELECT person_id, firstname, middlename, lastname, retired from crm5.person WHERE contact_id = " & ContactId & retiredSQL
    Set rsResult = GetDisconnectedRS(strSQL)
    
    Set GetPersonsSnapshotForContact = rsResult
End Function

Public Function GetPersonsForContact(ContactId As Long, retrieveRetired As Boolean) As ADODB.Recordset
    Dim strSQL As String
    Dim retiredSQL As String: retiredSQL = ""
    Dim rsResult As ADODB.Recordset
    
    If (retrieveRetired = False) Then
        retiredSQL = " AND retired = 0"
    End If
    
    strSQL = "SELECT * from crm5.person WHERE contact_id = " & ContactId & retiredSQL
    Set rsResult = GetDisconnectedRS(strSQL)
    
    Set GetPersonsForContact = rsResult
End Function

Public Function GetPersonSnapshotById(ByVal personId As Long) As ADODB.Recordset
    Dim strSQL As String
    Dim rsResult As ADODB.Recordset

    strSQL = "SELECT person_id, firstname, middlename, lastname, retired, dbo.fn_GetPersonEmails(person_id) AS [ContactEmail] from crm5.person WHERE Person_id = " & personId
    Set rsResult = GetDisconnectedRS(strSQL)

    Set GetPersonSnapshotById = rsResult
End Function


'Retrieves all children for a specific Contact, as a disconnected recordset
Public Function GetChildrenForContact(ContactId As Long) As ADODB.Recordset
    Dim strSQL As String
    Dim previousId
    Dim rsResult As ADODB.Recordset
        
    strSQL = "SELECT relations.destination_record AS Contact_id, contact.name " & vbNewLine & _
    " from crm5.relations, crm5.contact " & vbNewLine & _
    " WHERE crm5.relations.destination_record = crm5.contact.contact_id " & vbNewLine & _
    " AND crm5.relations.reldef_id = 1 " & vbNewLine & _
    " AND crm5.relations.source_table = 5 " & vbNewLine & _
    " AND crm5.relations.destination_table = 5 " & vbNewLine & _
    " AND crm5.relations.Source_record = " & ContactId & vbNewLine & _
    " ORDER BY crm5.contact.name ASC "

    Set rsResult = GetDisconnectedRS(strSQL)
    Set GetChildrenForContact = rsResult
End Function


'Retrieves all top level contacts (Contacts which have children but no parent),
'as a disconnected recordset.
Public Function GetAllRootContacts() As ADODB.Recordset
    Dim strSQL As String
    Dim previousId As Long
    Dim rsResult As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim excludeChildrenSQL As String
    Dim excludeChildrenHelper() As String
    Dim counter As Integer
    
    'Get all child records
    strSQL = " SELECT destination_record " & _
        "FROM crm5.relations " & _
        "WHERE " & _
        " reldef_id = 1 " & _
        " AND source_table = 5 " & _
        " AND destination_table = 5 " & _
        " ORDER BY destination_record DESC"
        
    'create "not in" statements
    Set rsTemp = GetDisconnectedRS(strSQL)
    If (Not rsTemp.EOF) Then
        previousId = 0
        ReDim excludeChildrenHelper(rsTemp.RecordCount)
        counter = 1
        While Not rsTemp.EOF
            If (CLng(previousId) <> CLng(rsTemp("destination_record").value)) Then
                excludeChildrenHelper(counter) = (" AND Source_record <> " & rsTemp("destination_record").value & vbNewLine)
                counter = counter + 1
                previousId = rsTemp("destination_record")
            End If
            rsTemp.MoveNext
        Wend
        ReDim Preserve excludeChildrenHelper(counter)
        excludeChildrenSQL = Join(excludeChildrenHelper)
    End If
    Set rsTemp = Nothing
    
    strSQL = "SELECT relations.source_record AS contact_id, contact.name " & vbNewLine & _
        "FROM crm5.relations, crm5.contact " & vbNewLine & _
        "WHERE " & vbNewLine & _
        " crm5.contact.contact_id = crm5.relations.source_record " & vbNewLine & _
        " AND reldef_id = 1 " & vbNewLine & _
        " AND source_table = 5 " & vbNewLine & _
        " AND destination_table = 5 " & vbNewLine & _
        excludeChildrenSQL & _
        " ORDER BY crm5.contact.name ASC"
    
    Set rsResult = GetDisconnectedRS(strSQL)
    'Remove duplicates (Superoffice SODBIF layer doesn't understand DISTINCT)
    If (Not rsResult.EOF) Then
        previousId = 0
        While (Not rsResult.EOF)
            If (CLng(previousId) = CLng(rsResult("contact_id").value)) Then
                rsResult.Delete (adAffectCurrent)
            Else
                previousId = rsResult("contact_id")
            End If
            rsResult.MoveNext
        Wend
        rsResult.MoveFirst
    End If
    
    Set GetAllRootContacts = rsResult
End Function

'Returns all the persons for the specified contact in a HTML dropdownbox.
'If one of the persons is selected pass the id as selectedId.
Public Function HTMLGetPersonsForContactAsDropDown(ContactId As Long, selectedId As Long, dropdownName As String, firstText As String, firstValue As String, additionalAttributes As String, retrieveRetired As Boolean) As String
    Dim result As String: result = ""
    Dim selected As String
    Dim contacts() As String
    Dim attributes() As String
    Dim name As String
    Dim rsResult As ADODB.Recordset
    Dim returnValue As String: returnValue = ""
    
    If Len(firstText) > 0 Then
        returnValue = vbTab & "<option value='" & firstValue & "'>" & firstText & "</option>" & vbCrLf
    End If
    
    Set rsResult = GetPersonsSnapshotForContact(ContactId, retrieveRetired)
        
    If (Not rsResult.EOF) Then
        result = rsResult.GetString(adClipString, -1, "|", ";")
        
        contacts = Split(result, ";")
        If UBound(contacts) > 0 Then
            Dim i As Long
            Dim x As Long
            For i = 0 To UBound(contacts) - 1
                attributes = Split(contacts(i), "|")
                If Len(attributes(2)) > 0 Then
                    name = attributes(1) & " " & attributes(2) & " " & attributes(3)
                Else
                    name = attributes(1) & " " & attributes(3)
                End If
                If (selectedId = CLng(attributes(0))) Then
                    selected = " selected "
                Else
                    selected = " "
                End If
                
                returnValue = returnValue & vbTab & "<option value='" & attributes(0) & "'" & selected & ">" & name & "</option>" & vbCrLf
            Next i
        End If
        
    End If
    rsResult.Close
    Set rsResult = Nothing
            
    HTMLGetPersonsForContactAsDropDown = "<select name='" & dropdownName & "'" & additionalAttributes & ">" & vbCrLf & returnValue & "</select>"
End Function

'Currently not in use
Public Function XMLGetRootContacts() As String
    Dim rsRootItems As ADODB.Recordset
    Dim xmlBuilder() As String
    Dim xmlResult As String
    Dim XmlBuilderGrowth As Long
    XmlBuilderGrowth = 15
    
    Set rsRootItems = Me.GetAllRootContacts()
    
    If (Not rsRootItems.EOF) Then
        xmlResult = xmlResult & "<ContactTree>"
        While (Not rsRootItems.EOF)
            xmlResult = xmlResult & "<ContactItem"
            xmlResult = xmlResult & " contact_id='"
            xmlResult = xmlResult & rsRootItems("contact_id").value
            xmlResult = xmlResult & "' Name='"
            xmlResult = xmlResult & rsRootItems("name").value
            xmlResult = xmlResult & "'>"
            xmlResult = xmlResult & XMLGetChildrenRecursiveForContact(rsRootItems("contact_id").value)
            xmlResult = xmlResult & "</ContactItem>"
            rsRootItems.MoveNext
        Wend
        xmlResult = xmlResult & "</ContactTree>"
    Else
        xmlResult = "<ContactTree/>"
    End If
    Set rsRootItems = Nothing
    XMLGetRootContacts = xmlResult
End Function

Public Function XMLGetRootContactsFast() As String
    Dim rsRootItems As ADODB.Recordset
    Dim util As Integration.FastString
    Dim xmlBuilder() As String
    Dim xmlResult As String
    Dim XmlBuilderGrowth As Long
    XmlBuilderGrowth = 15
    
    Set rsRootItems = Me.GetAllRootContacts()
    
    If (Not rsRootItems.EOF) Then
        Set util = New Integration.FastString
        util.Append "<ContactTree>"
        While (Not rsRootItems.EOF)
            util.Append "<ContactItem"
            util.Append " contact_id='"
            util.Append rsRootItems("contact_id").value
            util.Append "' Name='"
            util.Append rsRootItems("name").value
            util.Append "'>"
            util.Append XMLGetChildrenRecursiveForContactFast(rsRootItems("contact_id").value)
            util.Append "</ContactItem>"
            rsRootItems.MoveNext
        Wend
        util.Append "</ContactTree>"
        xmlResult = util.ToString()
        Set util = Nothing
    Else
        xmlResult = "<ContactTree/>"
    End If
    Set rsRootItems = Nothing
    XMLGetRootContactsFast = xmlResult
End Function

Private Function XMLGetChildrenRecursiveForContactFast(contact_id As Long) As String
    Dim rsSubItems As ADODB.Recordset
    Dim xmlResult As String
    Dim tmpResult As String
    Dim util As Integration.FastString
    
    Set rsSubItems = Me.GetChildrenForContact(contact_id)
    
    If (Not rsSubItems.EOF) Then
        Set util = New Integration.FastString
        util.Append "<ContactList>"
        While (Not rsSubItems.EOF)
            util.Append "<ContactItem"
            util.Append " contact_id='"
            util.Append rsSubItems("contact_id").value
            util.Append "' Name='"
            util.Append rsSubItems("name").value
            util.Append "'>"
            tmpResult = XMLGetChildrenRecursiveForContact(rsSubItems("contact_id").value)
            If (LenB(tmpResult) <> 0) Then
                util.Append XMLGetChildrenRecursiveForContactFast(rsSubItems("contact_id").value)
            End If
            util.Append "</ContactItem>"
            rsSubItems.MoveNext
        Wend
        util.Append "</ContactList>"
        xmlResult = util.ToString()
    Else
        xmlResult = vbNullString
    End If
    Set rsSubItems = Nothing
    XMLGetChildrenRecursiveForContactFast = xmlResult
End Function

Private Function XMLGetChildrenRecursiveForContact(contact_id As Long) As String
    Dim rsSubItems As ADODB.Recordset
    Dim xmlResult As String
    Dim tmpResult As String
    
    Set rsSubItems = Me.GetChildrenForContact(contact_id)
    
    If (Not rsSubItems.EOF) Then
        xmlResult = xmlResult & "<ContactList>"
        While (Not rsSubItems.EOF)
            xmlResult = xmlResult & "<ContactItem"
            xmlResult = xmlResult & " contact_id='"
            xmlResult = xmlResult & rsSubItems("contact_id").value
            xmlResult = xmlResult & "' Name='"
            xmlResult = xmlResult & rsSubItems("name").value
            xmlResult = xmlResult & "'>"
            tmpResult = XMLGetChildrenRecursiveForContact(rsSubItems("contact_id").value)
            If (LenB(tmpResult) <> 0) Then
                xmlResult = xmlResult & XMLGetChildrenRecursiveForContact(rsSubItems("contact_id").value)
            End If
            xmlResult = xmlResult & "</ContactItem>"
            rsSubItems.MoveNext
        Wend
        xmlResult = xmlResult & "</ContactList>"
    Else
        xmlResult = vbNullString
    End If
    Set rsSubItems = Nothing
    XMLGetChildrenRecursiveForContact = xmlResult
End Function

'Retrieves the web client superoffice connection string from registry, storing it in a private member and returning it to the callee.
'Subsequent calls will return will not access the registry but return it from the private member.
Private Function GetConnectionString() As String
    Dim conString As String
    If (cString = "") Then
        cString = GetRegistrySetting(SUPEROFFICE_WEBUSER_CONNECTIONSTRING_KEYPATH, SUPEROFFICE_WEBUSER_CONNECTIONSTRING_VALUENAME)
        'cString = "Provider=SODBP32.SOProvider.1;Data Source=CRM5;Location=ODBC:SuperOffice;User id=fredm:US;Password=;"
    End If
    GetConnectionString = cString
End Function

'Executes the specified SQL against the SuperOffice database, and retrieves the
'result as a disconnected recordset
Private Function GetDisconnectedRS(strSQL As String) As ADODB.Recordset
    Dim resultRs As New ADODB.Recordset
    Dim connection As New ADODB.connection
    
    connection.ConnectionString = GetConnectionString()
    connection.Open
    
    resultRs.CursorLocation = adUseClient
    resultRs.CursorType = adOpenStatic
    resultRs.LockType = adLockOptimistic

    resultRs.Open strSQL, connection
    
    Set resultRs.ActiveConnection = Nothing
    connection.Close

    Set GetDisconnectedRS = resultRs
    Set resultRs = Nothing
End Function

'Retrieves the specified Subkey and value from the registry starting at
'HKEY_LOCAL_MACHINE\SOFTWARE\Electric Farm.
'The subkey takes the form of "Integration\SuperOffice" and the value parameter
'must be a string value.
Private Function GetRegistrySetting(key As String, value As String) As String
    Dim regUtil As XisUtils.Registry
    
    Set regUtil = New XisUtils.Registry
    GetRegistrySetting = regUtil.GetKeyValue(key, value)
    Set regUtil = Nothing
End Function

