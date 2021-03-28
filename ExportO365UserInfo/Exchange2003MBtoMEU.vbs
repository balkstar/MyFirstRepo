'Globals/Constants
Const ADS_PROPERTY_APPEND = 3
Dim UserDN
Dim remoteSMTPAddress
Dim remoteLegacyDN
Dim domainController
Dim csvMode
csvMode = FALSE
Dim csvFileName
Dim lastADLookupFailed
Class UserInfo
    public OnPremiseEmailAddress
    public CloudEmailAddress
    public CloudLegacyDN
    public LegacyDN
    public ProxyAddresses
    public Mail
    public MailboxGUID
    public DistinguishedName
    Public Sub Class_Initialize()
        Set ProxyAddresses = CreateObject("Scripting.Dictionary")
    End Sub
End Class
'Command Line Parameters
If WScript.Arguments.Count = 0 Then
    'No parameters passed
    WScript.Echo("No parameters were passed.")
    ShowHelp()
ElseIf StrComp(WScript.Arguments(0), "-c", vbTextCompare) = 0 And WScript.Arguments.Count = 2 Then
    WScript.Echo("Missing DC Name.")
    ShowHelp()
ElseIf StrComp(WScript.Arguments(0), "-c", vbTextCompare) = 0 Then
    'CSV Mode
    csvFileName = WScript.Arguments(1)
    domainController = WScript.Arguments(2)
    csvMode = TRUE
    WScript.Echo("CSV mode detected. Filename: " & WScript.Arguments(1) & vbCrLf)
ElseIf wscript.Arguments.Count <> 4 Then
    'Invalid Arguments
    WScript.Echo WScript.Arguments.Count
    Call ShowHelp()
Else
    'Manual Mode
    UserDN = wscript.Arguments(0)
    remoteSMTPAddress = wscript.Arguments(1)
    remoteLegacyDN = wscript.Arguments(2)
    domainController = wscript.Arguments(3)
End If
Main()
'Main entry point
Sub Main
    'Check for CSV Mode
    If csvMode = TRUE Then
        UserInfoArray = GetUserInfoFromCSVFile()
    Else
        WScript.Echo "Manual Mode Detected" & vbCrLf
        Set info = New UserInfo
        info.CloudEmailAddress = remoteSMTPAddress
        info.DistinguishedName = UserDN
        info.CloudLegacyDN = remoteLegacyDN
        ProcessSingleUser(info)
    End If
End Sub
'Process a single user (manual mode)
Sub ProcessSingleUser(ByRef UserInfo)
    userADSIPath = "LDAP://" & domainController & "/" & UserInfo.DistinguishedName
    WScript.Echo "Processing user " & userADSIPath
    Set MyUser = GetObject(userADSIPath)
    proxyCounter = 1
    For Each address in MyUser.Get("proxyAddresses")
        UserInfo.ProxyAddresses.Add proxyCounter, address
        proxyCounter = proxyCounter + 1
    Next
    UserInfo.OnPremiseEmailAddress = GetPrimarySMTPAddress(UserInfo.ProxyAddresses)
    UserInfo.Mail = MyUser.Get("mail")
    UserInfo.MailboxGUID = MyUser.Get("msExchMailboxGUID")
    UserInfo.LegacyDN = MyUser.Get("legacyExchangeDN")
    ProcessMailbox(UserInfo)
End Sub
'Populate user info from CSV data
Function GetUserInfoFromCSVFile()
    CSVInfo = ReadCSVFile()
    For i = 0 To (UBound(CSVInfo)-1)
        lastADLookupFailed = false
        Set info = New UserInfo
        info.CloudLegacyDN = Split(CSVInfo(i+1), ",")(0)
        info.CloudEmailAddress = Split(CSVInfo(i+1), ",")(1)
        info.OnPremiseEmailAddress = Split(CSVInfo(i+1), ",")(2)
        WScript.Echo "Processing user " & info.OnPremiseEmailAddress
        WScript.Echo "Calling LookupADInformationFromSMTPAddress"
        LookupADInformationFromSMTPAddress(info)
        If lastADLookupFailed = false Then
            WScript.Echo "Calling ProcessMailbox"
            ProcessMailbox(info)
        End If
        set info = nothing
    Next
End Function
'Populate user info from AD
Sub LookupADInformationFromSMTPAddress(ByRef info)
    'Lookup the rest of the info in AD using the SMTP address
    Set objRootDSE = GetObject("LDAP://RootDSE")
    strDomain = objRootDSE.Get("DefaultNamingContext")
    Set objRootDSE = nothing
    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    Set objCommand = CreateObject("ADODB.Command")
    BaseDN = "<LDAP://" & domainController & "/" & strDomain & ">"
    adFilter = "(&(proxyAddresses=SMTP:" & info.OnPremiseEmailAddress & "))"
    Attributes = "distinguishedName,msExchMailboxGUID,mail,proxyAddresses,legacyExchangeDN"
    Query = BaseDN & ";" & adFilter & ";" & Attributes & ";subtree"
    objCommand.CommandText = Query
    Set objCommand.ActiveConnection = objConnection
    On Error Resume Next
    Set objRecordSet = objCommand.Execute
    'Handle any errors that result from the query
    If Err.Number <> 0 Then
        WScript.Echo "Error encountered on query " & Query & ". Skipping user."
        lastADLookupFailed = true
        return
    End If
    'Handle zero or ambiguous search results
    If objRecordSet.RecordCount = 0 Then
        WScript.Echo "No users found for address " & info.OnPremiseEmailAddress
        lastADLookupFailed = true
        return
    ElseIf objRecordSet.RecordCount > 1 Then
        WScript.Echo "Ambiguous search results for email address " & info.OnPremiseEmailAddress
        lastADLookupFailed = true
        return
    ElseIf Not objRecordSet.EOF Then
        info.LegacyDN = objRecordSet.Fields("legacyExchangeDN").Value
        info.Mail = objRecordSet.Fields("mail").Value
        info.MailboxGUID = objRecordSet.Fields("msExchMailboxGUID").Value
        proxyCounter = 1
        For Each address in objRecordSet.Fields("proxyAddresses").Value
            info.ProxyAddresses.Add proxyCounter, address
            proxyCounter = proxyCounter + 1
        Next
        info.DistinguishedName = objRecordSet.Fields("distinguishedName").Value
        objRecordSet.MoveNext
    End If
    objConnection = nothing
    objCommand = nothing
    objRecordSet = nothing
    On Error Goto 0
End Sub
'Populate data from the CSV file
Function ReadCSVFile()
    'Open file
    Set objFS = CreateObject("Scripting.FileSystemObject")
    Set objTextFile = objFS.OpenTextFile(csvFileName, 1, false, -1)
    'Loop through each line, putting each line of the CSV file into an array to be returned to the caller
    counter = 0
    Dim CSVArray()
    Do While NOT objTextFile.AtEndOfStream
        ReDim Preserve CSVArray(counter)
        CSVArray(counter) = objTextFile.ReadLine
        counter = counter + 1
    Loop
    'Close and return
    objTextFile.Close
    Set objTextFile = nothing
    Set objFS = nothing
    ReadCSVFile = CSVArray
End Function
'Process the migration
Sub ProcessMailbox(User)
    'Get user properties
    userADSIPath = "LDAP://" & domainController & "/" & User.DistinguishedName
    Set MyUser = GetObject(userADSIPath)
    'Add x.500 address to list of existing proxies
    existingLegDnFound = FALSE
    newLegDnFound = FALSE
    'Loop through each address in User.ProxyAddresses
    For i = 1 To User.ProxyAddresses.Count
        If StrComp(address, "x500:" & User.LegacyDN, vbTextCompare) = 0 Then
            WScript.Echo "x500 proxy " & User.LegacyDN & " already exists"
            existingLegDNFound = true
        End If
        If StrComp(address, "x500:" & User.CloudLegacyDN, vbTextCompare) = 0 Then
            WScript.Echo "x500 proxy " & User.CloudLegacyDN & " already exists"
            newLegDnFound = true
        End If
    Next
    'Add existing leg DN to proxy list
    If existingLegDnFound = FALSE Then
        WScript.Echo "Adding existing legacy DN " & User.LegacyDN & " to proxy addresses"
        User.ProxyAddresses.Add (User.ProxyAddresses.Count+1),("x500:" & User.LegacyDN)
    End If
    'Add new leg DN to proxy list
    If newLegDnFound = FALSE Then
        'Add new leg DN to proxy addresses
        WScript.Echo "Adding new legacy DN " & User.CloudLegacyDN & " to existing proxy addresses"
        User.ProxyAddresses.Add (User.ProxyAddresses.Count+1),("x500:" & User.CloudLegacyDN)
    End If
    'Dump out new list of addresses
    WScript.Echo "Original proxy addresses updated count: " & User.ProxyAddresses.Count
    For i = 1 to User.ProxyAddresses.Count
        WScript.Echo " proxyAddress " & i & ": " & User.ProxyAddresses(i)
    Next
    'Delete the Mailbox
    WScript.Echo "Opening " & userADSIPath & " as CDOEXM::IMailboxStore object"
    Set Mailbox = MyUser
    Wscript.Echo "Deleting Mailbox"
    On Error Resume Next
    Mailbox.DeleteMailbox
    'Handle any errors deleting the mailbox
    If Err.Number <> 0 Then
        WScript.Echo "Error " & Err.number & ". Skipping User." & vbCrLf & "Description: " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error Goto 0
    'Save and continue
    WScript.Echo "Saving Changes"
    MyUser.SetInfo
    WScript.Echo "Refeshing ADSI Cache"
    MyUser.GetInfo
    Set Mailbox = nothing
    'Mail Enable the User
    WScript.Echo "Opening " & userADSIPath & " as CDOEXM::IMailRecipient"
    Set MailUser = MyUser
    WScript.Echo "Mail Enabling user using targetAddress " & User.CloudEmailAddress
    MailUser.MailEnable User.CloudEmailAddress
    WScript.Echo "Disabling Recipient Update Service for user"
    MyUser.PutEx ADS_PROPERTY_APPEND, "msExchPoliciesExcluded", Array("{26491CFC-9E50-4857-861B-0CB8DF22B5D7}")
    WScript.Echo "Saving Changes"
    MyUser.SetInfo
    WScript.Echo "Refreshing ADSI Cache"
    MyUser.GetInfo
    'Add Legacy DN back on to the user
    WScript.Echo "Writing legacyExchangeDN as " & User.LegacyDN
    MyUser.Put "legacyExchangeDN", User.LegacyDN
    'Add old proxies list back on to the MEU
    WScript.Echo "Writing proxyAddresses back to the user"
    For j=1 To User.ProxyAddresses.Count
        MyUser.PutEx ADS_PROPERTY_APPEND, "proxyAddresses", Array(User.ProxyAddresses(j))
        MyUser.SetInfo
        MyUser.GetInfo
    Next
    'Add mail attribute back on to the MEU
    WScript.Echo "Writing mail attribute as " & User.Mail
    MyUser.Put "mail", User.Mail
    'Add msExchMailboxGUID back on to the MEU
    WScript.Echo "Converting mailbox GUID to writable format"
    Dim mbxGUIDByteArray
    Call ConvertHexStringToByteArray(OctetToHexString(User.MailboxGUID), mbxGUIDByteArray)
    WScript.Echo "Writing property msExchMailboxGUID to user object with value " & OctetToHexString(User.MailboxGUID)
    MyUser.Put "msExchMailboxGUID", mbxGUIDByteArray
    WScript.Echo "Saving Changes"
    MyUser.SetInfo
    WScript.Echo "Migration Complete!" & vbCrLf
End Sub
'Returns the primary SMTP address of a user
Function GetPrimarySMTPAddress(Addresses)
    For Each address in Addresses
        If Left(address, 4) = "SMTP" Then GetPrimarySMTPAddress = address
    Next
End Function
'Converts Hex string to byte array for writing to AD
Sub ConvertHexStringToByteArray(ByVal strHexString, ByRef pByteArray)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Stream = CreateObject("ADODB.Stream")
    Temp = FSO.GetTempName()
    Set TS = FSO.CreateTextFile(Temp)
    For i = 1 To (Len (strHexString) -1) Step 2
        TS.Write Chr("&h" & Mid (strHexString, i, 2))
    Next
    TS.Close
    Stream.Type = 1
    Stream.Open
    Stream.LoadFromFile Temp
    pByteArray = Stream.Read
    Stream.Close
    FSO.DeleteFile Temp
    Set Stream = nothing
    Set FSO = Nothing
End Sub
'Converts raw bytes from AD GUID to readable string
Function OctetToHexString (arrbytOctet)
    OctetToHexStr = ""
    For k = 1 To Lenb (arrbytOctet)
        OctetToHexString = OctetToHexString & Right("0" & Hex(Ascb(Midb(arrbytOctet, k, 1))), 2)
    Next
End Function
Sub ShowHelp()
    WScript.Echo("This script runs in two modes, CSV Mode and Manual Mode." & vbCrLf & "CSV Mode allows you to specify a CSV file from which to pull usernames." & vbCrLf& "Manual mode allows you to run the script against a single user.")
    WSCript.Echo("Both modes require you to specify the name of a DC to use in the local domain." & vbCrLf & "To run the script in CSV Mode, use the following syntax:")
    WScript.Echo("  cscript Exchange2003MBtoMEU.vbs -c x:\csv\csvfilename.csv dc.domain.com")
    WScript.Echo("To run the script in Manual Mode, you must specify the users AD Distinguished Name, Remote SMTP Address, Remote Legacy Exchange DN, and Domain Controller Name.")
    WSCript.Echo("  cscript Exchange2003MBtoMEU.vbs " & chr(34) & "CN=UserName,CN=Users,DC=domain,DC=com" & chr(34) & " " & chr(34) & "user@cloudaddress.com" & chr(34) & " " & chr(34) & "/o=Cloud Org/ou=Cloud Site/ou=Recipients/cn=CloudUser" & chr(34) & " dc.domain.com")
    WScript.Quit
End Sub