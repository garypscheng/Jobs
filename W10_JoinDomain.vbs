On Error Resume Next

Const ForReading = 1

Set WshShell = CreateObject("WScript.Shell")
DataPath    = WshShell.CurrentDirectory

strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objComputer in colSettings 
    strComputer = objComputer.Name
    strDomain = objComputer.Domain
    strDomainRole = objComputer.DomainRole
    strWorkgroup = objComputer.Workgroup
Next
If strDomainRole = "1" Then
   'WScript.Quit 0
end if

Set colChassis = objWMIService.ExecQuery ("Select * from Win32_SystemEnclosure")
For Each objChassis in colChassis
    For  Each strChassisType in objChassis.ChassisTypes
         Select Case strChassisType
                Case 3 HardwareType="Desktop"
                Case 4 HardwareType="Desktop"
                Case 5 HardwareType="Desktop"
                Case 6 HardwareType="Desktop"
                Case 7 HardwareType="Desktop"
                Case 13 HardwareType="Desktop"
                Case 15 HardwareType="Desktop"
                Case 16 HardwareType="Desktop"
                Case 8 HardwareType="NoteBook"  '#Portable
                Case 9 HardwareType="NoteBook"  '#Laptop      
                Case 10 HardwareType="NoteBook" '#Notebook
                Case 11 HardwareType="NoteBook" '#Hand Held
                Case 12 HardwareType="NoteBook" '#Docking Station
                Case 14 HardwareType="NoteBook" '#Sub Notebook
                Case 18 HardwareType="NoteBook" '#Expansion Chassis
                Case 21 HardwareType="NoteBook" '#Peripheral Chassis
                Case Else HardwareType="VDI"
            End Select
    Next
Next
strHardware = "OU=" & HardwareType & ","

Const HKEY_LOCAL_MACHINE = &H80000002
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
strValueName = "ProductName"
objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
If not IsNull(strValue) Then
   LTSBStr = "LTSB"
   SACStr = "SAC"
   Win7 = "Windows 7"
   Win10 = "Windows 10"
   If InStr(strValue, Win7) > 0 Then
       WinVer = "Win7"
   End If
   If InStr(strValue, Win10) > 0 Then
       WinVer = "Win10"
   End If
   If InStr(strValue, LTSBStr) > 0 Then
       WinEdition = "OU=" & WinVer & " LTSC,"
   End If
   If InStr(strValue, SACStr) > 0 Then
       WinEdition = "OU=" & WinVer & " LTSC,"
   End If
else
   WinEdition = ""
End If
strHKJCKeyPath = "SOFTWARE\HKJC\Windows"
strValueName2 = "DisableUSB"
objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strHKJCKeyPath,strValueName2,strValue2
If IsNull(strValue2) Then
   strUSB = ""
ELSE
   if strValue2 = "Yes" Then
      strUSB = "OU=Disable USB,"
   else
      strUSB = ""
   end if
End If

strValueName3 = "SharePC"
objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strHKJCKeyPath,strValueName3,strValue3
If IsNull(strValue3) Then
   strSharePC = ""
ELSE
   if strValue3 = "Yes" Then
      strUSB = ""
      strSharePC = "OU=Share,"
   else
      strSharePC = ""
   end if

End If

strValueName4 = "RISSSO"
objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strHKJCKeyPath,strValueName4,strValue4
If IsNull(strValue4) Then
   strRacing = ""
ELSE
   if strValue4 = "Yes" Then
      strRacing = "OU=RIS SSO,"
      strFirewall = ""
      strOpenAccess = ""
      strUSB = ""      
      strSharePC = ""
   else
      strRacing = ""
   end if
End If

strValueName5 = "AllowModifyFW"
objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strHKJCKeyPath,strValueName5,strValue5
If IsNull(strValue5) Then
   strFirewall = ""
ELSE
   if strValue5 = "Yes" Then
      strFirewall = "OU=Allow modify FW,"
      strRacing = ""
      strOpenAccess = ""
      strUSB = ""      
      strSharePC = ""
   else
      strFirewall = ""
   end if
End If

if strHardware = "VDI" then
      strUSB = ""
      strSharePC = ""
      strRacing = ""
      strFirewall = ""
      strOpenAccess = ""
End If

strValueName6 = "OpenAceessFW"
objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strHKJCKeyPath,strValueName6,strValue6
If IsNull(strValue6) Then
   strOpenAccess = ""
ELSE
   if strValue5 = "Yes" Then
      strOpenAccess = "OU=Open access FW,"
      strFirewall = ""
      strHardware = ""
      strRacing = ""
      strUSB = ""      
      strSharePC = ""
   else
      strOpenAccess = ""
   end if
End If

strValueName6 = "Location"
objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strHKJCKeyPath,strValueName6,strValue6
If IsNull(strValue6) Then
   strLocation = ""
ELSE
   strLocation = "OU=HK,"
   If InStr(strValue6, "BJOffice") > 0 Then
       strLocation = "OU=CN-SZ,"
   End If
   If InStr(strValue6, "CTC") > 0 Then
       strLocation = "OU=CN-CTC,"
   End If
   If InStr(strValue6, "SZOffice") > 0 Then
       strLocation = "OU=CN-SZ,"
   End If
End If

strValueName7 = "OU"
objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strHKJCKeyPath,strValueName7,strValue7
If IsNull(strValue7) Then
   strPredefine = ""
ELSE
   if strValue7 <> "" then
      arr = Split(strValue7, ";")
      For i=0 to ubound(arr)
          strarrOU = "OU=" & arr(i) & ","
          if strPredefine = "" then
             strPredefine = strarrOU
          else
             strPredefine = strarrOU & strPredefine
          end if
      next
      strUSB = ""
      strSharePC = ""
      strHardware = ""
      strRacing = ""
      strFirewall = ""
      strOpenAccess = ""
      WinEdition = ""
   end if
End If

WebFileName = "W10_Web.txt"
WebFile = DataPath & "\" & WebFileName
Set objFSO = CreateObject("Scripting.FileSystemObject")
If (objFSO.FileExists(WebFile)) Then
   Set objFile = objFSO.OpenTextFile(WebFile, ForReading)
   Do While objFile.AtEndOfStream <> True
      strline = objFile.ReadLine
      if not strline = "" then
         strURLSrv = strline
      end if
   Loop
   objFile.Close()
End if

if strURLSrv = "" then
   strURLSrv = "http://10.172.25.200/"
end if

KeyFileName = "W10_Key.txt"
KeyFile = DataPath & "\" & KeyFileName
KeyURL = strURLSrv & KeyFileName
strPassword = ""
strUser = ""
HTTPDownload KeyURL, DataPath, KeyFileName
If (objFSO.FileExists(KeyFile)) Then
   Set objFile = objFSO.OpenTextFile(KeyFile, ForReading)
   Do While objFile.AtEndOfStream <> True
      strline = objFile.ReadLine
      if not strline = "" then
         strPKey = strline
      end if
   Loop
   objFile.Close()
   'objFSO.DeleteFile(KeyFile)
End if
wscript.echo "Get Key"
strDomain = "corp" ' Domain to logon
strDomainFQDN = strDomain & ".hkjc.com" ' Domain to logon
strServer = "10.140.1.1"


If strPKey <> "" then
   For i=1 To (Len(strPKey)/3)
       strAsc = chr(StrReverse(Mid(strPKey,i+2*(i-1),3))/7)
       strEKey = strEKey & strAsc
   Next 
   arr = Split(strEKey, ",")
   if ubound(arr) = 1 then
      strAccount = arr(0)
      strPassword = arr(1)
      arruser = Split(strAccount, "\")
      if ubound(arruser) = 1 then
         struser = arruser(1)
      end if
   end if
End If
wscript.echo "Decrypted Key"
strOU = strUSB & strSharePC & strHardware & strRacing & strFirewall & strOpenAccess & strPredefine & WinEdition & strLocation & "OU=Workstation," & strDevOU & "OU=HKJC,DC=" & strDomain & ",DC=hkjc,DC=com" ' OU to place computer in
strDefaultOU = "CN=Computers,DC=" & strDomain & ",DC=hkjc,DC=com" ' default computer object

Const ADS_SCOPE_SUBTREE = 2
Const ADS_SECURE_AUTHENTICATION = 1
Const JOIN_DOMAIN = 1
Const ACCT_CREATE = 2
Const ACCT_DELETE = 4
Const WIN9X_UPGRADE = 16
Const DOMAIN_JOIN_IF_JOINED = 32
Const JOIN_UNSECURE = 64
Const MACHINE_PASSWORD_PASSED = 128
Const DEFERRED_SPN_SET = 256
Const INSTALL_INVOCATION = 262144

Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject" 
objConnection.Properties("User ID")=strUser
objConnection.Properties("Password")=strPassword
objConnection.Properties("Encrypt Password")=TRUE
objConnection.Properties("ADSI Flag")=1
objConnection.Open "Active Directory Provider" 
Set objCommand = CreateObject("ADODB.Command") 
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE

objCommand.CommandText = "SELECT Name FROM 'LDAP://" & strServer & "/" & strOU & "' WHERE objectCategory='organizationalUnit'"
Set objRecordSet = objCommand.Execute
If (IsNull(objRecordSet.eof)) then
   wscript.echo "No such defined OU"
   strjoinOU = strDefaultOU
else
   if objRecordSet.RecordCount > 0 then
      strjoinOU = strOU
      wscript.echo "Defined OU found" & strOU
   else
      wscript.echo "Defined OU not found"
      strjoinOU = strDefaultOU
   end if
end if

Set objNetwork = CreateObject("WScript.Network")
strComputer = objNetwork.ComputerName
strWin10OU = "OU=Workstation,OU=HKJC,DC=corp,DC=hkjc,DC=com"
objCommand.CommandText = "SELECT ADsPath FROM 'LDAP://" & strServer & "/" & strWin10OU & "' Where objectClass='computer' and Name='" & strComputer & "'"
Set objRecordSet = objCommand.Execute
If not (IsNull(objRecordSet.eof)) then
   wscript.echo "there is computer object in the Windows 10 OU"
   if objRecordSet.RecordCount = 1 then
      wscript.echo "exist computer object found in the Windows 10 OU"
      Do Until objRecordSet.EOF
         objComputer = objRecordSet.Fields("ADsPath").Value
         Set objNS = GetObject("LDAP:")
         wscript.echo "the computer object name in the Windows 10 OU" & objComputer
         Set objRemoveComputer =  objNS.OpenDSObject(objComputer,strUser,strPassword,ADS_SECURE_AUTHENTICATION)
         objRemoveComputer.DeleteObject (0)
         objRecordSet.MoveNext
	 wscript.sleep 60000
      Loop
   end if
end if

objCommand.CommandText = "SELECT ADsPath FROM 'LDAP://" & strServer & "/" & strDefaultOU & "'Where objectClass='computer'" & " and Name='" & strComputer & "'"
Set objRecordSet = objCommand.Execute 
If not (IsNull(objRecordSet.eof)) then
   wscript.echo "there is computer object in the default 10 OU"
   if objRecordSet.RecordCount = 1 then
      wscript.echo "exist computer object found in the default 10 OU"
      Do Until objRecordSet.EOF
         objComputer = objRecordSet.Fields("ADsPath").Value
         Set objNS = GetObject("LDAP:")
         wscript.echo "the computer object name in the default 10 OU" & objComputer
         Set objRemoveComputer =  objNS.OpenDSObject(objComputer,strUser,strPassword,ADS_SECURE_AUTHENTICATION)
         objRemoveComputer.DeleteObject (0)
         objRecordSet.MoveNext
	 wscript.sleep 60000
      Loop
   end if
end if

if ((strjoinOU <> "") and (strPassword <> "")) then
   wscript.echo "Join OU"
   Set objComputer = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer & "\root\cimv2:Win32_ComputerSystem.Name='" & strComputer & "'")
   ReturnValue = objComputer.JoinDomainOrWorkGroup(strDomainFQDN, strPassword, strDomainFQDN & "\" & strUser, strjoinOU, JOIN_DOMAIN + ACCT_CREATE + DOMAIN_JOIN_IF_JOINED)
   Select Case ReturnValue
      Case 0 Status = "Success"
      Case 2 Status = "Missing OU"
      Case 5 Status = "Access denied"
      Case 53 Status = "Network path not found"
      Case 87 Status = "Parameter incorrect"
      Case 1326 Status = "Logon failure, user or pass"
      Case 1355 Status = "Domain can not be contacted"
      Case 1909 Status = "User account locked out"
      Case 2224 Status = "Computer Account allready exists"
      Case 2691 Status = "Allready joined"
      Case Else Status = "UNKNOWN ERROR " & ReturnValue
   End Select
   if ReturnValue <> "0" then
      Msgbox ("it cannot join to domain and the status is: " & Status)
      WScript.Quit 1
   End if
else
   if strPassword = "" then
      Msgbox ("The network connection failed, please check and then re-run step5.bat again." )
      WScript.Quit 8787
   End if
end if

Sub HTTPDownload( myURL, myPath, myFileName )
    Dim i, objFile, objFSO, objHTTP, strFile, strMsg

    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    myFile = myPath & "\" & myFileName
    If objFSO.FolderExists( myFile ) Then
        strFile = objFSO.BuildPath( myFileName )
    ElseIf objFSO.FolderExists( myPath ) Then
        strFile = myFile
    ElseIf not objFSO.FolderExists(myPath ) Then
        objFSO.CreateFolder myPath
        strFile = myFile
    else
        Exit Sub
    End If

    Set objFile = objFSO.OpenTextFile( strFile, ForWriting, True )
    Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )
    objHTTP.Open "GET", myURL, False
    objHTTP.Send

    For i = 1 To LenB( objHTTP.ResponseBody )
        objFile.Write Chr( AscB( MidB( objHTTP.ResponseBody, i, 1 ) ) )
    Next

    objFile.Close( )
End Sub