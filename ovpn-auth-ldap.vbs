'
' OpenVPN AD Authentication for Windows
'
' Nicholas Caito
' xenomorph@gmail.com
' http://xenomorph.net/
' https://github.com/BitingChaos/
'
' Original file copyright Jesse Jordan
' http://ovpn-auth-ldap.sourceforge.net/
'
'


'Enum Definitions for the ADS_AUTHENTICATION_ENUM used with the IADsOpenDSObject
'------------------------------------------
'*ADS_AUTHENTICATION_ENUM*
'------------------------------------------
Const ADS_SECURE_AUTHENTICATION   = &h01
Const ADS_USE_ENCRYPTION          = &h02
Const ADS_USE_SSL                 = &h02
Const ADS_READONLY_SERVER         = &h04
Const ADS_PROMPT_CREDENTIALS      = &h08
Const ADS_NO_AUTHENTICATION       = &h10
Const ADS_FAST_BIND               = &h20
Const ADS_USE_SIGNING             = &h40
Const ADS_USE_SEALING             = &h80
Const ADS_USE_DELEGATION          = &h100
Const ADS_SERVER_BIND             = &h200
Const ADS_NO_REFERRAL_CHASING     = &h400
Const ADS_AUTH_RESERVED           = &h80000000

'------------------------------------------
'*ADS_ERROR_CODE_ENUM*
'------------------------------------------
Const E_ADS_INVALID_DOMAIN_OBJECT = &h80005001
Const E_ADS_BAD_PATHNAME= &h80005000

'------------------------------------------
'*LOGEVENT_ENUM*
'------------------------------------------
Const SEVERITY_INFO = &H00
Const SEVERITY_ERROR = &H01
Const SEVERITY_WARNING = &h02

'------------------------------------------
'*GLOBALS*
'------------------------------------------
Dim strAppPath:strAppPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") ' This variable DOES have a trailing '\'

Dim strIniFile:strIniFile = "ovpn-auth-ldap.ini"
Dim strIniFileOrig:strIniFileOrig = "ovpn-auth-ldap.orig.ini"

Dim strLogFile:strLogFile = "ovpn-auth-ldap.log"
Dim strLogLevel:strLogLevel = "2"

Dim strLogText:strLogText = "No Error" 

Dim strRootLDAP, strADDomain, strADServer, strADGroupDN, strADUser, strADPass

Public Sub Main

	Dim retVal:retVal = False
	
	Call LoadSettings(strAppPath, strIniFile)
	retVal = ParseParams()
	retVal = IsValidUser()
	
	If retVal = False Then
		strLogText = "Invalid login for Username: '" & strADUser & "'. " & _
		"Ensure that the user is a member of the correct VPN Group!"
		WriteLog(SEVERITY_ERROR)
		WScript.Quit(1)
	Else
		strLogText = "Successful login for Username: " & strADUser
		WriteLog(SEVERITY_INFO)
	End If
	
End Sub

'Loads the settings from the provided path and Ini filename, strPath should include trailing '\'
Public Sub LoadSettings(strPath, strFile)

	Dim objFso, objIni
	Dim strFilePath:strFilePath = strPath & strFile	
		
	' --------------------------------------------------------
	' Load Settings from INI file
	Set objFso = CreateObject("Scripting.FileSystemObject")
	If objFso.FileExists(strFilePath) Then
		Set objIni = New Ini
		
		objIni.IniFileName = strFilePath 
	
		objIni.IniSection = "ActiveDirectory"
		
		objIni.IniKey = "ADServer"
		strADServer = objIni.Value
		If strADServer = "" Then
			strLogText = "INI Value 'ADServer' cannot be blank!"
			WriteLog(SEVERITY_ERROR)
			WScript.Quit(1)
		End If
		
		objIni.IniKey = "Domain"
		strADDomain = objIni.Value
		If strADDomain = "" Then
			strLogText = "INI Value 'Domain' cannot be blank!"
			WriteLog(SEVERITY_ERROR)
			WScript.Quit(1)
		End If
		
		objIni.IniKey = "DN"
		strRootLDAP = objIni.Value
		If strRootLDAP = "" Then
			strLogText = "INI Value 'DN' cannot be blank!"
			WriteLog(SEVERITY_ERROR)
			WScript.Quit(1)
		End If
		
		objIni.IniKey = "ADGroupDN"
		strADGroupDN = objIni.Value
		If strADGroupDN = "" Then
			strLogText = "INI Value 'ADGroup' cannot be blank!"
			WriteLog(SEVERITY_ERROR)
			WScript.Quit(1)
		End If
		
		objIni.IniSection = "Logging"
		
		objIni.IniKey = "LogLevel"
		strLogLevel = objIni.Value
		
		If strADDomain = "" Then
			strLogText = "INI Value 'LogLevel' not defined, defaulting to LogLevel 2. " & _
						 "Errors, Warnings and Authentication Status Messages will be logged."
			WriteLog(SEVERITY_WARNING)
		End If
	Else
			strLogText = "INI File '" & strFilePath & "' could not be opened!"
			WriteLog(SEVERITY_ERROR)
			WScript.Quit(1)	
	End If
	
End Sub

'Parse command line parameters
Function ParseParams()

	Dim intParamCount
	Dim retVal:retVal = False
	
	strLogText = "None"
	strADUser = ""
	strADPass = ""	

	Select Case WScript.Arguments.Count
		Case 0
			'Use OpenVPN Credentials
			Set WshShell = CreateObject("WScript.Shell")
	    	strADUser = WshShell.ExpandEnvironmentStrings("%USERNAME%")
	    	strADPass = WshShell.ExpandEnvironmentStrings("%PASSWORD%")
	    	If Not (WshShell Is Nothing) Then Set WshShell = Nothing
	    
	    Case 1
	    	'Only 1 Parameter provided, check for '/?' or '/help' or '--help' or '--?'
			If WScript.Arguments.Item(0) = "/?" Or _
			   WScript.Arguments.Item(0) = "/help" Or _
			   WScript.Arguments.Item(0) = "--help" Or _
			   WScript.Arguments.Item(0) = "--?" Then
			   Call PrintHelp()
			   WScript.Quit(1)
			End If
		
		Case 2
			'Test with suppplied credentials
			strADUser = WScript.Arguments.Item(0)
			strADPass = WScript.Arguments.Item(1)
			
		Case Else
			' too many parameters given, error
			strLogText = "Too many parameters passed."
			WriteLog(SEVERITY_ERROR)
			WScript.Quit(1)
	End Select

	'Now Check the validity of the params
	If strADUser = "" Or strADPass = "" Then
		If strADUser = "" Then
			strLogText = "No Username provided."
		ElseIf strADPass = "" Then
			strLogText = "No Password provided for User:" & vbNewLine & strADUser & "."
		ElseIf strADUser = "" And strADPass = "" Then
			strLogText = "No user credentials provided."
		End If					
		WriteLog(SEVERITY_ERROR)
		WScript.Quit(1)
	End If

	If strLogText = "None" Then
		retVal = True
	End If

	ParseParams = retVal

End Function


'Attempt user login using provided credentials,
'if user has valid login then check group membership.
Function IsValidUser()

	On Error Resume Next 
	Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D

	Dim arrMemberOf
	Dim strPath, strUser, strPassword, strMember
	Dim objDSO, objConnection, objCommand, objRecordSet, objRoot, objGroup
	Dim retVal:retVal = False
	
	strPath = "LDAP://cn=" & strADUser & "," & strRootLDAP 
	strUser = strADDomain & "\" & strADUser
	strPassword = strADPass
	
	Set objDSO = GetObject("LDAP:")

	Set objRoot = objDSO.OpenDSObject("LDAP://" & strADServer & "/" & _
                                      "RootDSE", strADDomain & "\" & strADUser, _
                                      strADPass, ADS_SERVER_BIND And _
                                      ADS_USE_ENCRYPTION)

	If Err Then
		strLogText = "An authentication error has occurred!"
		WriteLog(Err.Number)
		WScript.Quit(1)
	End If
	
	Set objGroup = GetObject("LDAP://" & strADGroupDN)
	objGroup.GetInfo()
 
	arrMemberOf = objGroup.GetEx("member")
 
	For Each strMember in arrMemberOf
		Set objMember = GetObject("LDAP://" & strMember)
		If (StrComp(strADUser, objMember.SamAccountName, vbTextCompare) = 0) Then ' fix issue with case-sensitive username 
			retVal = True
			Exit For
		End If
	Next
	
	IsValidUser = retVal
	    
End Function

Public Sub PrintHelp()

	WScript.Echo "ovpn-auth-ldap.vbs - Secure LDAP Authentiation for User Logins" & vbNewLine & _
				 "Built for OpenVPN 2.1-Up running on Windows Server/XP/Vista/7 x86/x64"
	WScript.Echo "---------------------------------------------------------------------"
	WScript.Echo "1.) Start by configuring the ovpn-auth-ldap.ini file " & vbNewLine & _
				 "    to match your network configuration." & vbNewLine
	WScript.Echo "2.) Add the following to your Server.ovpn configuration file: " & vbNewLine & _
				 vbTab & "script-security 3" & vbNewLine & _
				 vbTab & "auth-user-pass-verify ""C:/Windows/System32/cscript.exe C:/Progra~1/OpenVPN/config/ovpn-auth-ldap.vbs"" via-env" & vbNewLine
	WScript.Echo "3.) The script uses the username and password from the OpenVPN " & vbNewLine & _
				 "    environment variables to authenticate the user." & vbNewLine
	WScript.Echo "4.) Remember to check the ovpn-auth-ldap.log file or your Windows " & vbNewLine & _
				 "    Event Log if you are having problems." & vbNewLine
	WScript.Echo "---------------------------------------------------------------------"
	WScript.Echo "Valid Command Line Arguments:"
	WScript.Echo "[/?] [/help] [--help] [--?] - This Help Screen"
	WScript.Echo "[Username] [Password] - Test the script using command line arguments"
	WScript.Echo "---------------------------------------------------------------------"
	WScript.Echo "Script Returns '0' for success '1' for failure" & vbNewLine
	WScript.Quit(1)

End Sub

'Logs Errors to Event Log and Log File
Function WriteLog(byVal errNum)

	Dim objFso, objFolder
	Dim strFile, strFolder, strFileName, strPath, strErrTxt
		
	' --------------------------------------------------------
	' Set the folder and file name
	strFileName = "ovpn-auth-ldap.log"
	strFolder = strAppPath
	strPath = strAppPath & strFileName
			
	' --------------------------------------------------------
	' Section to create folder and hold file.
	Set objFso = CreateObject("Scripting.FileSystemObject")
	If objFso.FolderExists(strFolder) Then
		Set objFolder = objFSO.GetFolder(strFolder)
	Else
		Set objFolder = objFSO.CreateFolder(strFolder)
	End If
	
	If objFso.FileExists(strPath) Then
		Set strFile = objFso.OpenTextFile(strPath, 8)
	Else
		Set strFile = objFso.CreateTextFile(strPath, True)
	End If
	
	' --------------------------------------------------------
	' output messages to log
	Select Case errNum
		Case 0
		     strErrTxt = "OpenVPN AD Login - Successfully logged in Username: '" & strADUser & _
			 "' into Domain Name: '" & strADDomain & "'."
		     LogEvent strErrTxt, SEVERITY_INFO ' write to system Event Log
		     strFile.WriteLine FormatDateTime(now(),0) & " - LOGIN: " & strErrTxt ' write to Log file
		     
	    Case -2147217911, -2147023570
	         strErrTxt = "OpenVPN AD Warning - " & strLogText & " " & _
	         			 "Invalid credentials for Username: '" & strADUser & "' " & _
	                  	 "on attempted login to Domain Name: '" & strADDomain & "'."
	         LogEvent strErrTxt, SEVERITY_WARNING
	         strFile.WriteLine FormatDateTime(now(),0) & " - WARNING: " & strErrTxt
	         
	    Case -2147217865, -2147023541
	         strErrTxt = "OpenVPN AD Error - " & strLogText & " " & _
	         			 "Cannot find server or invalid LDAP path supplied." & " " & _
	         			 "LDAP Path: '" & strRootLDAP & "'."
	         LogEvent strErrTxt, SEVERITY_ERROR
	         strFile.WriteLine FormatDateTime(now(),0) & " - ERROR: " & strErrTxt
	         
		Case E_ADS_INVALID_DOMAIN_OBJECT
		     strErrTxt = "OpenVPN AD Error - " & strLogText & _
		    			 "The specified domain either does not exist or could not be contacted" & _
		    			 " | " & "Domain Name: '" & strADDomain & "'."
		     LogEvent strErrTxt, SEVERITY_ERROR
		     strFile.WriteLine FormatDateTime(now(),0) & " - ERROR: " & strErrTxt
		     
		Case SEVERITY_WARNING
	         strErrTxt = "OpenVPN AD Warning - " & strLogText & " | " & Err.Number & ", " & Err.Description 
	         LogEvent strErrTxt, SEVERITY_WARNING
	         strFile.WriteLine FormatDateTime(now(),0) & " - WARNING: " & strErrTxt
	         
		Case SEVERITY_ERROR ' we usually get this one when a user isn't in the VPN group
	         strErrTxt = "OpenVPN AD Error - " & strLogText ' & " | " & Err.Number ' & ", " & Err.Description 
	         LogEvent strErrTxt, SEVERITY_ERROR
	         strFile.WriteLine FormatDateTime(now(),0) & " - ERROR: " & strErrTxt
	         
		Case SEVERITY_INFO
	         strErrTxt = "OpenVPN AD Info - " & strLogText ' & " | " & Hex(Err.Number) & ", " & Err.Description 
	         LogEvent strErrTxt, SEVERITY_INFO
	         strFile.WriteLine FormatDateTime(now(),0) & " - INFO: " & strErrTxt
	         
		Case Else
	         strErrTxt = "OpenVPN AD Unknown - " & strLogText ' & " | " & Hex(Err.Number) & ", " & Err.Description 
	         LogEvent strErrTxt, SEVERITY_ERROR
	         strFile.WriteLine FormatDateTime(now(),0) & " - ERROR: " & strErrTxt
	End Select
	
	strFile.Close
	If Not (strFile Is Nothing) Then Set strFile = Nothing

End Function

'Write logs to Event Log Service
Public Sub LogEvent(EventDescription, EventType)
	
	On Error Resume Next
	
	Dim objShell
	Set objShell = Wscript.CreateObject("Wscript.Shell")
	objShell.LogEvent EventType, EventDescription
	If Not (objShell Is Nothing) Then Set objShell = Nothing
	
End Sub

' -----------------------------------------------------
' ***** A class to read INI files *********************
' -----------------------------------------------------
Class Ini

	Private fso, f
	Private x, A, s, i, j
		
	Public Filename 'Path to the ini File
	Public Section  '[section]
	Public Key  	'Key=Value
	Public Default  'Return it when an error occurs
	
	Private Sub Class_Initialize() ' Setup Initialize event.
	    Default = ""
	    Set fso = CreateObject("Scripting.FileSystemObject")
	End Sub

	Private Sub Class_Terminate() ' Setup Terminate event.
	    Set fso = Nothing
	End Sub
	
	Property Let IniFileName(sFileName)
		Filename = sFileName
	End Property
	
	Property Get IniFileName()
		IniFileName = FileName
	End Property
	
	Property Let IniSection(sSection)
		Section = sSection
	End Property
	
	Property Get IniSection()
		IniSection = Section
	End Property

	Property Let IniKey(sKey)
		Key = sKey
	End Property
		
	Property Get Content()
	    'All the file in a string
	    If fso.FileExists(Filename) Then
	        Set f = fso.OpenTextFile(Filename, 1)
	        Content = f.ReadAll
	        f.Close
	        Set f = Nothing
	    Else
	        Content = ""
	    End If
	End Property
	
	Property Let Content(sContent)
	    'Create a brand new ini file
	    Set f = fso.CreateTextFile(Filename, True)
	    f.Write sContent
	    f.Close
	    Set f = Nothing
	End Property
	
	Property Get ContentArray()
	    'All the file in an array of lines
	    ContentArray = Split(Content, vbCrLf, -1, 1)
	End Property
	
	Private Sub FindSection(ByRef StartLine, ByRef EndLine)
	    StartLine = -1
	    EndLine = -2
	    A = ContentArray
	    For x = 0 To UBound(A)
	        s = UCase(Trim(A(x)))
	        If s = "[" & UCase(Section) & "]" Then
	            StartLine = x
	        Else
	            If (Left(s, 1) = "[") And (Right(s, 1) = "]") Then
	                If StartLine >= 0 Then
	                    EndLine = x - 1
	                    If EndLine > 0 Then 'A Space before the next section ?
	                        If Trim(A(EndLine)) = "" Then EndLine = EndLine - 1
	                    End If
	                    Exit Sub
	                End If
	            End If
	        End If
	    Next
	    If (StartLine >= 0) And (EndLine < 0) Then EndLine = UBound(A)
	End Sub
	
	Property Get Value()
	    'Retrieve the value for the current key in the current section
	    FindSection i, j
	    A = ContentArray
	    Value = Default
	    'Search only in the good section
	    For x = i + 1 To j
	        s = Trim(A(x))
	        If UCase(Left(s, Len(Key))) = UCase(Key) Then
	            Select Case Mid(s, Len(Key) + 1, 1)
	            Case "="
	                Value = Trim(Mid(s, Len(Key) + 2))
	                Exit Property
	            Case " ", Chr(9)
	                x = InStr(Len(Key), s, "=")
	                Value = Trim(Mid(s, x + 1))
	                Exit Property
	            End Select
	        End If
	    Next
	End Property
	
	Property Let Value(sValue)
	    ' Write the value for a key in a section
	    FindSection i, j
	    If i < 0 Then 'Session doesn't exist
	        Content = Content & vbCrLf & "[" & Section & "]" & vbCrLf & Key & "=" & sValue
	    Else
	        A = ContentArray
	        f = -1
	        'Search for the key, either the key exists or not
	        For x = i + 1 To j
	            s = Trim(A(x))
	            If UCase(Left(s, Len(Key))) = UCase(Key) Then
	                Select Case Mid(s, Len(Key) + 1, 1)
	                Case " ", Chr(9), "="
	                    f = x 'Key found
	                    A(x) = Key & "=" & sValue
	                End Select
	            End If
	        Next
	        If f = -1 Then
	            'Not found, add it at the end of the section
	            ReDim Preserve A(UBound(A) + 1)
	            For x = UBound(A) To j + 2 Step -1
	                A(x) = A(x - 1)
	            Next
	            A(j + 1) = Key & "=" & sValue
	        End If
	        'Define the content
	        s = ""
	        For x = 0 To UBound(A)
	            s = s & A(x) & vbCrLf
	        Next
	        'Suppress the last CRLF
	        If Right(s, 2) = vbCrLf Then s = Left(s, Len(s) - 2)
	        Content = s 'Write it
	    End If
	End Property
	
End Class

'Start Script
Call Main