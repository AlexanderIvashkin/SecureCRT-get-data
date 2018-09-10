# $language = "VBScript"
# $interface = "1.0"

' Execute commands on Cisco devices and capture the output.
' Dec 2011, Alexander Ivashkin
' 16 Jan 2012: added Safeword authentication failure count with confirmation
' 21 Sep 2012: Added preliminary support of CatOS (not really working)
' 20 Mar 2013: New connection mode: telnet from syslog/tftp servers
' 03 April 2013: Major update. CSV export; failed connections list; passwords are separate from the device list and are reusable; minor bugfixes and flowchart updates
' 04 April 2013: Full support of static usernames/passwords. Optimised and reorganised into subroutines.
' 04 April 2013: Full support of CatOS
' 14 January 2014: bugfix: "set len 0" for CatOS
' 03 March 2014: Flushing passwords on every iteration just in case something goes wrong.
' 12 May 2017: MsgBox to signal completion; removal of blank line in the CSV file
' 13 September 2017: Minor improvements and major CSV output feature: additional column with the command executed to allow for easier filtering and comparison in Excel
' 15 September 2017: Preliminary code for the SSH support. GoTo does not exist. Rewrite of blocks into subroutines will be required.
' 18 September 2017: Bugfix: After encountering a single CatOS switch, "set len 0" became persistent and used on all the subsequent devices
' 20 September 2017: Version control is introduced. Deeming this the Stable Version 1.0.
' (no more version comments will be here, everything will be in Git)



Option Explicit
'==========================================================================
    ' These constants are to be tuned for your liking or needs
	Const safewordUsername = "alexander_ivashkin" ' Username to use for login
	Const staticUsername = "cisco" ' Login to use for locally managed devices
	Const staticPassword = "cisco" ' Password to use locally managed devices
	Const DEVICE_FILE_PATH = "devices.txt"
	Const OUTPUT_FILE_PATH = "output.txt"
	Const COMMANDS_FILE_PATH = "commands.txt"
	Const PASSWORDS_FILE_PATH = "passwords.txt"
	Const RESULTS_FILE_PATH = "devices_processed.csv"
	Const CSV_FILE_PATH = "output.csv"
	Const waitingTimeout = 30 ' How long to wait for a response from devices (in seconds)
'==========================================================================
' END OF CUSTOMISABLE PART.
' HERE BE DRAGONS!

	Const ForWriting = 2
	Const ForAppending = 8
	Const ForReading = 1
	
	' Global variables and objects
	Dim password
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	Dim fso_out
	Set fso_out = CreateObject("Scripting.FileSystemObject")
	Dim fso_cmd
	Set fso_cmd = CreateObject("Scripting.FileSystemObject")
	Dim fso_pswd
	Set fso_pswd = CreateObject("Scripting.FileSystemObject")
	Dim fso_results
	Set fso_results = CreateObject("Scripting.FileSystemObject")
	Dim fso_csv
	Set fso_csv = CreateObject("Scripting.FileSystemObject")
	
	Dim fil
	Set fil = fso.OpenTextFile(DEVICE_FILE_PATH)
	Dim fil_out
	Set fil_out = fso_out.OpenTextFile(OUTPUT_FILE_PATH, ForWriting, True)
	Dim fil_cmd
	Set fil_cmd = fso_cmd.OpenTextFile(COMMANDS_FILE_PATH)
	Dim fil_pswd
	Set fil_pswd = fso_cmd.OpenTextFile(PASSWORDS_FILE_PATH)
	Dim fil_results
	Set fil_results = fso_cmd.OpenTextFile(RESULTS_FILE_PATH, ForWriting, True)
	Dim fil_csv
	Set fil_csv = fso_cmd.OpenTextFile(CSV_FILE_PATH, ForWriting, True)

	Dim line
	Dim commands(100)
	Dim cmdIndex
	
	' Execute "terminal length 0" as the first command in all cases except when running CatOS (this is dealt with separately later)
	' And "show clock" as well.
	commands(0) = "ter len 0"
	commands(1) = "show clock"
	cmdIndex = 2
	' Read all the commands into an array
	While Not fil_cmd.AtEndOfStream
	    commands(cmdIndex) = fil_cmd.ReadLine
	    cmdIndex = cmdIndex + 1
	Wend

	' Read passwords
	Dim usedPasswords
	usedPasswords = 0
	Dim passwords()
	Dim pswdIndex
	Dim maxPswd
	maxPswd = 0
	ReDim passwords(9)
	While Not fil_pswd.AtEndOfStream
	    passwords(pswdIndex) = fil_pswd.ReadLine
	    pswdIndex = pswdIndex +1
	    maxPswd = maxPswd +1
	    If pswdIndex = UBound(passwords) then
		ReDim Preserve passwords(UBound(passwords)*2)
	    End If
	Wend
	fil_pswd.Close

	Dim name
	Dim connectedOK
	Dim isLocallyManaged
	Dim cnxnString
	Dim cnxnType
	Dim errorID
	Dim errorDescr
	Dim authFailureCount
	Dim readline
	Dim crtResult
	Dim crtResult2
	Dim crtResult3


Sub Main
    authFailureCount = 0
    pswdIndex = 0
    crt.Screen.Synchronous = True

    ' Print headers for the CSV
    name = "Hostname"
    WriteToCSVWithCmd "Output", "Command"

    Do While Not fil.AtEndOfStream
	line = fil.ReadLine '
	name = LCase(line)
	fil_out.Write "-----------------------------" & vbCrLf & name & vbCrLf & "-----------------------------" & vbCrLf
	password = passwords( pswdIndex )
	
	' Timestamp for every device (in addtion to the show clock)
	WriteToResults "Timestamp: " & Now()

	' We'll try telnet by default. The LoginOntoDevice function will handle SSH automatically.
	If LoginOntoDevice( "Telnet", name ) Then
	    authFailureCount = 0 '{{{
	    cmdIndex = 0
	    While commands(cmdIndex) <> ""
		crt.Screen.Send commands(cmdIndex) & vbCr
		' stillOutputting = true
		readline = crt.Screen.ReadString( name & "#", name & ">", name & "> (enable)" )
		DumpOutput readline, commands(cmdIndex)
		cmdIndex = cmdIndex + 1
	    Wend
	    crt.Screen.Send "exit" & vbCr

	    If isLocallyManaged = True Then
		WriteToResults  "SUCCESSFUL. LOCALLY MANAGED!"
	    Else
		WriteToResults  "SUCCESSFUL"
	    End If '}}}
	End If ' If connectedOK

	WriteToOUT "Safeword passwords left: " & maxPswd - pswdIndex
	FlushPasswords

    Loop ' While NotEOF
	
    fil.Close

    MsgBox("Script execution has been finished.")
End Sub

Sub WriteToCSVWithCmd ( message, cmd )
    fil_csv.Write name & "`" & cmd & "`" & message & vbCrLf
End Sub

Sub WriteToOUT( message )
    fil_out.Write message & vbCrLf
    WriteToCSVWithCmd  message, "AI Script" 
End Sub

Sub WriteToResults( message )
    fil_results.Write name & "`" & message & vbCrLf
    WriteToOUT( message )
    ' fil_out.Write message & vbCrLf
    ' fil_csv.Write name & "`" & message & vbCrLf
End Sub


' Copy unused passwords back
Sub FlushPasswords
    Set fil_pswd = fso_pswd.OpenTextFile(PASSWORDS_FILE_PATH, ForWriting, True)
    dim currPswd
    For currPswd = pswdIndex To maxPswd - 1
	fil_pswd.Write passwords( currPswd ) & vbCrLf
    Next
    fil_pswd.Close
End Sub

' Locally managed devices
Sub LoginStatic
    crt.Screen.Send staticPassword & vbCr '{{{
    Dim index
    Dim index3
    index = crt.Screen.WaitForStrings("#", "% Bad passwords", "Password:", "(enable)", ">", "Username:", "% Bad secrets", "% Authentication failed", waitingTimeout)
    Select Case index
	Case 1 ' #
	    connectedOK = true
	Case 2 ' Bad passwords
	    WriteToResults  "Static password authentication failure: Bad passwords"
	Case 3 ' Asking for password again
	    WriteToResults  "Static password authentication failure: Asking for password again"
	Case 4 ' CatOS
	    WriteToOUT "CatOS"
	    commands(0) = "set len 0"
	    connectedOK = true
	Case 5 ' >
	    WriteToOUT "Not in 'enable'" '{{{
	    crt.Screen.Send "enable" & vbCr
	    index3 = crt.Screen.WaitForStrings("#", "% Bad secrets", "Password:", "(enable)", ">", waitingTimeout)
	    Select Case index3
		Case 1 ' #
		    connectedOK = true
		Case 2 ' Bad secrets
		    WriteToResults  "Static password enable failure: Bad secrets"
		Case 3 ' Password:
		    LoginStatic
		Case 4 ' (enable)
		    connectedOK = true
		    WriteToOUT "CatOS"
		    commands(0) = "set len 0"
		Case 5 ' >
		    WriteToResults  "Static password enable failure: Unknown issue"
	    End Select '}}}
	Case 6 ' Asking for username again, will try static one now
	    crt.Screen.Send staticUsername & vbCr '{{{
	    index = crt.Screen.WaitForStrings( "Password:", waitingTimeout )
	    Select Case index
		Case 0
		    WriteToResults  "Timed out waiting for static password authentication"
		Case 1 ' Will send static password again
		    LoginStatic
	    End Select '}}}
	Case 7 ' Bad enable secrets
	    WriteToResults  "Static password enable failure: Bad secrets"
	Case 8 ' TACACS	responding to static passwords
	    WriteToResults  "TACACS issue: username not recognised or TACACS connectivity problem"
	Case 0
	    WriteToResults  "Timed out waiting for static password authentication"
    End Select '}}}
End Sub

Sub DumpOutput( rawData, cmd )
    Dim strings
    Dim i
    strings = Split( rawData, Chr(13)&Chr(10) )
    For i = 0 To Ubound(strings) - 1
	WriteToCSVWithCmd  strings(i), cmd 
    Next
End Sub

' Issue a connect command to the shell (presumably a TFTP/Syslog server)
' Telnet and SSH are currently supported as "cnxnType"
Sub ConnectFromShell( cnxnType, hostname )
    Select Case cnxnType ' {{{
	Case "Telnet"
	    cnxnString = "telnet " & hostname & vbCr
	Case "SSH"
	    cnxnString = "ssh " & hostname & vbCr
    End Select

    crt.Screen.Send vbCrLf & vbCrLf
    crtResult = crt.Screen.WaitForStrings("$", waitingTimeout)
    Select Case crtResult
	Case 1
	    crt.Screen.Send cnxnString
	Case 0
	    TotalMeltdown  "FATAL ERROR: Timed out waiting for shell prompt ($)"
    End Select
    isLocallyManaged = false
    commands(0) = "terminal length 0"
    connectedOK = false ' }}}
End Sub


' Graceful script shutdown
Sub TotalMeltdown( messageForThePosterity )
    WriteToResults  messageForThePosterity
    FlushPasswords
    fil.Close
End Sub


' Login onto a device
' Parametres:
' cnxnType: "Telnet" and "SSH" support
' hostname: can be an IP address as well
' Returns: TRUE or FALSE (success or failure)
Function LoginOntoDevice( cnxnType, hostname )

   Call ConnectFromShell( cnxnType, hostname ) 
   
    ' TODO: try to use a next pswd if this one fails instead of going to the next device
    crtResult = crt.Screen.WaitForStrings("Username:", "Password:", "telnet: Unable to connect to remote host", "telnet: Name or service not known", "Connection refused", "failed to resolve", "Connection timed out", "'s password", "open failed", waitingTimeout)
	Select Case crtResult '{{{

	    ' Tired of waiting
	    Case 0
		WriteToResults  "Timed out waiting for login prompt"
		LoginOntoDevice = False

	    ' "Username:"
	    Case 1
		crt.Screen.Send safewordUsername & vbCr
		errorID = crt.Screen.WaitForStrings( "Enter Safeword Password:", "Password:", waitingTimeout )
		Select Case errorID ' {{{
		    Case 0
			WriteToResults  "Timed out waiting for Safeword password prompt"
			LoginOntoDevice = False
		    Case 1, 2 ' TACACS. "Password" is the new TACACS response!
			LoginOntoDevice = SendSafeWordPassword
		End Select '}}}

	    ' "Password:" or "'s password"
	    Case 2, 8
		Select Case cnxnType
		    Case "Telnet"
		    ' This must be locally managed
			isLocallyManaged = True
			LoginStatic

		    Case "SSH"
			' Let's login using SafeWord
			LoginOntoDevice = SendSafeWordPassword
		End Select

	    ' "telnet: Unable to connect to remote host"
	    Case 3
		WriteToResults "Telnet: Unable to connect to remote host"
		LoginOntoDevice = False

	    ' "telnet: Name or service not known"
	    Case 4
		WriteToResults "Telnet: Hostname could not be resolved in DNS"
		LoginOntoDevice = False

	    ' "Connection refused" - let's try SSH!
	    ' If we have already done so, then it's a connection failure. Dump into the log and move on.
	    ' TODO
	     Case 5
		 Select Case cnxnType
		    Case "Telnet"
			' Let's try SSH then!
			LoginOntoDevice = LoginOntoDevice( "SSH", hostname )
		    Case "SSH"
			WriteToResults "SSH: Connection refused"
			LoginOntoDevice = False
		 End Select

	    ' failed to resolve
	    Case 6
		WriteToResults "SSH: failed to resolve hostname"
		LoginOntoDevice = False
	
	    ' Connection timed out
	    Case 7
		WriteToResults "SSH: connection timed out"
		LoginOntoDevice = False

	    ' open failed
	    Case 9
		WriteToResults "SSH: connection open failed"
		LoginOntoDevice = False

	End Select '}}}
End Function

' Send a TACACS password
Function SendSafeWordPassword()
    If pswdIndex = maxPswd Then  '{{{
	WriteToResults  "No Safeword passwords available. Terminating connection"
	crt.Screen.Send Chr(3)
	SendSafeWordPassword = False
    Else
	crt.Screen.Send password & vbCr
	crtResult3 = crt.Screen.WaitForStrings("#", "% Authentication failed", "(enable)", "password", waitingTimeout)
	Select Case crtResult3
	    ' "#"
	    Case 1
		pswdIndex = pswdIndex +1
		SendSafeWordPassword = True
		
	    ' "% Authentication failed" and "password"
	    Case 2, 4
		WriteToResults  "Safeword authentication failure" '{{{
		pswdIndex = pswdIndex +1
		authFailureCount = authFailureCount + 1
		If authFailureCount = 3 then
		    If MsgBox( "Your Safeword password was not accepted three times. Do you really want to continue? Persistence could cost you dear! (ask Alex Vezyr - he knows)", vbYesNo + vbDefaultButton2 + vbCritical, "Safeword authentication failure") = vbNo Then
			TotalMeltdown( "Safeword Authentication Failure: more than thrice!" )
			SendSafeWordPassword = False
			Exit Function
		    End if
		End if '}}}
		SendSafeWordPassword = False

	    ' (enable)
	    Case 3
		WriteToOUT "CatOS"
		pswdIndex = pswdIndex +1
		commands(0) = "set len 0"
		SendSafeWordPassword = True
	'     Case 4
	' 	WriteToResults "TACACS issue: Not in 'enable'"
	' 	pswdIndex = pswdIndex +1
	' 	crt.Screen.Send "exit" & vbCr
	    Case 0
		WriteToResults  "Timed out waiting for Safeword password prompt"
		SendSafeWordPassword = False
	End Select
    End If '}}}
End Function
