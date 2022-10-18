' BASED ON: https://www.codeproject.com/Articles/14053/Adding-an-LDAP-address-book-to-MS-Outlook
' MODIFIED BY nacken2008

' IGNORE ANY ERRORS:
'On Error Resume Next

const HKEY_CURRENT_USER = &H80000001

Set objShell = WScript.CreateObject("Wscript.Shell")
Set objRegistry = GetObject( "winmgmts:{impersonationLevel=impersonate}!\\" & CreateObject("WScript.Network").ComputerName & "\root\default:StdRegProv")

strRegistryFolder = "Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook"
strMailAccountsSubKey = "9375CFF0413111d3B88A00104B2A6676"
strSearchString = "Service UID"
strLDAPDisplayNameID = "001e3001"
strLDAPServerNameID = "001e6600"
strLDAPDisplayName = "ADRESSES_NAME_BOOK"
strLDAPServerName = "Ip-addresses or Hostname your AD"
strLDAPPort = "389"
strLDAPSearchBase = "OU=YOUR OU,DC=test,DC=local"
strLDAPUserLogin = "AD_USER"
strLDAPUserPassword = "AD_PASS"

If objRegistry.EnumKey(HKEY_CURRENT_USER, strRegistryFolder, arrSubKeys) <> 0 Then
		'msgbox("Outlook Profile does not exist")
		WScript.Quit
End If
If objRegistry.EnumKey(HKEY_CURRENT_USER, strRegistryFolder & "\" & "e8cb48869c395445ade13e3c1c80d154", arrSubKeys) = 0 Then
		'msgbox("Key e8cb48869c395445ade13e3c1c80d154 already exists")
		WScript.Quit
End If

strCMD = "powershell.exe -noninteractive -command " & Chr(34) & "$Password = '" & strLDAPUserPassword & "' | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString; Write-Host $Password" & Chr(34)
Set objScriptExec = objShell.Exec(strCMD)
objScriptExec.StdIn.Close()
strLDAPUserPassword = objScriptExec.StdOut.ReadAll
If (objScriptExec.status = 1) Then
	DELIMITER = "||"
	i = 0
	do while len(strLDAPUserPassword) <> 0
		current = left(strLDAPUserPassword, 2)
		strLDAPUserPassword = right(strLDAPUserPassword, len(strLDAPUserPassword) - len(current))
		output = output & "&H" & current & DELIMITER
		If ((3 < i) And (i <= 19)) Then
			strS001e67f1_1 = strS001e67f1_1 & "&H" & current & DELIMITER
		End If
		If ((23 < i) And (i <= 39)) Then
			strS001e67f1_2 = strS001e67f1_2 & "&H" & current & DELIMITER
		End If
		i = i + 1
	loop
	output = left(output, len(output) - len(DELIMITER) - 5)
	objLDAPUserPassword = split(output, DELIMITER)
	strS001e67f1 = "&H01||&H00||&H00||&H00||" & strS001e67f1_1 & "&H01||&H00||&H00||&H00||" & strS001e67f1_2 & "&H00||&H00||&H00||&H00||&H14||&H00||&H00||&H00||&H53||&H00||&H30||&H00||&H30||&H00||&H31||&H00||&H65||&H00||&H36||&H00||&H37||&H00||&H66||&H00||&H31||&H00||&H00||&H00||&H03||&H66||&H00||&H00||&Hc0||&H00||&H00||&H00||&H10||&H00||&H00||&H00||&H53||&H31||&Hf3||&H19||&H7e||&Hbb||&H8a||&Hb6||&H59||&Hcd||&H26||&Hf6||&H3d||&H75||&Hc8||&Hc2||&H00||&H00||&H00||&H00||&H04||&H80||&H00||&H00||&Ha0||&H00||&H00||&H00||&H10||&H00||&H00||&H00||&H60||&H21||&H78||&H2d||&Hb2||&H24||&He4||&H4c||&H77||&Hb4||&H1b||&H98||&Hbc||&Hec||&H40||&H3e||&H08||&H00||&H00||&H00||&Hba||&H6f||&Hed||&Ha4||&H68||&Hcd||&H84||&Hd5||&H14||&H00||&H00||&H00||&H49||&H88||&H84||&H2d||&Hc8||&H50||&H35||&Hcb||&Hfa||&H43||&He7||&H15||&Hdd||&Hbb||&H9c||&H39||&Hc1||&Hf9||&H09||&H94"	
	objS001e67f1 = split(strS001e67f1, DELIMITER)
	strFlagCreateAccount = "yes"
Else
	strFlagCreateAccount = "no"
End if

objRegistry.EnumKey HKEY_CURRENT_USER, strRegistryFolder & "\" & strMailAccountsSubKey, arrProfiles
For Each strSubfolder In arrProfiles
	'msgbox (strSubfolder)
	objRegistry.GetBinaryValue HKEY_CURRENT_USER, strRegistryFolder & "\" & strMailAccountsSubKey & "\" & strSubfolder, strSearchString, strRetVal
	strSubfolderName = ""
	For i = lBound(strRetVal) to uBound(strRetVal)
		strRetVal_temp = Dec2Hex(strRetVal(i))
		If (Len (strRetVal_temp) < 2) Then
			strSubfolderName = strSubfolderName & "0" & strRetVal_temp
		Else 
			strSubfolderName = strSubfolderName & strRetVal_temp
		End if
	Next
	If (strSubfolderName <> "") Then
		'msgbox (strRegistryFolder & "\" & LCase (strSubfolderName))
		objRegistry.GetStringValue HKEY_CURRENT_USER, strRegistryFolder & "\" & LCase (strSubfolderName), strLDAPDisplayNameID, strLDAPDisplayNameValue
		If (strLDAPDisplayNameValue <> "") Then
			If (strLDAPDisplayNameValue = strLDAPDisplayName) Then
				strFlagCreateAccount = "no"
			End if
		End if
		objRegistry.GetStringValue HKEY_CURRENT_USER, strRegistryFolder & "\" & LCase (strSubfolderName), strLDAPServerNameID, strLDAPServerNameValue
		If (strLDAPServerNameValue <> "") Then
			If (strLDAPServerNameValue = strLDAPServerName) Then
				strFlagCreateAccount = "no"
			End if
		End if
	End if
Next

If (strFlagCreateAccount = "yes") Then
	CreateAccount objRegistry, HKEY_CURRENT_USER, strRegistryFolder, strLDAPDisplayName, strLDAPServerName, strLDAPPort, strLDAPSearchBase, strLDAPUserLogin, objLDAPUserPassword, objS001e67f1
End If

Function Dec2Hex (ByVal numAny)
	Dim Sign
	Const maxNum = 9007199254740991
	Const HexChars = "0123456789ABCDEF"
	Sign = Sgn(numAny)
	numAny = Fix(Abs(CDbl(numAny)))
	If numAny > CDbl(maxNum) Then
		Wscript.Echo "Dec2Hex Error: " & numAny & " must be greater/less than +/- 9,007,199,254,740,991"
		Dec2Hex = Empty
	Exit Function
	End If 'numAny > maxNum
	If numAny = 0 Then
		Dec2Hex = "0"
		Exit Function
	End If
	While numAny > 0
		Dec2Hex = Mid(HexChars, 1 + (numAny - 16 * Fix(numAny / 16)), 1) & Dec2Hex
		numAny = Fix(numAny/16)
	WEnd
	If Sign = -1 Then Dec2Hex = "-" & Dec2Hex
End Function 'Dec2Hex

Function CreateAccount (objRegistry, HKEY_CURRENT_USER, strRegistryFolder, strLDAPDisplayName, strLDAPServerName, strLDAPPort, strLDAPSearchBase, strLDAPUserLogin, objLDAPUserPassword, objS001e67f1)
'Add Ldap Type Key
	sKeyPath = strRegistryFolder & "\" & "e8cb48869c395445ade13e3c1c80d154\"
	objRegistry.CreateKey HKEY_CURRENT_USER, sKeyPath
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "001f300a", Array(&H45,&H00,&H4d,&H00,&H41,&H00,&H42,&H00,&H4c,&H00,&H54,&H00,&H2e,&H00,&H44,&H00,&H4c,&H00,&H4c,&H00,&H00,&H00)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "001f3d13", Array(&H7b,&H00,&H36,&H00,&H34,&H00,&H38,&H00,&H35,&H00,&H44,&H00,&H32,&H00,&H36,&H00,&H38,&H00,&H2d,&H00,&H43,&H00,&H32,&H00,&H41,&H00,&H43,&H00,&H2d,&H00,&H31,&H00,&H31,&H00,&H44,&H00,&H31,&H00,&H2d,&H00,&H41,&H00,&H44,&H00,&H33,&H00,&H45,&H00,&H2d,&H00,&H31,&H00,&H30,&H00,&H41,&H00,&H30,&H00,&H43,&H00,&H39,&H00,&H31,&H00,&H31,&H00,&H43,&H00,&H39,&H00,&H43,&H00,&H30,&H00,&H7d,&H00,&H00,&H00)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "001f3006", Array(&H4d,&H00,&H69,&H00,&H63,&H00,&H72,&H00,&H6f,&H00,&H73,&H00,&H6f,&H00,&H66,&H00,&H74,&H00,&H20,&H00,&H4c,&H00,&H44,&H00,&H41,&H00,&H50,&H00,&H2d,&H00,&H56,&H00,&H65,&H00,&H72,&H00,&H7a,&H00,&H65,&H00,&H69,&H00,&H63,&H00,&H68,&H00,&H6e,&H00,&H69,&H00,&H73,&H00,&H00,&H00)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "00033e03", Array(&H23,0,0,0)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "01023d0c", Array(&H5c,&Hb9,&H3b,&H24,&Hff,&H71,&H07,&H41,&Hb7,&Hd8,&H3b,&H9c,&Hb6,&H31,&H79,&H92)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "001f3d09", Array(&H45,&H00,&H4d,&H00,&H41,&H00,&H42,&H00,&H4c,&H00,&H54,&H00,&H00,&H00)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "001f3001", Array(&H4d,&H00,&H69,&H00,&H63,&H00,&H72,&H00,&H6f,&H00,&H73,&H00,&H6f,&H00,&H66,&H00,&H74,&H00,&H20,&H00,&H4c,&H00,&H44,&H00,&H41,&H00,&H50,&H00,&H2d,&H00,&H56,&H00,&H65,&H00,&H72,&H00,&H7a,&H00,&H65,&H00,&H69,&H00,&H63,&H00,&H68,&H00,&H6e,&H00,&H69,&H00,&H73,&H00,&H00,&H00)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "00033009", Array(0,0,0,0)
'Add Ldap connection settings key
	sKeyPath = strRegistryFolder & "\" & "5cb93b24ff710741b7d83b9cb6317992\"
	objRegistry.CreateKey HKEY_CURRENT_USER, sKeyPath
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "001f3d13", Array(&H7b,&H00,&H36,&H00,&H34,&H00,&H38,&H00,&H35,&H00,&H44,&H00,&H32,&H00,&H36,&H00,&H38,&H00,&H2d,&H00,&H43,&H00,&H32,&H00,&H41,&H00,&H43,&H00,&H2d,&H00,&H31,&H00,&H31,&H00,&H44,&H00,&H31,&H00,&H2d,&H00,&H41,&H00,&H44,&H00,&H33,&H00,&H45,&H00,&H2d,&H00,&H31,&H00,&H30,&H00,&H41,&H00,&H30,&H00,&H43,&H00,&H39,&H00,&H31,&H00,&H31,&H00,&H43,&H00,&H39,&H00,&H43,&H00,&H30,&H00,&H7d,&H00,&H00,&H00)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "001f3d0a", Array(&H45,&H00,&H4d,&H00,&H41,&H00,&H42,&H00,&H4c,&H00,&H54,&H00,&H2e,&H00,&H44,&H00,&H4c,&H00,&H4c,&H00,&H00,&H00)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "101e3d0f", Array(&H01,&H00,&H00,&H00,&H08,&H00,&H00,&H00,&H45,&H4d,&H41,&H42,&H4c,&H54,&H2e,&H44,&H4c,&H4c,&H00)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "001f3d0b", Array(&H53,&H00,&H65,&H00,&H72,&H00,&H76,&H00,&H69,&H00,&H63,&H00,&H65,&H00,&H45,&H00,&H6e,&H00,&H74,&H00,&H72,&H00,&H79,&H00,&H00,&H00)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "00033009", Array(&H20,&H00,&H00,&H00)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "001f6604", Array(&H28,&H00,&H26,&H00,&H28,&H00,&H6d,&H00,&H61,&H00,&H69,&H00,&H6c,&H00,&H3d,&H00,&H2a,&H00,&H29,&H00,&H28,&H00,&H7c,&H00,&H28,&H00,&H6d,&H00,&H61,&H00,&H69,&H00,&H6c,&H00,&H3d,&H00,&H25,&H00,&H73,&H00,&H2a,&H00,&H29,&H00,&H28,&H00,&H7c,&H00,&H28,&H00,&H63,&H00,&H6e,&H00,&H3d,&H00,&H25,&H00,&H73,&H00,&H2a,&H00,&H29,&H00,&H28,&H00,&H7c,&H00,&H28,&H00,&H73,&H00,&H6e,&H00,&H3d,&H00,&H25,&H00,&H73,&H00,&H2a,&H00,&H29,&H00,&H28,&H00,&H67,&H00,&H69,&H00,&H76,&H00,&H65,&H00,&H6e,&H00,&H4e,&H00,&H61,&H00,&H6d,&H00,&H65,&H00,&H3d,&H00,&H25,&H00,&H73,&H00,&H2a,&H00,&H29,&H00,&H29,&H00,&H29,&H00,&H29,&H00,&H29,&H00,&H00,&H00)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "001f3d09", Array(&H45,&H00,&H4d,&H00,&H41,&H00,&H42,&H00,&H4c,&H00,&H54,&H00,&H00,&H00)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "01023d01", Array(&He8,&Hcb,&H48,&H86,&H9c,&H39,&H54,&H45,&Had,&He1,&H3e,&H3c,&H1c,&H80,&Hd1,&H54)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "01023615", Array(&H50,&Ha7,&H0a,&H61,&H55,&Hde,&Hd3,&H11,&H9d,&H60,&H00,&Hc0,&H4f,&H4c,&H8e,&Hfa)
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e6600", strLDAPServerName
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e6601", strLDAPPort
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e6602", strLDAPUserLogin
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "S001e67f1", objS001e67f1
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "000b6622", Array(&H01,&H00)
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e6603", strLDAPSearchBase
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e6605", "SMTP"
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e6606", "mail"
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e6607", "60"
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e6608", "100"
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e6609", "120"
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e660a", "15"
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e660b", ""
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e660c", "OFF"
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e660d", "OFF"
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e660e", "NONE"
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e660f", "OFF"
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e6610", "postalAddress"
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e6611", "cn"
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e6612", "1"
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e3001", strLDAPDisplayName
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "000b6613", Array(&H00,&H00)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "000b6615", Array(&H01,&H00)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "01026617", objLDAPUserPassword
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "00036623", Array(&H00,&H00,&H00,&H00)
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "01026631", Array(&H5b,&Hfe,&H3f,&He9,&H65,&H55,&H19,&H48,&H9c,&H52,&H2d,&H68,&Hfc,&Hb9,&H89,&Hbf)
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e3d09", "EMABLT"
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e3d0a", "BJABLR.DLL"
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e3d0b", "ServiceEntry"
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e3d13", "{6485D268-C2AC-11D1-AD3E-10A0C911C9C0}"
	objRegistry.SetStringValue HKEY_CURRENT_USER, sKeyPath, "001e6604", "(&(mail=*)(|(mail=%s*)" & "(|(cn=%s*)(|(sn=%s*)(givenName=%s*)))))"'objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "001e67f1", Array(&H0a)
'Append to Backup Key for ldap types
	sKeyPath = strRegistryFolder & "\" & "9207f3e0a3b11019908b08002b2a56c2\"
	objRegistry.getBinaryValue HKEY_CURRENT_USER, sKeyPath, "01023d01", Backup
	Dim oldLength
	oldLength = UBound(Backup)
	ReDim Preserve Backup(oldLength+16)
	Backup(oldLength+1) = &He8
	Backup(oldLength+2) = &Hcb
	Backup(oldLength+3) = &H48
	Backup(oldLength+4) = &H86
	Backup(oldLength+5) = &H9c
	Backup(oldLength+6) = &H39
	Backup(oldLength+7) = &H54
	Backup(oldLength+8) = &H45
	Backup(oldLength+9) = &Had
	Backup(oldLength+10) = &He1
	Backup(oldLength+11) = &H3e
	Backup(oldLength+12) = &H3c
	Backup(oldLength+13) = &H1c
	Backup(oldLength+14) = &H80
	Backup(oldLength+15) = &Hd1
	Backup(oldLength+16) = &H54
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "01023d01", Backup
	
'Get Contacts Registry Key
	For num = LBound(Backup) To LBound(Backup) + 15
		strRetVal_temp = Dec2Hex(Backup(num))
		If (Len (strRetVal_temp) < 2) Then
			contactskey = contactskey & "0" & strRetVal_temp
		Else 
			contactskey = contactskey & strRetVal_temp
		End if
	Next

'array for ABSearchOrder
	Dim ABSearchOrder(203)
	ABSearchOrder(0) = &H03
	ABSearchOrder(1) = &H00
	ABSearchOrder(2) = &H00
	ABSearchOrder(3) = &H00
	ABSearchOrder(4) = &H1e
	ABSearchOrder(5) = &H00
	ABSearchOrder(6) = &H00
	ABSearchOrder(7) = &H00
	ABSearchOrder(8) = &H1c
	ABSearchOrder(9) = &H00
	ABSearchOrder(10) = &H00
	ABSearchOrder(11) = &H00
	ABSearchOrder(12) = &H5a
	ABSearchOrder(13) = &H00
	ABSearchOrder(14) = &H00
	ABSearchOrder(15) = &H00
	ABSearchOrder(16) = &H3c
	ABSearchOrder(17) = &H00
	ABSearchOrder(18) = &H00
	ABSearchOrder(19) = &H00
	ABSearchOrder(20) = &H33
	ABSearchOrder(21) = &H00
	ABSearchOrder(22) = &H00
	ABSearchOrder(23) = &H00
	ABSearchOrder(24) = &H98
	ABSearchOrder(25) = &H00
	ABSearchOrder(26) = &H00
	ABSearchOrder(27) = &H00
	ABSearchOrder(28) = &H00
	ABSearchOrder(29) = &H00
	ABSearchOrder(30) = &H00
	ABSearchOrder(31) = &H00
'GAL ID 1, length 16
	For n = 32 To 47
		ABSearchOrder(n) = Backup(n - 16)
	Next
'GAL ID 1 end
	ABSearchOrder(48) = &H01
	ABSearchOrder(49) = &H00
	ABSearchOrder(50) = &H00
	ABSearchOrder(51) = &H00
	ABSearchOrder(52) = &H00
	ABSearchOrder(53) = &H01
	ABSearchOrder(54) = &H00
	ABSearchOrder(55) = &H00
	ABSearchOrder(56) = &H2f
	ABSearchOrder(57) = &H00
	ABSearchOrder(58) = &H00
	ABSearchOrder(59) = &H00
	ABSearchOrder(60) = &H00
	ABSearchOrder(61) = &H00
	ABSearchOrder(62) = &H00
	ABSearchOrder(63) = &H00
	ABSearchOrder(64) = &Hfe
	ABSearchOrder(65) = &H42
	ABSearchOrder(66) = &Haa
	ABSearchOrder(67) = &H0a
	ABSearchOrder(68) = &H18
	ABSearchOrder(69) = &Hc7
	ABSearchOrder(70) = &H1a
	ABSearchOrder(71) = &H10
	ABSearchOrder(72) = &He8
	ABSearchOrder(73) = &H85
	ABSearchOrder(74) = &H0b
	ABSearchOrder(75) = &H65
	ABSearchOrder(76) = &H1c
	ABSearchOrder(77) = &H24
	ABSearchOrder(78) = &H00
	ABSearchOrder(79) = &H00
	ABSearchOrder(80) = &H03
	ABSearchOrder(81) = &H00
	ABSearchOrder(82) = &H00
	ABSearchOrder(83) = &H00
	ABSearchOrder(84) = &H03
	ABSearchOrder(85) = &H00
	ABSearchOrder(86) = &H00
	ABSearchOrder(87) = &H00
'Contacts ID 1, length 16 (from 01023d01 (0..15) -> 01026601)
	sKeyPath = strRegistryFolder & "\" & contactskey & "\"	
	objRegistry.getBinaryValue HKEY_CURRENT_USER, sKeyPath, "01026601", Contacts
	For n = 88 To 103
		ABSearchOrder(n) = Contacts(n - 88)
	Next
'Contacts ID 1 end	
'Contacts ID 2, length 48 (from 01023d01 (0..15) -> 11026620 (12..59))
	objRegistry.getBinaryValue HKEY_CURRENT_USER, sKeyPath, "11026620", Contacts
	For n = 104 To 151
		ABSearchOrder(n) = Contacts(n - 92)
	Next
'Contacts ID 2 end
	ABSearchOrder(152) = &H00
	ABSearchOrder(153) = &H00
	ABSearchOrder(154) = &H00
	ABSearchOrder(155) = &H00
'LDAP Adress Book ID, length 16
	ABSearchOrder(156) = &H50
	ABSearchOrder(157) = &Ha7
	ABSearchOrder(158) = &H0a
	ABSearchOrder(159) = &H61
	ABSearchOrder(160) = &H55
	ABSearchOrder(161) = &Hde
	ABSearchOrder(162) = &Hd3
	ABSearchOrder(163) = &H11
	ABSearchOrder(164) = &H9d
	ABSearchOrder(165) = &H60
	ABSearchOrder(166) = &H00
	ABSearchOrder(167) = &Hc0
	ABSearchOrder(168) = &H4f
	ABSearchOrder(169) = &H4c
	ABSearchOrder(170) = &H8e
	ABSearchOrder(171) = &Hfa
'LDAP Adress Book ID end
	ABSearchOrder(172) = &H01
	ABSearchOrder(173) = &H04
	ABSearchOrder(174) = &H00
	ABSearchOrder(175) = &H00
	ABSearchOrder(176) = &Hfe
	ABSearchOrder(177) = &Hff
	ABSearchOrder(178) = &Hff
	ABSearchOrder(179) = &Hff
	ABSearchOrder(180) = &H00
	ABSearchOrder(181) = &H75
	ABSearchOrder(182) = &H72
	ABSearchOrder(183) = &H62
	ABSearchOrder(184) = &H6f
	ABSearchOrder(185) = &H6b
	ABSearchOrder(186) = &H2e
	ABSearchOrder(187) = &H62
	ABSearchOrder(188) = &H65
	ABSearchOrder(189) = &H72
	ABSearchOrder(190) = &H2e
	ABSearchOrder(191) = &H6d
	ABSearchOrder(192) = &H79
	ABSearchOrder(193) = &H74
	ABSearchOrder(194) = &H6f
	ABSearchOrder(195) = &H79
	ABSearchOrder(196) = &H73
	ABSearchOrder(197) = &H2e
	ABSearchOrder(198) = &H64
	ABSearchOrder(199) = &H65
	ABSearchOrder(200) = &H00
	ABSearchOrder(201) = &H00
	ABSearchOrder(202) = &H00
	ABSearchOrder(203) = &H00
	sKeyPath = strRegistryFolder & "\" & "9207f3e0a3b11019908b08002b2a56c2\"
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "11023d05", ABSearchOrder
'Append to Backup Key for ldap connection settings
	sKeyPath = strRegistryFolder & "\" & "9207f3e0a3b11019908b08002b2a56c2\"
	objRegistry.getBinaryValue HKEY_CURRENT_USER, sKeyPath, "01023d0e", Backup
	oldLength = UBound(Backup)
	ReDim Preserve Backup(oldLength+16)
	Backup(oldLength+1) = &H5c
	Backup(oldLength+2) = &Hb9
	Backup(oldLength+3) = &H3b
	Backup(oldLength+4) = &H24
	Backup(oldLength+5) = &Hff
	Backup(oldLength+6) = &H71
	Backup(oldLength+7) = &H07
	Backup(oldLength+8) = &H41
	Backup(oldLength+9) = &Hb7
	Backup(oldLength+10) = &Hd8
	Backup(oldLength+11) = &H3b
	Backup(oldLength+12) = &H9c
	Backup(oldLength+13) = &Hb6
	Backup(oldLength+14) = &H31
	Backup(oldLength+15) = &H79
	Backup(oldLength+16) = &H92
	objRegistry.SetBinaryValue HKEY_CURRENT_USER, sKeyPath, "01023d0e", Backup
'Delete Active Books List Key
	sKeyPath = strRegistryFolder & "\" & "9375CFF0413111d3B88A00104B2A6676"
	objRegistry.DeleteValue HKEY_CURRENT_USER, sKeyPath, "{ED475419-B0D6-11D2-8C3B-00104B2A6676}"
End Function 'CreateAccount