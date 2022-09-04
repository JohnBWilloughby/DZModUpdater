Attribute VB_Name = "Module1"
Option Explicit


Global sModPath As Variant
Global FindTheMods As Variant

Global cCompName    As Variant
Global cDriveLetter As Variant
Global cModPath As Variant

Global RemoteDrive As Variant
Global ConfigFile As Variant

Global strFullFilename As String
Global sMessage As String

Global WhattoMap As String
Global WheretoMap As String



Private Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Private Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpszName As String, ByVal bForce As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Const WN_ACCESS_DENIED = 5& 'ERROR_ACCESS_DENIED
Private Const WN_ALREADY_CONNECTED = 85& 'ERROR_ALREADY_ASSIGNED
Private Const WN_BAD_LOCALNAME = 1200& 'ERROR_BAD_DEVICE
Private Const WN_BAD_NETNAME = 67& 'ERROR_BAD_NET_NAME
Private Const WN_BAD_PASSWORD = 86& 'ERROR_INVALID_PASSWORD
Private Const WN_BAD_POINTER = 487& 'ERROR_INVALID_ADDRESS
Private Const WN_BAD_VALUE = 87 'ERROR_INVALID_PARAMETER
Private Const WN_MORE_DATA = 234 'ERROR_MORE_DATA
Private Const WN_NET_ERROR = 59& 'ERROR_UNEXP_NET_ERR
Private Const WN_NOT_CONNECTED = 2250 'ERROR_NOT_CONNECTED
Private Const WN_NOT_SUPPORTED = 50& 'ERROR_NOT_SUPPORTED
Private Const WN_OPEN_FILES = 2401& 'ERROR_OPEN_FILES
Private Const WN_OUT_OF_MEMORY = 8 'ERROR_NOT_ENOUGH_MEMORY
Private Const WN_SUCCESS = 0 'NO_ERROR


Function GetUNCPath(DriveLetter As String, DrivePath, ErrorMsg As _
String) As Long

On Local Error GoTo GetUNCPath_Err
Dim status As Long
Dim lpszLocalName As String
Dim lpszRemoteName As String
Dim cbRemoteName As Long
lpszLocalName = DriveLetter
If Right$(lpszLocalName, 1) <> Chr$(0) Then lpszLocalName = lpszLocalName & Chr$(0)
lpszRemoteName = String$(255, Chr$(32))
cbRemoteName = Len(lpszRemoteName)
status = WNetGetConnection(lpszLocalName, lpszRemoteName, cbRemoteName)

GetUNCPath = status

Select Case status
Case WN_SUCCESS
    ErrorMsg = "Function Performed successfully"
Case WN_NOT_SUPPORTED
    ErrorMsg = "This function is not supported"
Case WN_OUT_OF_MEMORY
    ErrorMsg = "The System is Out of Memory."
Case WN_NET_ERROR
    ErrorMsg = "An error occurred on the network"
Case WN_BAD_POINTER
    ErrorMsg = "The network path is invalid"
Case WN_BAD_VALUE
    ErrorMsg = "Invalid local device name"
Case WN_NOT_CONNECTED
    ErrorMsg = "The drive is not connected"
Case WN_MORE_DATA
    ErrorMsg = "The buffer was too small to return the fileservice name"
Case Else
    ErrorMsg = "Unrecognized error - " & Str$(status) & "."
End Select

If Len(ErrorMsg) Then
DrivePath = ""
Else
' Trim it, and remove any nulls
DrivePath = StripNulls(lpszRemoteName)
End If
Exit Function

GetUNCPath_Err:
MsgBox Err.Description, vbInformation
End Function

Function sGetUserName() As String

Dim lpBuffer As String * 255
Dim lret As Long
lret = GetUserName(lpBuffer, 255)
sGetUserName = StripNulls(lpBuffer)
End Function

Private Function StripNulls(s As String) As String

'Truncates string at first null character, any text after first null is lost
Dim I As Integer
StripNulls = s
If Len(s) Then
I = InStr(s, Chr$(0))
If I Then StripNulls = Left$(s, I - 1)
End If
End Function


Function MapNetworkDrive(UNCname As String, Password As String, DriveLetter As String, ErrorMsg As String) As Long

Dim status As Long
Dim tUNCname As String, tPassword As String, tDriveLetter As String
On Local Error GoTo MapNetworkDrive_Err
tUNCname = UNCname
tPassword = Password
tDriveLetter = DriveLetter
'If Right$(tUNCname, 1) <> Chr$(0) Then tUNCname = tUNCname & Chr$(0)
'If Right$(tPassword, 1) <> Chr$(0) Then tPassword = tPassword & Chr$(0)
'If Right$(tDriveLetter, 1) <> Chr$(0) Then tDriveLetter = tDriveLetter & Chr$(0)
status = WNetAddConnection(tUNCname, tPassword, tDriveLetter)

Select Case status
Case WN_SUCCESS
    ErrorMsg = ""
Case WN_NOT_SUPPORTED
    ErrorMsg = "Function is not supported."
Case WN_OUT_OF_MEMORY:
    ErrorMsg = "The system is out of memory."
Case WN_NET_ERROR
    ErrorMsg = "An error occurred on the network."
Case WN_BAD_POINTER
    ErrorMsg = "The network path is invalid."
Case WN_BAD_NETNAME
    ErrorMsg = "Invalid network resource name."
Case WN_BAD_PASSWORD
    ErrorMsg = "The password is invalid."
Case WN_BAD_LOCALNAME
    ErrorMsg = "The local device name is invalid."
Case WN_ACCESS_DENIED
    ErrorMsg = "A security violation occurred."
Case WN_ALREADY_CONNECTED
    ErrorMsg = "This drive letter is already connected to a network drive."
Case Else
    ErrorMsg = "Unrecognized error - " & Str$(status) & "."
End Select


MapNetworkDrive = status
MapNetworkDrive_End:
Exit Function
MapNetworkDrive_Err:
MsgBox Err.Description, vbInformation
Resume MapNetworkDrive_End
End Function



Function DisconnectNetworkDrive(DriveLetter As String, ForceFileClose As Long, ErrorMsg As String) As Long

Dim status As Long
Dim tDriveLetter As String
On Local Error GoTo DisconnectNetworkDrive_Err
tDriveLetter = DriveLetter
If Right$(tDriveLetter, 1) <> Chr$(0) Then tDriveLetter = tDriveLetter & Chr$(0)
status = WNetCancelConnection(tDriveLetter, ForceFileClose)

Select Case status
Case WN_SUCCESS
    ErrorMsg = ""
Case WN_BAD_POINTER:
    ErrorMsg = "The network path is invalid."
Case WN_BAD_VALUE
    ErrorMsg = "Invalid local device name"
Case WN_NET_ERROR:
    ErrorMsg = "An error occurred on the network."
Case WN_NOT_CONNECTED
    ErrorMsg = "The drive is not connected"
Case WN_NOT_SUPPORTED
    ErrorMsg = "This function is not supported"
Case WN_OPEN_FILES
    ErrorMsg = "Files are in use on this service. Drive was not disconnected."
Case WN_OUT_OF_MEMORY:
    ErrorMsg = "The System is Out of Memory"
Case Else:
    ErrorMsg = "Unrecognized error - " & Str$(status) & "."
End Select


DisconnectNetworkDrive = status
Exit Function
DisconnectNetworkDrive_Err:
MsgBox Err.Description, vbInformation
End Function


