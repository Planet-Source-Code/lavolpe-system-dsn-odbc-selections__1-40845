VERSION 5.00
Begin VB.Form frmLstODBC 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Current System DSN ODBC Configurations"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5550
   HelpContextID   =   104
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Refresh"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   735
      Width           =   900
   End
   Begin VB.ComboBox lstODBC 
      Height          =   315
      ItemData        =   "ODBClist.frx":0000
      Left            =   120
      List            =   "ODBClist.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
   Begin VB.CommandButton cmdODBCexe 
      Caption         =   "Control Panel's ODBC Config"
      Height          =   375
      Left            =   1020
      TabIndex        =   3
      Top             =   735
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3450
      TabIndex        =   2
      Top             =   735
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4665
      TabIndex        =   1
      Top             =   735
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "Select an ODBC configuration from the list below"
      Height          =   255
      Left            =   135
      TabIndex        =   4
      Top             =   120
      Width           =   5280
   End
End
Attribute VB_Name = "frmLstODBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PURPOSE: A reusable form that will quickly display all current System ODBC DSNs for selection
' The programmer then can use the connection name in an ADO or DAO connection

'NOTES:
' 1. This project does not test an ODBC connection once it has been selected. That is up to you.
' 2. The form does not offer user to enter user name or password and is intended for trusted connections
'  -- However, if needed, you can easily add that option to this form or offer that when form closes
' 3. If the user's permissions are locked down and cannot access the registry to read the information
'     therein, then this project may be useless to you.
' 4. The form should be opened modally (i.e., frmLstODBC.Show 1, Me)

Option Explicit
Option Compare Text

' Windows Registry Root Key Constants
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_DYN_DATA = &H80000006
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_USERS = &H80000003

' Registry Storage Type Constants

Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_EXPAND_SZ = 2
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7
Private Const REG_NONE = 0
Private Const REG_SZ = 1

''' Security Constants

Private Const READ_CONTROL = &H20000
Private Const READ_WRITE = 2

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)

Private Const SYNCHRONIZE = &H100000

''' Registry Security Constants

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_NOTIFY = &H10
Private Const KEY_EVENT = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_SET_VALUE = &H2
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

' Types

'' Security Types

''' Security Attributes

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

''' Access Control

Private Type ACL
        AclRevision As Byte
        Sbz1 As Byte
        AclSize As Integer
        AceCount As Integer
        Sbz2 As Integer
End Type

''' Security Descriptor

Private Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Long
        Owner As Long
        Group As Long
        Sacl As ACL
        Dacl As ACL
End Type

'' Other Types

''' FileTime

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type


' Registry Key Functions

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
    ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
    lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" _
(ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

' Registry Value Functions

'' Registry value functions declared more than once.  This is a kludge for overloading
'' that will be solved in the next version of Visual Studio, to the best of my knowledge,
'' as there will be support for overloading and polymorphism in the next version of Visual Basic

''' Standard
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
''' String Value
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
''' Long Value (DO NOT PASS POINTERS)
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
''' Standard
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
''' Pointer Value
Private Declare Function RegSetValueExPtr Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As Long, ByVal cbData As Long) As Long
''' String Value
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
''' Long Value (DO NOT PASS POINTERS)
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long


' VB Enum of the standard registry key constants

Private Enum StandardRegistryKeys
    HKeyClassesRoot = HKEY_CLASSES_ROOT
    HKeyCurrentConfig = HKEY_CURRENT_CONFIG
    Hkeycurrentuser = HKEY_CURRENT_USER
    HKeyDynamicData = HKEY_DYN_DATA
    HKeyLocalMachine = HKEY_LOCAL_MACHINE
    HKeyPerformanceData = HKEY_PERFORMANCE_DATA
    HKeyUsers = HKEY_USERS
End Enum

Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As String) As Long
Private Const CB_FINDSTRING As Long = &H14C
Private Const CB_FINDSTRINGEXACT As Long = &H158
Private Const LB_FINDSTRINGEXACT As Long = &H1A2
Private Const LB_FINDSTRING As Long = &H18F

''' a handy function that ports to VB well.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private ODBCsetting() As ODBCinfo

Private Sub cmdODBCexe_Click()
' Run the ODBC data sources appelate from the control panel

On Error GoTo BadExecute
' If the user has this file on their computer, & they should have, the following line will execute it
Shell "Odbcad32.exe"
Exit Sub

BadExecute:
MsgBox Err.Description, vbOKOnly + vbExclamation
End Sub

Private Sub cmdOK_Click(Index As Integer)
' Exit and update the DSN, if needed

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo cmdOK_Click_General_ErrTrap

If Index = 0 Then   ' then OK button was clicked
    If lstODBC.ListIndex > -1 Then
        ' update the ODBC info selected
        NewDSNname.Name = lstODBC 'update the DSN connection name
        NewDSNname.Type = ODBCsetting(lstODBC.ItemData(lstODBC.ListIndex)).Type
    End If
Else
    If Index = 2 Then   ' refreshing
        ' save currently highlighted item (if any)
        lstODBC.Tag = lstODBC.Text
        LoadODBCtypes       ' refresh list
        ' re-select the previous item if any
        If Len(lstODBC.Tag) Then lstODBC.ListIndex = FindListItem(lstODBC.hWnd, False, True, lstODBC.Tag)
        lstODBC.Tag = ""
        lstODBC.SetFocus
        Exit Sub
    End If
End If
Unload Me
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
cmdOK_Click_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: cmdOK_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub Form_Load()
NewDSNname.Name = vbNullString
LoadODBCtypes
End Sub

Private Sub Form_Terminate()

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo Form_Terminate_General_ErrTrap
Erase ODBCsetting
Unload Me
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
Form_Terminate_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: Form_Terminate" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub LoadODBCtypes()
' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo Form_Load_General_ErrTrap

Dim HKeyBase As String, SKeyNr As Long, SKeyName As String, strTemp As String
Dim sSubKeys() As Variant

lstODBC.Clear

' Registry key where system DSNs can be located
HKeyBase = "Software\ODBC\ODBC.INI\ODBC Data Sources"
' Returns each value in the key in array (0 to #values, 0 to 1)
'   -- the (0 to 1) are:  0=Key Value Name, 1=Key Value Data
GetAllValues HKeyLocalMachine, HKeyBase, sSubKeys()
' No load the results
ReDim ODBCsetting(LBound(sSubKeys) To UBound(sSubKeys))
For SKeyNr = LBound(sSubKeys) To UBound(sSubKeys)   ' Begin enumerating the key
    ODBCsetting(SKeyNr).Name = sSubKeys(SKeyNr, 0)
    ODBCsetting(SKeyNr).Type = GetRegistryValue(HKeyLocalMachine, HKeyBase, ODBCsetting(SKeyNr).Name)
    ' Add result to combo box which is sorted
    lstODBC.AddItem ODBCsetting(SKeyNr).Name
    ' Add reference to which ODBCsetting to reference
    lstODBC.ItemData(lstODBC.NewIndex) = SKeyNr
Next
' if the current setting was passed in variable DSNname, then show it
lstODBC.ListIndex = FindListItem(lstODBC.hWnd, False, True, DSNname)
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
Form_Load_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: Form_Load" & vbCrLf & Err.Description, vbExclamation + vbOKOnly

End Sub

Private Function OpenKey(ByVal uKey As StandardRegistryKeys, ByVal szSubKey As String, Optional ByVal vntIndexKey As String) As Long
''' Function to open the key

    Dim l As Long, uOpened As Long, SA As SECURITY_ATTRIBUTES
    Dim d As Long
    
    l = RegCreateKeyEx(uKey, szSubKey, 0, vbNullString, 0, KEY_ALL_ACCESS, SA, uOpened, d)
    OpenKey = uOpened
End Function

''' Get Value
Private Function GetRegistryValue(ByRef uKey As StandardRegistryKeys, ByVal szValueKey As String, _
    Optional szSubKey As String = "", Optional DefaultValue As Variant, Optional HKeyCur As Long = 0) As Variant
    Dim lBytes As Long, lType As Long, sINI As String
    Dim vData As Variant, lData As Long, wData As Integer
    Dim sdata As String, bData() As Byte
    
    If IsMissing(DefaultValue) Then DefaultValue = "False"
    If HKeyCur = 0 Then uKey = OpenKey(uKey, szValueKey) Else uKey = HKeyCur
    
    ' Null query to retrieve the type and length of data/check for existence
    If uKey <> 0 Then RegQueryValueEx uKey, szSubKey, 0&, lType, 0&, lBytes
    
    If lBytes = 0 Then
        ' Failed to query value : (If DefaultValue <> Empty then create the key)
        ' then assign the Default Value and exit
        GetRegistryValue = DefaultValue
        If uKey <> 0 Then RegCloseKey uKey
        Exit Function
    End If
    
    ' Add 2 bytes to the length to give us a little room.
    lBytes = lBytes + 2
    
    ' look at the return type from the Null call, and call the
    ' appropriate 'overload' of the QueryValue function
    Select Case lType
    Case REG_SZ     ' For Strings
        vData = String(lBytes, 0)
        sdata = vData
        RegQueryValueExString uKey, szSubKey, 0, lType, sdata, lBytes
        GetRegistryValue = Mid(sdata, 1, lBytes - 1)
    Case REG_MULTI_SZ   ' More than one string, call private function to parse the null terminated data
        vData = String(lBytes, 0)
        sdata = vData
        RegQueryValueExString uKey, szSubKey, 0, lType, sdata, lBytes
        GetRegistryValue = GetMultiSZ(sdata)
    Case REG_BINARY
        ' Get as string data and convert to byte data
        ' ,
        ' if 1 byte then as a single byte
        
        sdata = String(lBytes, 0)
        RegQueryValueExString uKey, szSubKey, 0, lType, sdata, lBytes
        ReDim bData(lBytes)
        CopyMemory bData(0), ByVal StrPtr(sdata), lBytes
        Select Case lBytes
        Case 4  'if 4 bytes then a Long
            CopyMemory lData, bData(0), 4
            GetRegistryValue = lData
        Case 2 'If the binary data is 2 bytes long, the value is returned as a word
            CopyMemory wData, bData(0), 2
            GetRegistryValue = wData
        Case 1 ' if 1 byte then as a single byte
            GetRegistryValue = bData(0)
        Case Else ' if 3 or 5 or more, then return as Byte() array
            GetRegistryValue = bData
        End Select
    ' Value is a dword, call the Long value function
    Case REG_DWORD
        RegQueryValueExLong uKey, szSubKey, 0, lType, lData, lBytes
        GetRegistryValue = lData
        
    ' Retrieve the data into a string if it's anything else.
        
    Case Else
        vData = String(lBytes, 0)
        RegQueryValueExString uKey, szValueKey, 0, lType, vData, lBytes
        vData = Mid(vData, 1, lBytes)
        GetRegistryValue = vData
       
    End Select
RegCloseKey uKey
End Function

' Convert a Registry MultiSZ entry into a String array

Private Function GetMultiSZ(ByVal vData As String) As String()

    Dim I As Long, sMultiString() As String, J As Integer
    
    J = 0
    ReDim sMultiString(J)
    For I = 1 To Len(vData)
    
        If Mid(vData, I, 1) = Chr(0) Then
            J = J + 1
            ReDim Preserve sMultiString(J)
        Else
            sMultiString(J) = sMultiString(J) + Mid(vData, I, 1)
        End If
    Next I

    GetMultiSZ = sMultiString
End Function

Private Sub GetAllValues(hKey As Long, strPath As String, KeyData() As Variant)
'Returns: a 2D array.
'(x,0) is value name
'(x,1) is value type (see constants)

Dim lRegResult As Long
Dim hCurKey As Long
Dim lValueNameSize As Long
Dim strValueName As String
Dim lCounter As Long
Dim byDataBuffer(4000) As Byte
Dim lDataBufferSize As Long
Dim lValueType As Long
Dim strNames() As String
Dim lTypes() As Long
Dim intZeroPos As Integer

hCurKey = OpenKey(hKey, strPath)

Do
    'Initialise bufffers
    lValueNameSize = 255
    strValueName = String$(lValueNameSize, " ")
    lDataBufferSize = 4000
    lRegResult = RegEnumValue(hCurKey, lCounter, strValueName, lValueNameSize, 0&, lValueType, byDataBuffer(0), lDataBufferSize)
    If lRegResult = 0 Then
        'Save the type
        ReDim Preserve strNames(lCounter) As String
        ReDim Preserve lTypes(lCounter) As Long
        lTypes(UBound(lTypes)) = lValueType
        'Tidy up string and save it
        intZeroPos = InStr(strValueName, Chr$(0))
        If intZeroPos > 0 Then
            strNames(UBound(strNames)) = _
            Left$(strValueName, intZeroPos - 1)
        Else
            strNames(UBound(strNames)) = strValueName
        End If
        lCounter = lCounter + 1
    Else
        Exit Do
    End If
Loop

'Move data into array
RegCloseKey hCurKey
ReDim KeyData(UBound(strNames), 0 To 1) As Variant

For lCounter = 0 To UBound(strNames)
    KeyData(lCounter, 0) = strNames(lCounter)
    KeyData(lCounter, 1) = lTypes(lCounter)
Next
'GetAllValues = Finisheddata
'Erase Finisheddata
End Sub

Private Function FindListItem(ObjectHwnd As Long, bListBox As Boolean, bExactMatch As Boolean, sCriteria As String) As Long
' Function checks listbox contents for match of sCriteria if bListBox = True, otherwise
' checks combobox contents for match of sCriteria if bListBox = False
Dim lMatchType As Long
If bListBox = True Then
    If bExactMatch = False Then lMatchType = LB_FINDSTRING Else lMatchType = LB_FINDSTRINGEXACT
Else
    If bExactMatch = False Then lMatchType = CB_FINDSTRING Else lMatchType = CB_FINDSTRINGEXACT
End If
FindListItem = SendMessageStr(ObjectHwnd, lMatchType, -1, sCriteria)
End Function



