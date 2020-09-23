Attribute VB_Name = "RegistryFunctions"
Option Explicit
' used to store ODBC data
Public Type ODBCinfo
    Name As String
    Type As String
End Type

' The form will add selected DSN info to this variable.
' If the user pressed cancel then the Name property will be vbNullString
Public NewDSNname As ODBCinfo
' To see if the user selected a different ODBC setting, compare
' If NewDSNname.Name = DSNname Then      < selected different setting
' If NewDSNname.Name = vbNullString            < user clicked cancel
' The type of ODBC connection will be contained in NewDSNname.Type

' The form will use this variable to populate currently selected ODBC string if you set the value
Public DSNname As String
' Suggest setting DSNname = NewDSNname.Name after form is closed.

'NOTE: This project does not test an ODBC connection once it has been selected. That is up to you.

