Attribute VB_Name = "mdlMain"
Public fArray() As Byte
Public FileLoaded As Boolean
Public Forms() As Form2
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Public PersonCount As Double
Public Type Person
    id As Integer
    name As String
    surname As String
    sex As Byte
    job As String
    address As String
    mail As String
    country As String
    fathername As String
    mothername As String
    ip As String
End Type
Public Persons() As Person
Public Function Exists(sFileName As String) As Boolean
    Exists = CStr(CBool(PathFileExists(sFileName)))
End Function

