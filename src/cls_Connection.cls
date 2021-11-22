VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Connection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
 
Private m_ConfigFile As String

Public Property Get ConfigFile() As String: ConfigFile = m_ConfigFile: End Property
Public Property Let ConfigFile(ByVal p_ConfigFile As String): m_ConfigFile = p_ConfigFile: End Property

 
Public Function GetINIString(ByVal sApp As String, ByVal sKey As String) As String
    Dim sBuf As String * 256
    Dim lBuf As Long
    
    lBuf = GetPrivateProfileString(sApp, sKey, "", sBuf, Len(sBuf), m_ConfigFile)
    GetINIString = Left$(sBuf, lBuf)
End Function
 
 
Public Function WriteINI(ByVal sApp As String, ByVal sKey As String, ByVal sValue As String) As String
    WritePrivateProfileString sApp, sKey, sValue, m_ConfigFile
End Function

