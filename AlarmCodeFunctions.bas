Attribute VB_Name = "AlarmCodeFunctions"
'This module is used to make API declarations for use with reading and writing to the AlarmAckMsg.ini file

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
                                                                                            ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
                                                                                                ByVal lpFileName As String) As Long
                                                                                                
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
                                                                                            ByVal lpFileName As String) As Long
Public FilePath As String
Private Const key As String = "AlarmType"
Private Const default As String = "GeneralAlarm"
Public alarm As String

Public Property Get path() As String
    path = FilePath
End Property

Public Property Get keyName() As String
    keyName = key
End Property

Public Property Get defaultSection() As String
    defaultSection = default
End Property

Public Function getKeyValue(fileName As String, SectionName As String, key As String, default As String) As String
    Dim strTemp As String
    Dim Length As Integer
    strTemp = Space$(255)
    Length = GetPrivateProfileString(SectionName, key, "", strTemp, 255, fileName)
    
    getKeyValue = Left$(strTemp, Length)

End Function

Public Function getSectionValues(fileName As String, SectionName As String) As String
    Dim strTemp As String
    Dim Length As Integer
    strTemp = Space$(32767)
    Length = GetPrivateProfileSection(SectionName, strTemp, 32767, fileName)
    
    getSectionValues = Left$(strTemp, Length)
    
End Function

Public Function SplitToArray(keyVal As String, delim As String) As String
    Dim arr() As String
    arr = Split(keyVal, delim)
    SplitToArray = arr(1)
    
End Function
