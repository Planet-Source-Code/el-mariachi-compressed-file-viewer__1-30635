Attribute VB_Name = "NetworkAccess"
Option Explicit

' public Constants
Public Const INVALID_HANDLE_VALUE = -1
'
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
'Public Const OPEN_ALWAYS = 4
'Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
'Public Const GENERIC_WRITE = &H40000000
Public Const GENERIC_READ = &H80000000
'
Public Const FILE_BEGIN = 0
'Public Const FILE_CURRENT = 1
'Public Const FILE_END = 2

Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
                         ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, _
                         ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, _
                         lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, _
                        ByVal dwMoveMethod As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

