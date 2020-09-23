Attribute VB_Name = "OpenFile"
Option Explicit

'API function declarations for the open file with app function
Private Declare Function apiShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) _
    As Long

'***App Window Constants***
Public Const WIN_NORMAL = 1         'Open Normal
'Public Const WIN_MAX = 3            'Open Maximized
'Public Const WIN_MIN = 2            'Open Minimized
'AGP - I added the following two constants
Public Const WIN_EXPLORE = "Explore" 'Open with Explorer
Public Const WIN_OPEN = "Open"       'Open with a Folder View

'***Error Codes***
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&

Public Function fHandleFile(stFile As String, lShowHow As Long, Optional OpenType As String) As Long
    Dim lRet As Long, varTaskID As Variant
    Dim stRet As String
    
    If OpenType <> WIN_EXPLORE And OpenType <> WIN_OPEN Then OpenType = vbNullString
    
    'First try ShellExecute
    lRet = apiShellExecute(&O0, OpenType, stFile, vbNullString, vbNullString, lShowHow)
    
    If lRet > ERROR_SUCCESS Then
        stRet = vbNullString
        lRet = -1
    Else
        Select Case lRet
            Case ERROR_NO_ASSOC:
                'Try the OpenWith dialog
                varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & stFile, WIN_NORMAL)
                lRet = (varTaskID <> 0)
            Case ERROR_OUT_OF_MEM:
                stRet = "Error: Out of Memory/Resources. Couldn't Execute!"
            Case ERROR_FILE_NOT_FOUND:
                stRet = "Error: File not found.  Couldn't Execute!"
            Case ERROR_PATH_NOT_FOUND:
                stRet = "Error: Path not found. Couldn't Execute!"
            Case ERROR_BAD_FORMAT:
                stRet = "Error:  Bad File Format. Couldn't Execute!"
            Case Else:
        End Select
        If lRet <> -1 Then MsgBox stFile & vbCrLf & stRet, vbCritical, "Shell execute failed..."
    End If
    fHandleFile = lRet '& IIf(stRet = "", vbNullString, ", " & stRet)
End Function
