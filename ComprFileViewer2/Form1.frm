VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compressed File Viewer 2"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   8175
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Dump to Text File"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   7800
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdg1 
      Left            =   6000
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   7080
      ItemData        =   "Form1.frx":1CFA
      Left            =   120
      List            =   "Form1.frx":1CFC
      TabIndex        =   3
      Top             =   600
      Width           =   7935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go"
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Input your full file path or browse for it."
      Top             =   120
      Width           =   6375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents Archive As cArchive
Attribute Archive.VB_VarHelpID = -1

'I stripped Dana Seaman's FolderView code so that the result was only the code that
'enumerated compressed files. Although her FolderView code is excellent, the code that I
'needed was only to see the contents of compressed files (zip, cab, rar, and ace). Note
'that this code does not compress or uncompress files, it simply enumerates the contents
'of the compressed files.
'Dana 's original code is at:
'http://www.planet-source-code.com/xq/ASP/txtCodeId.23292/lngWId.1/qx/vb/scripts/ShowCode.htm
'All of the code here is Dana's except for the minor stuff like the "Dump to Text File"
'subroutine and the browsing function. You will notice that there are alot of variables that
'have been commented out. Those are just the byproducts of the original and bigger project.
'They were not needed for the compressed file enumeration but i left them in the code. Also,
'as a request from me, Dana tweaked the code to work with shared files over a network.
'I want to personally thank Dana for such good work.
'
'~El Mariachi
'
'ps The enumeration of zip and cab files is done all through the code in this project.
'   The enumeration of ace and rar files is done via the two attached DLL's. I had to
'   rename them as .dl_ but just rename them to .dll and it should work. I forget if these
'   DLL's are Dana's or from a third party. You might want to contact her for further
'   inquiries about this code and the DLL's.

Public Sub Archive_FileFound(ByVal Index As Long, ByVal Total As Long, ByVal FileName As String, _
                              ByVal ArchiveExt As String, ByVal Modified As Date, ByVal Size As Long, _
                              ByVal CompSize As Long, ByVal Method As Long, ByVal Attr As Long, _
                              ByVal Path As String, ByVal flags As Long, ByVal Crc As Long, _
                              ByVal Comments As String)
    Dim sMethod As String
    ', sExt As String, FakePath As String
    'Dim fType As String Long
    'Dim FakeFile
    'Dim MyIcon AsAs Integer
    Dim Ratio As Single
    Dim Encrypt As Boolean
    'Dim Item As ListItem

    On Error GoTo ProcedureError

    'Trap division by zero
    If Size Then
       Ratio = 1 - CompSize / Size
       'Don't allow negative values (per PkZip/WinZip)
       'Occurs on stored+encrypted files
       If Ratio < 0 Then Ratio = 0
    Else
       Ratio = 0
    End If
    'Ratio is single. Format as desired

    Select Case ArchiveExt
       Case ace_
          sMethod = MethodVerboseAce(Method, flags)
          Encrypt = (flags And 4) * -1
       Case cab_
          Select Case Method
             Case 0: sMethod = "None"
             Case 1: sMethod = "MsZip"
             Case 2: sMethod = "Lzx"
          End Select
          Encrypt = False
       Case rar_
          sMethod = MethodVerboseRar(Method, flags)
          'Flag bit 2 is Encryption True/False
          Encrypt = (flags And 4) * -1
       Case zip_
          sMethod = MethodVerboseZip(Method, flags)
          Encrypt = (flags And 1) * -1
    End Select
  
    Me.List1.AddItem "Total=" & Total
    Me.List1.AddItem "Index=" & Index
    Me.List1.AddItem "Path=" & Path
    Me.List1.AddItem "FileName=" & FileName
    Me.List1.AddItem "ArchiveExt=" & ArchiveExt
    Me.List1.AddItem "Modified=" & Modified
    Me.List1.AddItem "Size=" & Size
    Me.List1.AddItem "CompSize=" & CompSize
    Me.List1.AddItem "Ratio=" & Format$(Ratio, "00.0%")
    Me.List1.AddItem "Method=" & Method
    Me.List1.AddItem "sMethod=" & sMethod
    Me.List1.AddItem "Encrypt=" & Encrypt
    Me.List1.AddItem "flags=" & flags
    Me.List1.AddItem "Attr=" & Attr
    Me.List1.AddItem "..Attr=" & GetAttrString(Attr)
    Me.List1.AddItem "Crc=" & Crc
    Me.List1.AddItem "hexCrc=" & Hex$(Crc)
    Me.List1.AddItem "Comments=" & Comments

    Me.List1.AddItem "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    Me.List1.AddItem ""

ProcedureExit:
    Exit Sub
ProcedureError:
    If ErrMsgBox(Me.Name & "Archive_FileFound") = vbRetry Then Resume Next
End Sub

Private Sub Command1_Click()
    'Open an archive
    
    On Error Resume Next
    Me.List1.Clear
    cdg1.DialogTitle = "Browse to the Compressed File..."
    cdg1.FileName = ""
    cdg1.InitDir = App.Path
    cdg1.Filter = "Zip Files (*.zip)|*.zip|Cab Files (*.cab)|*.cab|RAR Files (*.rar)|*.rar|Ace Files (*.ace)|*.ace|All compressed files|*.zip;*.cab;*.rar;*.ace|All Files (*.*)|*.*"
    cdg1.ShowOpen
    'Check if cancel was pressed
    If Err = cdlCancel Then Exit Sub
    
    Me.Text1.Text = cdg1.FileName
End Sub

Private Sub Command2_Click()
    Dim Path As String, sExt As String
    Dim tmi As Variant
    
    Me.List1.Clear
    Path = Trim$(Me.Text1.Text) 'the full path and file name
    If LenB(Dir(Path)) = 0 Then
        MsgBox "File does not exist!", vbCritical, "Error reading file"
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    Command1.Enabled = False
    Command2.Enabled = False
    sExt = GetExt(Path) 'the extension
    tmi = Now
    Select Case sExt
        Case ace_, cab_, rar_, zip_
            'InZip = True
            'LoadStart
            Set Archive = New cArchive
            Archive.ArchiveName = Path
            Archive.ArchiveExt = sExt
            Archive.GetInfo
            'LoadCleanup 1
            'ShowProgress Start, Archive.FileCount, Path
        Case Else
            'do nothing
    End Select
   
    Me.Caption = "Done in " & DateDiff("s", tmi, Now) & " seconds"
    Command1.Enabled = True
    Command2.Enabled = True
    Me.MousePointer = vbNormal
End Sub


Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    Dim i As Long
    
    Me.MousePointer = vbHourglass
    Command1.Enabled = False
    Command2.Enabled = False
    
    Open App.Path & "\" & "cfv2.dat" For Output As #23
    For i = 0 To List1.ListCount - 1
        Print #23, List1.List(i)
    Next i
    Close #23
    Command1.Enabled = True
    Command2.Enabled = True
    Me.MousePointer = vbNormal
    If MsgBox("Do you want to view the file?", vbYesNo + vbQuestion, "View It?") = vbYes Then
        fHandleFile App.Path & "\" & "cfv2.dat", WIN_NORMAL
    End If
End Sub
