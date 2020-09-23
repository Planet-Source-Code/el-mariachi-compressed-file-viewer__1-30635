Attribute VB_Name = "mDeclares"
Option Explicit

'Public Const HKEY_CLASSES_ROOT = &H80000000
'Public Const KEY_ALL_ACCESS = &H2003F
'Public Const hNull& = 0
Public Const MAX_PATH = 260
'Public Const NOERROR = 0
'Public Const CourierWhite As String = "<b><font face='courier new'   color='white' size='2'>"
'Public Const CourierBlack As String = "<b><font face='courier new'   color='black' size='2'>"
' Difference between day zero for VB dates and Win32 dates
' (or #12-30-1899# - #01-01-1601#)
Private Const rDayZeroBias As Double = 109205#    ' Abs(CDbl(#01-01-1601#))
' 10000000 nanoseconds * 60 seconds * 60 minutes * 24 hours / 10000
' comes to 86400000 (the 10000 adjusts for fixed point in Currency)
Private Const rMillisecondPerDay As Double = 10000000# * 60# * 60# * 24# / 10000#
'Public Const INVALID_HANDLE_VALUE = -1
'Public Const LVM_FIRST = &H1000
'Public Const LVS_SHAREIMAGELISTS = &H40&
'Public Const GWL_STYLE = (-16)
'Public Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
'Public Const LVSIL_NORMAL = 0
'Public Const LVSIL_SMALL = 1
'Public Const LVIF_IMAGE = &H2
'Public Const LVM_SETITEM = (LVM_FIRST + 6)

'Public Const LARGE_ICON As Integer = 32
'Public Const SMALL_ICON As Integer = 16
'Public Const ILD_TRANSPARENT = &H1                                     'Display transparent
'ShellInfo Flags
'Public Const SHGFI_DISPLAYNAME = &H200
'Public Const SHGFI_EXETYPE = &H2000
'Public Const SHGFI_SYSICONINDEX = &H4000                               'System icon index
'Public Const SHGFI_LARGEICON = &H0                                     'Large icon
'Public Const SHGFI_SMALLICON = &H1                                     'Small icon
'Public Const SHGFI_SHELLICONSIZE = &H4
'Public Const SHGFI_TYPENAME = &H400
'Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
'        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
'        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
'Public Const SMALLSYS_SHGFI_FLAGS = SHGFI_SYSICONINDEX Or SHGFI_SMALLICON

'Public Const YMDHMS As String = "yyyymmddhhnnss"
'Public Const HMS As String = "hhnnss"
'----------------------------------
'Public Const chk_ As String = "chk"
'Public Const nam_ As String = "nam"
'Public Const ext_ As String = "ext"
'Public Const siz_ As String = "siz"
'Public Const typ_ As String = "typ"
'Public Const mod_ As String = "mod"
'Public Const tim_ As String = "tim"
'Public Const cre_ As String = "cre"
'Public Const acc_ As String = "acc"
'Public Const atr_ As String = "atr"
'Public Const dos_ As String = "dos"
'----------------------------------
'Public Const cmp_ As String = "cmp"
'Public Const rat_ As String = "rat"
'Public Const crc_ As String = "crc"
'Public Const enc_ As String = "enc"
'Public Const mtd_ As String = "mtd"
'Public Const pth_ As String = "pth"
'Public Const com_ As String = "com"
'Public Const sig_ As String = "sig"
'----------------------------------
Public Const ace_ As String = "ace"
Public Const cab_ As String = "cab"
Public Const rar_ As String = "rar"
Public Const zip_ As String = "zip"
'----------------------------------
'Public Buffer As String * MAX_PATH
'Public f_Type As String * 80
'Public Lang   As Long
'Type SHFILEINFO
'        hIcon As Long                      '  out: icon
'        iIcon As Long                      '  out: icon index
'        dwAttributes As Long               '  out: SFGAO_ flags
'        szDisplayName As String * MAX_PATH '  out: display name (or path)
'        szTypeName As String * 80          '  out: type name
'End Type
'Public SFI As SHFILEINFO
'Public Const cbSFI As Long = 12 + MAX_PATH + 80 'size of SFI

'Public Type LV_ITEM
'    mask As Long
'    iItem As Long
'    iSubItem As Long
'    state As Long
'    stateMask As Long
'    pszText As String
'    cchTextMax As Long
'    iImage As Long
'    lParam As Long '(~ ItemData)
''#if (_WIN32_IE >= 0x0300)
'    iIndent As Long
''#End If
'End Type

   'Public lvi As LV_ITEM 'api list item struc


'Modified for Faster Date Conversion & 64-bit NTFS Filesizes
'Public Type WIN32_FIND_DATA
'   dwFileAttributes  As Long
'   '-------------------------------------------------------
'   'Example: MyDate = UTCCurrToVbDate(W32FD.ftLastWriteTime)
'   ftCreationTime    As Currency   'As FILETIME
'   ftLastAccessTime  As Currency   'As FILETIME
'   ftLastWriteTime   As Currency   'As FILETIME
'   '---------------------------------------------------------------
'   'Example: MySize = CVC(W32FD.nFileSizeLow & W32FD.nFileSizeHigh)
'   nFileSizeHigh     As String * 4 'As Long
'   nFileSizeLow      As String * 4 'As Long
'   '---------------------------------------
'   dwReserved0       As Long
'   dwReserved1       As Long
'   cFileName         As String * MAX_PATH
'   cAlternate        As String * 14
'End Type

'Public Type FTs
'   Ext As String
'   Type As String
'   IconIndex As Long
'End Type

'Public Enum SHFolders
'    CSIDL_DESKTOP = &H0
'    CSIDL_INTERNET = &H1
'    CSIDL_PROGRAMS = &H2
'    CSIDL_CONTROLS = &H3
'    CSIDL_PRINTERS = &H4
'    CSIDL_PERSONAL = &H5
'    CSIDL_FAVORITES = &H6
'    CSIDL_STARTUP = &H7
'    CSIDL_RECENT = &H8
'    CSIDL_SENDTO = &H9
'    CSIDL_BITBUCKET = &HA
'    CSIDL_STARTMENU = &HB
'    CSIDL_DESKTOPDIRECTORY = &H10
'    CSIDL_DRIVES = &H11
'    CSIDL_NETWORK = &H12
'    CSIDL_NETHOOD = &H13
'    CSIDL_FONTS = &H14
'    CSIDL_TEMPLATES = &H15
'    CSIDL_COMMON_STARTMENU = &H16
'    CSIDL_COMMON_PROGRAMS = &H17
'    CSIDL_COMMON_STARTUP = &H18
'    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
'    CSIDL_APPDATA = &H1A
'    CSIDL_PRINTHOOD = &H1B
'    CSIDL_ALTSTARTUP = &H1D '// DBCS
'    CSIDL_COMMON_ALTSTARTUP = &H1E '// DBCS
'    CSIDL_COMMON_FAVORITES = &H1F
'    CSIDL_INTERNET_CACHE = &H20
'    CSIDL_COOKIES = &H21
'    CSIDL_HISTORY = &H22
'End Enum
'NOTE!! Some declares changed to 'As Any' to
'       accomodate Currency as well as Filetime
Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As Any) As Long
'Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
'Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
'Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
'Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
'Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
'Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
'Declare Function GetLogicalDrives Lib "kernel32" () As Long
Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Any, lpLocalFileTime As Any) As Long
'Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
'Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal uID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
'Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
'Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
'Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
'Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal flags&) As Long
'Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Declare Function GetTickCount Lib "kernel32" () As Long
'Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
'Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'Declare Function CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As Long
'Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As Long
'Declare Function SHGetMalloc Lib "shell32" (ppMalloc As IMalloc) As Long
'Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Declare Function lstrlenptr Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
'Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Sub CopyMemoryLpToStr Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal lpvDest As String, lpvSource As Long, ByVal cbCopy As Long)
'Declare Function AlphaBlending Lib "Alphablending.dll" _
            (ByVal destHDC As Long, ByVal XDest As Long, ByVal YDest As Long, _
            ByVal destWidth As Long, ByVal destHeight As Long, ByVal srcHDC As Long, _
            ByVal xSrc As Long, ByVal ySrc As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal AlphaSource As Long) As Long
'Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
'Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
'Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

'Public Sub Blend(Destination As Long, Source As Long, Amount As Long, X, Y, X2, Y2)
'    AlphaBlending Destination, X, Y, X2, Y2, Source, X, Y, X2, Y2, Amount
'End Sub


Public Function PointerToString(lPtr As Long) As String
Dim lLen As Long
Dim sR As String
    ' Get length of Unicode string to first null
    lLen = lstrlenptr(lPtr)
    ' Allocate a string of that length
    sR = String$(lLen, 0)
    ' Copy the pointer data to the string
    CopyMemoryLpToStr sR, ByVal lPtr, lLen
    PointerToString = sR
End Function
'Public Function QualifyPath(ByVal MyString As String) As String
'   If Right$(MyString, 1) <> "\" Then
'      QualifyPath = MyString & "\"
'   Else
'      QualifyPath = MyString
'   End If
'End Function
Public Function GetMyDate(ZipDate As Integer, ZipTime As Integer) As Date
    Dim FTime As Currency 'Makes it much easier to convert
    'Convert the dos stamp into a file time
    DosDateTimeToFileTime CLng(ZipDate), CLng(ZipTime), FTime
    'Filetime to VbDate
    GetMyDate = UTCCurrToVbDate(FTime, False)
End Function
'Public Function GetResourceStringFromFile(sModule As String, idString) As String
'
'   Dim hModule As Long
'   Dim nChars As Long
'
'   hModule = LoadLibrary(sModule)
'   If hModule Then
'      nChars = LoadString(hModule, idString, Buffer, MAX_PATH)
'      If nChars Then
'         GetResourceStringFromFile = Left$(Buffer, nChars)
'      End If
'      FreeLibrary hModule
'   End If
'End Function

'Public Function GetResourceString(Num) As String
'   On Error Resume Next
'   Select Case Num
'      Case 1000 To 1999
'         GetResourceString = LoadResString(Lang + Num)
'      Case Else
'         GetResourceString = GetResourceStringFromFile("Shell32.Dll", Num)
'   End Select
'End Function

Public Function ErrMsgBox(Msg As String) As Integer
    ErrMsgBox = MsgBox("Error: " & Err.Number & ". " & Err.Description, vbRetryCancel + vbCritical, Msg)
End Function

Public Function UTCCurrToVbDate(ByVal MyCurr As Currency, Optional ToLocal As Boolean = True) As Date
   Dim UTC As Currency
   ' Discrepancy in WIN32_FIND_DATA:
   ' Win2000 correctly reports 0 as 01-01-1980, Win98/ME does not.
   If MyCurr = 0 Then MyCurr = 11960017200000# ' 01-01-1980
   If ToLocal Then
      FileTimeToLocalFileTime MyCurr, UTC
   Else
      UTC = MyCurr
   End If
   UTCCurrToVbDate = (UTC / rMillisecondPerDay) - rDayZeroBias

End Function
'Public Function CVC(ValToConvert As String) As Currency
'    'Converts 8-byte string to Currency
'    'NOTE: Stores value as 64-bit integer (up to 8 Exabytes - 1)
'    '      Must supply 4 byte Low & 4 byte High in that order.
'    '      Scale * 10000 when retrieving value
'    CopyMemory CVC, ByVal ValToConvert, 8
'End Function

'Public Function DirSpace(sPath As String) As Currency
'   Dim Win32Fd As WIN32_FIND_DATA
'   Dim lHandle As Long
'   Const FILE_ATTRIBUTE_DIRECTORY = &H10
'   sPath = QualifyPath(sPath)
'   lHandle = FindFirstFile(sPath & "*.*", Win32Fd)
'   If lHandle > 0 Then
'      Do
'         If Asc(Win32Fd.cFileName) <> 46 Then  'skip . and .. entries
'            If (Win32Fd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
'               DirSpace = DirSpace + CVC(Win32Fd.nFileSizeLow & Win32Fd.nFileSizeHigh)
'            Else 'Recurse
'               DirSpace = DirSpace + DirSpace(sPath & StripNull(Win32Fd.cFileName))
'            End If
'         End If
'      Loop While FindNextFile(lHandle, Win32Fd) > 0
'   End If
'   FindClose (lHandle)
'
'End Function

Public Sub ParseFullPath(ByVal FullPath As String, JustPath As String, JustName As String)
   
   Dim lSlash As Integer
   
   ' Given a full path, parse it and return
   ' the path and file name.
   lSlash = InStrRev(FullPath, "/")
   If lSlash = 0 Then
      lSlash = InStrRev(FullPath, "\")
   End If
   If lSlash > 0 Then
      JustName = Mid$(FullPath, lSlash + 1)
      JustPath = Left$(FullPath, lSlash)
   Else
      JustName = FullPath
      JustPath = vbNullString
   End If

End Sub
Public Function StringToPointer(sStr As String, ByRef ByteArray() As Byte) As Long
    Dim X As Long
    Dim lstrlen As Long
    
    lstrlen = Len(sStr)
    For X = 1 To lstrlen
        ByteArray(X - 1) = AscB(Mid$(sStr, X, 1))
    Next
    ByteArray(X - 1) = 0
    StringToPointer = VarPtr(ByteArray(LBound(ByteArray)))
End Function
Public Function StripNull(ByVal StrIn As String) As String
   On Error GoTo ProcedureError
   Dim nul As Long
   '
   ' Truncate input string at first null.
   ' If no nulls, perform ordinary Trim.
   '
   nul = InStr(StrIn, vbNullChar)
   Select Case nul
      Case Is > 1
         StripNull = Left$(StrIn, nul - 1)
      Case 1
         StripNull = vbNullString
      Case 0
         StripNull = Trim$(StrIn)
   End Select

ProcedureExit:
  Exit Function
ProcedureError:
     If ErrMsgBox("mDeclares.StripNull") = vbRetry Then Resume Next


End Function
'Public Function FolderLocation(lFolder As SHFolders) As String
'
'   Dim lp As Long
'   'Get the PIDL for this folder
'   SHGetSpecialFolderLocation 0&, lFolder, lp
'   SHGetPathFromIDList lp, Buffer
'   FolderLocation = StripNull(Buffer)
'   'Free the PIDL
'   CoTaskMemFree lp
'
'End Function
'Public Function UTCCurrToVbLocal(ByVal MyCurr As Currency) As Date
'   Dim UTC As Currency
'   ' Discrepancy in WIN32_FIND_DATA:
'   ' Win2000 correctly reports 0 as 01-01-1980, Win98/ME does not.
'   If MyCurr = 0 Then MyCurr = 11960017200000# ' 01-01-1980
'   If FileTimeToLocalFileTime(MyCurr, UTC) Then
'      UTCCurrToVbLocal = (UTC / rMillisecondPerDay) - rDayZeroBias
'   End If
'End Function
'Public Function GenDate(MyDate As Date, Optional JustDate As Boolean = False) As String
'
'   If JustDate Then
'      GenDate = FormatDateTime(MyDate, vbShortDate)
'   Else
'      GenDate = FormatDateTime(MyDate, vbGeneralDate)
'   End If
'
'End Function
'Public Function FileExistsW32FD(sSource As String) As WIN32_FIND_DATA
'
'   Dim hFile As Long
'   'Returns True in dwReserved1 if file exists as well as
'   'raw data in WIN32_FIND_DATA structure
'   hFile = FindFirstFile(sSource, FileExistsW32FD)
'   FileExistsW32FD.dwReserved1 = hFile <> INVALID_HANDLE_VALUE
'   FindClose hFile
'
'End Function
'Public Function FormatSize(ByVal Size As Variant, Optional ByVal ReturnBytes As Boolean = False) As String
''Copyright Dana Seaman
''dgs@natallink.com.br
''Handles up to 999.9 Yottabytes.
''Formats are:
''   #.###
''   ##.##
''   ###.#
'   Dim Decimals As Integer, Group As Integer, Pwr As Integer
'   Dim SizeKb
'   Const KB& = 1024
'   Const Letter$ = "KMGTPEZY"
'   On Error GoTo PROC_ERR
'
'   If Size < KB Then 'Return bytes
'      FormatSize = FormatNumber(Size, 0) & " b"
'      ' Vb5 FormatSize = Format$(Size, "#,##0 b")
'   Else
'      SizeKb = Size / KB
'      For Pwr = 0 To 23
'         If SizeKb < 10 ^ (Pwr + 1) Then    ' Fits our criteria
'            Group = Pwr \ 3                 ' Kb(0), Mb(1), etc.
'            SizeKb = SizeKb / KB ^ Group    ' Scale to group
'            Decimals = 4 - Len(Int(SizeKb)) ' NumDigitsAfterDecimal
'            FormatSize = FormatNumber(SizeKb, Decimals) & " " & _
'                         Mid(Letter, Group + 1, 1) & "b"
'            'Vb5 FormatSize = Format$(SizeKb, "#,###." & String(Decimals, 48)) & " " & _
'                         Mid(Letter, Group + 1, 1) & "b"
'
'            Exit For
'         End If
'      Next
'      If FormatSize = "" Then FormatSize = "Out of bounds"
'   End If
'
'   ' return bytes
'   If ReturnBytes Then
'      If Size >= KB Then
'         FormatSize = FormatSize & " (" & FormatNumber(Size, 0) & " b)"
'         'Vb5 FormatSize = FormatSize & Format$(Size, " (#,### b)")
'      End If
'   End If
'
'PROC_EXIT:
'  Exit Function
'PROC_ERR:
'   FormatSize = "Overflow"
'   'was If ErrMsgBox("Module1.FormatSize") = vbRetry Then Resume Next
'
'End Function


'***extra stuff************
Public Function GetExt(ByVal Name As String) As String
   On Error GoTo ProcedureError
   Dim j As Integer
   j = InStrRev(Name, ".")
   If j > 0 And j < (Len(Name)) Then
      GetExt = Mid$(Name, j + 1)
      GetExt = LCase$(GetExt)
   End If

ProcedureExit:
  Exit Function
ProcedureError:
     If ErrMsgBox(".GetExt") = vbRetry Then Resume Next

End Function

Public Function GetAttrString(ByVal Attr As Variant) As String
   On Error GoTo ProcedureError
   Dim j As Integer
   
   'Const sFill As String = "..............."
   Const sFill As String = "               "
   Const sAttr As String = "rhsvdalnt?lco?e"

'00 r   0001  "Read Only"
'01 h   0002  "Hidden"
'02 s   0004  "System"
'03 v   0008  "Volume Label"
'04 f   0016  "Folder"
'05 a   0032  "Archive"
'06 l   0064  "Alias"
'07 n   0128  "Normal"
'08 t   0256  "Temporary"
'09 ?   0512   ??
'10 l   1024  "Alias"
'11 c   2048  "Compressed"
'12 o   4096  "Offline"
'13 ?   8192   ??
'14 e  16384  "Encrypted"

   GetAttrString = sFill
   If Attr Then
      For j = 0 To 14
         If Attr And (2 ^ j) Then ' Set letter
            Mid$(GetAttrString, j + 1, 1) = Mid$(sAttr, j + 1, 1)
         End If
      Next
  End If
  GetAttrString = UCase$(Trim$(GetAttrString))

ProcedureExit:
  Exit Function
ProcedureError:
     If ErrMsgBox(".GetAttrString") = vbRetry Then Resume Next

End Function

Public Function MethodVerboseZip(ByVal Method, ByVal BitFlag) As String
   On Error Resume Next
   'Conforms to PkZip 2.04g Specifications
'Methods are
'0    Stored (None)
'1    Shrunk
'2-5  Reduced:1,2,3,4
'(For Method 6 - Imploding)
' general purpose bit flag: (2 bytes)
'Bit 1: If the compression method used was type 6,
'       Imploding, then this bit, if set, indicates
'       an 8K sliding dictionary was used.  If clear,
'       then a 4K sliding dictionary was used.
'Bit 2: If the compression method used was type 6,
'       Imploding, then this bit, if set, indicates
'       3 Shannon-Fano trees were used to encode the
'       sliding dictionary output.  If clear, then 2
'       Shannon-Fano trees were used.
'6    Imploded:8kDict/4kDict:3Tree/2Tree
'7    Tokenized
'(For Method 8 - Deflating)
' general purpose bit flag: (2 bytes)
'Bit 2  Bit 1
'  0      0    Normal (-en) compression option was used.
'  0      1    Maximum (-ex) compression option was used.
'  1      0    Fast (-ef) compression option was used.
'  1      1    Super Fast (-es) compression option was used.
'8    Deflated:N,X,F,S
'9    EnhDefl
'10   ImplDCL
   
   BitFlag = (BitFlag \ 2) And 3 'Isolate bits 1, 2

   Select Case Method
      'Since deflated is the most common check for it first
      Case 8
         MethodVerboseZip = "Deflated:" & Choose(BitFlag + 1, "N", "X", "F", "S")
      Case 0
         MethodVerboseZip = "Stored"
      Case 1
         MethodVerboseZip = "Shrunk"
      Case 2 To 5
         MethodVerboseZip = "Reduced:" & Method - 1
      Case 6
         MethodVerboseZip = "Imploded:" & Choose(BitFlag + 1, "8KDict:2Tree", "4KDict:2Tree", "8KDict:3Tree", "4KDict:3Tree")
      Case 7
         MethodVerboseZip = "Tokenized"
      Case 9
         MethodVerboseZip = "EnhDef"
      Case 10
         MethodVerboseZip = "ImplDCL"
      Case Else
         MethodVerboseZip = "Unknown"
   End Select

End Function

Public Function MethodVerboseRar(ByVal Method As Long, ByVal BitFlag As Long) As String
   On Error Resume Next
   Dim Dict As Integer
  'Flags
  '      0 0 0 0 0 0 0 0  &H00&   - dictionary size    64 KB
  '      0 0 1 0 0 0 0 0  &H20&   - dictionary size   128 KB
  '      0 1 0 0 0 0 0 0  &H40&   - dictionary size   256 KB
  '      0 1 1 0 0 0 0 0  &H60&   - dictionary size   512 KB
  '      1 0 0 0 0 0 0 0  &H80&   - dictionary size  1024 KB
  Dict = 2 ^ (6 + ((BitFlag \ 32) And 7)) 'Isolate bits 5 to 7
   Select Case Method
      Case 48
         MethodVerboseRar = "Stored"
      Case 51
         MethodVerboseRar = "Deflated:" & Dict & "Kb"
      Case Else
         MethodVerboseRar = "Unknown"
   End Select

End Function

Public Function MethodVerboseAce(ByVal Method, ByVal BitFlag) As String
   On Error Resume Next
   Dim Dict As Integer

   Dict = 1024 'Need to confirm this!!!
   Select Case Method
      Case 0
         MethodVerboseAce = "Stored"
      Case 1
         MethodVerboseAce = "Deflated:" & Dict & "Kb"
      Case Else
         MethodVerboseAce = "Unknown"
   End Select
End Function
