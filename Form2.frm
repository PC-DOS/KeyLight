VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Key Light - PC-DOS Workshop - [Config]"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command7 
      Caption         =   "信息(&I)"
      Height          =   360
      Left            =   2505
      TabIndex        =   15
      Top             =   3045
      Width           =   780
   End
   Begin VB.CommandButton Command6 
      Caption         =   "復位(&R)"
      Height          =   360
      Left            =   1575
      TabIndex        =   14
      Top             =   3045
      Width           =   900
   End
   Begin VB.CommandButton Command5 
      Caption         =   "關閉應用程序(&E)"
      Height          =   360
      Left            =   30
      TabIndex        =   13
      Top             =   3045
      Width           =   1515
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   360
      Left            =   4425
      TabIndex        =   12
      Top             =   3045
      Width           =   1035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   360
      Left            =   3330
      TabIndex        =   11
      Top             =   3045
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Caption         =   "顏色"
      Height          =   570
      Left            =   30
      TabIndex        =   5
      Top             =   2415
      Width           =   5445
      Begin VB.CommandButton Command2 
         BackColor       =   &H00000000&
         Height          =   315
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   1755
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "按鈕關閉時的顏色:"
         Height          =   180
         Left            =   2970
         TabIndex        =   8
         Top             =   240
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "按鈕啟用時的顏色:"
         Height          =   180
         Left            =   135
         TabIndex        =   6
         Top             =   240
         Width           =   1530
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "透明度"
      Enabled         =   0   'False
      Height          =   660
      Left            =   255
      TabIndex        =   2
      Top             =   1680
      Width           =   5205
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   255
         LargeChange     =   10
         Left            =   150
         Max             =   255
         Min             =   155
         SmallChange     =   5
         TabIndex        =   3
         Top             =   285
         Value           =   255
         Width           =   4110
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "255"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   12
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4350
         TabIndex        =   4
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "程序窗口透明(&T)"
      Height          =   375
      Left            =   30
      TabIndex        =   1
      Top             =   1230
      Width           =   1755
   End
   Begin VB.CheckBox Check1 
      Caption         =   "程序總是保持在所有窗口前面(&K)"
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   825
      Value           =   1  'Checked
      Width           =   2985
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   105
      Picture         =   "Form2.frx":030A
      Top             =   60
      Width           =   720
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Key Light 應用程序設置中心"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   870
      TabIndex        =   10
      Top             =   225
      Width           =   4530
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type SETTINGS
lpEnableColor As Long
lpDisableColor As Long
lpTopMost As Long
lpTrans As Long
lpTransValue As Long
End Type
Dim uFlags As SETTINGS
Private Type KEYSTATE
dwCapsLock As Integer
dwNumLock As Integer
dwScrollLock As Integer
End Type
Private Enum KEYSTATERETVALUE
dwEnable = 1
dwDisable = 0
End Enum
Dim lpKeyState As KEYSTATE
Dim lpValue As KEYSTATERETVALUE
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Dim lpszCaptionNew As String
Private Const SC_MINIMIZE = &HF020&
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_ICONIC = WS_MINIMIZE
Const SC_ICON = SC_MINIMIZE
Const SC_TASKLIST = &HF130&
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Dim bCodeUse As Boolean
Private Const WS_CAPTION = &HC00000
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * 1024
End Type
Const SC_RESTORE = &HF120&
Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Dim lMeWinStyle As Long
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOOWNERZORDER = &H200
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SC_MOVE = &HF010&
Private Const SC_SIZE = &HF000&
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Const WS_EX_APPWINDOW = &H40000
Private Type WINDOWINFORMATION
hWindow As Long
hWindowDC As Long
hThreadProcess As Long
hThreadProcessID As Long
lpszCaption As String
lpszClassName As String
lpszThreadProcessName As String * 1024
lpszThreadProcessPath As String
lpszExe As String
lpszPath As String
End Type
Private Type WINDOWPARAM
bEnabled As Boolean
bHide As Boolean
bTrans As Boolean
bClosable As Boolean
bSizable As Boolean
bMinisizable As Boolean
bTop As Boolean
lpTransValue As Integer
End Type
Dim lpWindow As WINDOWINFORMATION
Dim lpWindowParam() As WINDOWPARAM
Dim lpCur As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Dim lpRtn As Long
Dim hWindow As Long
Dim lpLength As Long
Dim lpArray() As Byte
Dim lpArray2() As Byte
Dim lpBuff As String
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const LWA_COLORKEY = &H1
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&
Private Const WS_SYSMENU = &H80000
Private Const GWL_STYLE = (-16)
Private Const MF_BYCOMMAND = &H0
Private Const SC_CLOSE = &HF060&
Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Private Const MF_INSERT = &H0&
Private Const SC_MAXIMIZE = &HF030&
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Type WINDOWINFOBOXDATA
lpszCaption As String
lpszClass As String
lpszThread As String
lpszHandle As String
lpszDC As String
End Type
Dim dwWinInfo As WINDOWINFOBOXDATA
Private Type SmFontAttr
 FontName As String
 FontSize As Integer
 FontBod As Boolean
 FontItalic As Boolean
 FontUnderLine As Boolean
 FontStrikeou As Boolean
 FontColor As Long
 WinHwnd As Long
 End Type
 Dim M_GetFont As SmFontAttr
Private Const RESOURCETYPE_DISK = &H1
Private Const RESOURCETYPE_PRINT = &H2
Private Const NoError = 0
Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_NETHOOD = &H13
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_TEMPLATES = &H15
Private Const LF_FACESIZE = 32
Private Const MAX_PATH = 260
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_FIXEDPITCHONLY = &H4000&
Private Const CF_EFFECTS = &H100&
Private Const ITALIC_FONTTYPE = &H200
Private Const BOLD_FONTTYPE = &H100
Private Const CF_NOFACESEL = &H80000
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCALABLEONLY = &H20000
Private Const CF_SCREENFONTS = &H1
Private Const CF_SHOWHELP = &H4&
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Type CHOOSECOLOR
 lStructSize As Long
 hwndOwner As Long
 hInstance As Long
 rgbResult As Long
 lpCustColors As String
 flags As Long
 lCustData As Long
 lpfnHook As Long
 lpTemplateName As String
 End Type
Private Type OPENFILENAME
 lStructSize As Long
 hwndOwner As Long
 hInstance As Long
 lpstrFilter As String
 lpstrCustomFilter As String
 nMaxCustFilter As Long
 nFilterIndex As Long
 lpstrFile As String
 nMaxFile As Long
 lpstrFileTitle As String
 nMaxFileTitle As Long
 lpstrInitialDir As String
 lpstrTitle As String
 flags As Long
 nFileOffset As Integer
 nFileExtension As Integer
 lpstrDefExt As String
 lCustData As Long
 lpfnHook As Long
 lpTemplateName As String
 End Type
Private Type LOGFONT
 lfHeight As Long
 lfWidth As Long
 lfEscapement As Long
 lfOrientation As Long
 lfWeight As Long
 lfItalic As Byte
 lfUnderline As Byte
 lfStrikeOut As Byte
 lfCharSet As Byte
 lfOutPrecision As Byte
 lfClipPrecision As Byte
 lfQuality As Byte
 lfPitchAndFamily As Byte
 lfFaceName As String * LF_FACESIZE
 End Type
 Dim MyComputer As Long
Private Type CHOOSEFONT
 lStructSize As Long
 hwndOwner As Long
 hDC As Long
 lpLogFont As Long
 iPointSize As Long
 flags As Long
 rgbColors As Long
 lCustData As Long
 lpfnHook As Long
 lpTemplateName As String
 hInstance As Long
 lpszStyle As String
 nFontType As Integer
 MISSING_ALIGNMENT As Integer
 nSizeMin As Long
 nSizeMax As Long
 End Type
Private Type SHITEMID
 cb As Long
 abID() As Byte
 End Type
Private Type ITEMIDLIST
 mkid As SHITEMID
 End Type
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal Pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
 (ByVal hwndOwner As Long, ByVal nFolder As Long, _
 Pidl As ITEMIDLIST) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChooseFont As CHOOSEFONT) As Long
Private Declare Function WNetDisconnectDialog Lib "mpr.dll" _
 (ByVal hwnd As Long, ByVal dwType As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Type BROWSEINFO
 hOwner As Long
 pidlRoot As Long
 pszDisplayName As String
 lpszTitle As String
 ulFlags As Long
 lpfn As Long
 lParam As Long
 iImage As Long
 End Type
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Dim FontInfo As SmFontAttr
Private Function GetFolderValue(wIdx As Integer) As Long
 If wIdx < 2 Then
 GetFolderValue = 0
 ElseIf wIdx < 12 Then
 GetFolderValue = wIdx
 Else
 GetFolderValue = wIdx + 4
 End If
 End Function
Private Function GetReturnType() As Long
 Dim dwRtn As Long
 dwRtn = dwRtn Or BIF_RETURNONLYFSDIRS
 GetReturnType = dwRtn
 End Function
 Public Function GetFolder(Optional Title As String, _
 Optional hwnd As Long, _
 Optional FolderID As Long = 1) As String
 Dim Bi As BROWSEINFO
 Dim Pidl As Long
 Dim Folder As String
 Dim IDL As ITEMIDLIST
 Dim nFolder As Long
 Dim ReturnFol As String
 Dim Fid As Integer
 Fid = FolderID
 Folder = String$(255, Chr$(0))
 With Bi
 .hOwner = hwnd
 nFolder = GetFolderValue(Fid)
 If SHGetSpecialFolderLocation(ByVal hwnd, ByVal nFolder, IDL) = NoError Then
 .pidlRoot = IDL.mkid.cb
 End If
 .pszDisplayName = String$(MAX_PATH, Fid)
 If Len(Title) > 0 Then
 .lpszTitle = Title & Chr$(0)
 Else
 .lpszTitle = "请选择文件夹:" & Chr$(0)
 End If
 .ulFlags = GetReturnType()
 End With
 Pidl = SHBrowseForFolder(Bi)
 If SHGetPathFromIDList(ByVal Pidl, ByVal Folder) Then
 ReturnFol = Left$(Folder, InStr(Folder, Chr$(0)) - 1)
 If Right$(Trim$(ReturnFol), 1) <> "\" Then ReturnFol = ReturnFol & "\"
 GetFolder = ReturnFol
 Else
 GetFolder = ""
 End If
 End Function
 Public Function SaveFile(WinHwnd As Long, _
 Optional BoxLabel As String = "", _
 Optional StartPath As String = "", _
 Optional FilterStr = "*.*|*.*", _
 Optional Flag As Variant = &H4 Or &H200000) As String
 Dim rc As Long
 Dim pOpenfilename As OPENFILENAME
 Dim Fstr1() As String
 Dim Fstr As String
 Dim I As Long
 Const MAX_Buffer_LENGTH = 256
 On Error Resume Next
 If Len(Trim$(StartPath)) > 0 Then
 If Right$(StartPath, 1) <> "\" Then StartPath = StartPath & "\"
 If Dir$(StartPath, vbDirectory + vbHidden) = "" Then
 StartPath = App.Path
 End If
 Else
 StartPath = App.Path
 End If
 If Len(Trim$(FilterStr)) = 0 Then
 Fstr = "*.*|*.*"
 End If
 Fstr1 = Split(FilterStr, "|")
 For I = 0 To UBound(Fstr1)
 Fstr = Fstr & Fstr1(I) & vbNullChar
 Next
 With pOpenfilename
 .hwndOwner = WinHwnd
 .hInstance = App.hInstance
 .lpstrTitle = BoxLabel
 .lpstrInitialDir = StartPath
 .lpstrFilter = Fstr
 .nFilterIndex = 1
 .lpstrDefExt = vbNullChar & vbNullChar
 .lpstrFile = String(MAX_Buffer_LENGTH, 0)
 .nMaxFile = MAX_Buffer_LENGTH - 1
 .lpstrFileTitle = .lpstrFile
 .nMaxFileTitle = MAX_Buffer_LENGTH
 .lStructSize = Len(pOpenfilename)
 .flags = Flag
 End With
 rc = GetSaveFileName(pOpenfilename)
 If rc Then
 SaveFile = Left$(pOpenfilename.lpstrFile, pOpenfilename.nMaxFile)
 Else
 SaveFile = ""
 End If
 End Function
 Public Function OpenFile(WinHwnd As Long, _
 Optional BoxLabel As String = "", _
 Optional StartPath As String = "", _
 Optional FilterStr = "*.*|*.*", _
 Optional Flag As Variant = &H8 Or &H200000) As String
 Dim rc As Long
 Dim pOpenfilename As OPENFILENAME
 Dim Fstr1() As String
 Dim Fstr As String
 Dim I As Long
 Const MAX_Buffer_LENGTH = 256
 On Error Resume Next
 If Len(Trim$(StartPath)) > 0 Then
 If Right$(StartPath, 1) <> "\" Then StartPath = StartPath & "\"
 If Dir$(StartPath, vbDirectory + vbHidden) = "" Then
 StartPath = App.Path
 End If
 Else
 StartPath = App.Path
 End If
 If Len(Trim$(FilterStr)) = 0 Then
 Fstr = "*.*|*.*"
 End If
 Fstr = ""
 Fstr1 = Split(FilterStr, "|")
 For I = 0 To UBound(Fstr1)
 Fstr = Fstr & Fstr1(I) & vbNullChar
 Next
 With pOpenfilename
 .hwndOwner = WinHwnd
 .hInstance = App.hInstance
 .lpstrTitle = BoxLabel
 .lpstrInitialDir = StartPath
 .lpstrFilter = Fstr
 .nFilterIndex = 1
 .lpstrDefExt = vbNullChar & vbNullChar
 .lpstrFile = String(MAX_Buffer_LENGTH, 0)
 .nMaxFile = MAX_Buffer_LENGTH - 1
 .lpstrFileTitle = .lpstrFile
 .nMaxFileTitle = MAX_Buffer_LENGTH
 .lStructSize = Len(pOpenfilename)
 .flags = Flag
 End With
 rc = GetOpenFileName(pOpenfilename)
 If rc Then
 OpenFile = Left$(pOpenfilename.lpstrFile, pOpenfilename.nMaxFile)
 Else
 OpenFile = ""
 End If
 End Function
 Public Function GetColor() As Long
 Dim rc As Long
 Dim pChoosecolor As CHOOSECOLOR
 Dim CustomColor() As Byte
 With pChoosecolor
 .hwndOwner = 0
 .hInstance = App.hInstance
 .lpCustColors = StrConv(CustomColor, vbUnicode)
 .flags = 0
 .lStructSize = Len(pChoosecolor)
 End With
 rc = CHOOSECOLOR(pChoosecolor)
 If rc Then
 GetColor = pChoosecolor.rgbResult
 Else
 GetColor = -1
 End If
 End Function
 Public Function ConnectDisk(Optional hwnd As Long) As Long
 Dim rc As Long
 If IsNumeric(hwnd) Then
 rc = WNetConnectionDialog(hwnd, RESOURCETYPE_DISK)
 Else
 rc = WNetConnectionDialog(0, RESOURCETYPE_DISK)
 End If
 ConnectDisk = rc
 End Function
 Public Function ConnectPrint(Optional hwnd As Long) As Long
 Dim rc As Long
 If IsNumeric(hwnd) Then
 rc = WNetConnectionDialog(hwnd, RESOURCETYPE_PRINT)
 Else
 rc = WNetConnectionDialog(0, RESOURCETYPE_PRINT)
 End If
 End Function
 Public Function DisconnectDisk(Optional hwnd As Long) As Long
 Dim rc As Long
 If IsNumeric(hwnd) Then
 rc = WNetDisconnectDialog(hwnd, RESOURCETYPE_DISK)
 Else
 rc = WNetDisconnectDialog(0, RESOURCETYPE_DISK)
 End If
 End Function
 Public Function DisconnectPrint(Optional hwnd As Long) As Long
 Dim rc As Long
 If IsNumeric(hwnd) Then
 rc = WNetDisconnectDialog(hwnd, RESOURCETYPE_PRINT)
 Else
 rc = WNetDisconnectDialog(0, RESOURCETYPE_PRINT)
 End If
 End Function
 Private Function GetFont(WinHwnd As Long) As SmFontAttr
 Dim rc As Long
 Dim pChooseFont As CHOOSEFONT
 Dim pLogFont As LOGFONT
 With pLogFont
 .lfFaceName = StrConv(FontInfo.FontName, vbFromUnicode)
 .lfItalic = FontInfo.FontItalic
 .lfUnderline = FontInfo.FontUnderLine
 .lfStrikeOut = FontInfo.FontStrikeou
 End With
 With pChooseFont
 .hInstance = App.hInstance
 If IsNumeric(WinHwnd) Then .hwndOwner = WinHwnd
 .flags = CF_BOTH + CF_INITTOLOGFONTSTRUCT + CF_EFFECTS + CF_NOSCRIPTSEL
 If IsNumeric(FontInfo.FontSize) Then .iPointSize = FontInfo.FontSize * 10
 If FontInfo.FontBod Then .nFontType = .nFontType + BOLD_FONTTYPE
 If IsNumeric(FontInfo.FontColor) Then .rgbColors = FontInfo.FontColor
 .lStructSize = Len(pChooseFont)
 .lpLogFont = VarPtr(pLogFont)
 End With
 rc = CHOOSEFONT(pChooseFont)
 If rc Then
 FontInfo.FontName = StrConv(pLogFont.lfFaceName, vbUnicode)
 FontInfo.FontName = Left$(FontInfo.FontName, InStr(FontInfo.FontName, vbNullChar) - 1)
 With pChooseFont
 FontInfo.FontSize = .iPointSize / 10
 FontInfo.FontBod = (.nFontType And BOLD_FONTTYPE)
 FontInfo.FontItalic = (.nFontType And ITALIC_FONTTYPE)
 FontInfo.FontUnderLine = (pLogFont.lfUnderline)
 FontInfo.FontStrikeou = (pLogFont.lfStrikeOut)
 FontInfo.FontColor = .rgbColors
 End With
 End If
 GetFont = FontInfo
 End Function
Sub GetProcessName(ByVal processID As Long, szExeName As String, szPathName As String)
On Error Resume Next
Dim my As PROCESSENTRY32
Dim hProcessHandle As Long
Dim success As Long
Dim l As Long
l = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If l Then
my.dwSize = 1060
If (Process32First(l, my)) Then
Do
If my.th32ProcessID = processID Then
CloseHandle l
szExeName = Left$(my.szExeFile, InStr(1, my.szExeFile, Chr$(0)) - 1)
For l = Len(szExeName) To 1 Step -1
If Mid$(szExeName, l, 1) = "\" Then
Exit For
End If
Next l
szPathName = Left$(szExeName, l)
Exit Sub
End If
Loop Until (Process32Next(l, my) < 1)
End If
CloseHandle l
End If
End Sub
Private Sub DisableClose(hwnd As Long, Optional ByVal MDIChild As Boolean)
On Error Resume Next
Exit Sub
Dim hSysMenu As Long
Dim nCnt As Long
Dim cID As Long
hSysMenu = GetSystemMenu(hwnd, False)
If hSysMenu = 0 Then
Exit Sub
End If
nCnt = GetMenuItemCount(hSysMenu)
If MDIChild Then
cID = 3
Else
cID = 1
End If
If nCnt Then
RemoveMenu hSysMenu, nCnt - cID, MF_BYPOSITION Or MF_REMOVE
RemoveMenu hSysMenu, nCnt - cID - 1, MF_BYPOSITION Or MF_REMOVE
DrawMenuBar hwnd
End If
End Sub
Private Function GetPassword(hwnd As Long) As String
On Error Resume Next
lpLength = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
If lpLength > 0 Then
ReDim lpArray(lpLength + 1) As Byte
ReDim lpArray2(lpLength - 1) As Byte
CopyMemory lpArray(0), lpLength, 2
SendMessage hwnd, WM_GETTEXT, lpLength + 1, lpArray(0)
CopyMemory lpArray2(0), lpArray(0), lpLength
GetPassword = StrConv(lpArray2, vbUnicode)
Else
GetPassword = ""
End If
End Function
Private Function GetWindowClassName(hwnd As Long) As String
On Error Resume Next
Dim lpszWindowClassName As String * 256
lpszWindowClassName = Space(256)
GetClassName hwnd, lpszWindowClassName, 256
lpszWindowClassName = Trim(lpszWindowClassName)
GetWindowClassName = lpszWindowClassName
End Function
Private Sub Check1_Click()
On Error Resume Next
Exit Sub
Exit Sub
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
If 1 = 245 Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
End If
End Sub
Private Sub Check1_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
If 1 = 245 Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
End If
End Sub
Private Sub Check2_Click()
On Error Resume Next
Select Case Check2.Value
Case 0
With Frame1
.Enabled = False
End With
With HScroll1
.Max = 255
.Min = 155
.LargeChange = 10
.SmallChange = 5
.Enabled = False
End With
With Label1
.Alignment = 2
.Caption = HScroll1.Value
.Enabled = False
.BackStyle = 1
.BackColor = RGB(255, 255, 255)
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
Case 1
With Frame1
.Enabled = True
End With
With HScroll1
.Max = 255
.Min = 155
.LargeChange = 10
.SmallChange = 5
.Enabled = True
End With
With Label1
.Alignment = 2
.Caption = HScroll1.Value
.Enabled = True
.BackStyle = 1
.BackColor = RGB(255, 255, 255)
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
End Select
End Sub
Private Sub Check2_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
If 1 = 245 Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
End If
End Sub
Private Sub Command1_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
With Me
.Height = 3840
.Width = 5580
End With
Command1.BackColor = GetColor()
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
With Me
.Height = 3840
.Width = 5580
End With
With Me
.Height = 3840
.Width = 5580
End With
End Sub
Private Sub Command1_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
If 1 = 245 Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
End If
End Sub
Private Sub Command2_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
With Me
.Height = 3840
.Width = 5580
End With
Command2.BackColor = GetColor()
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
With Me
.Height = 3840
.Width = 5580
End With
With Me
.Height = 3840
.Width = 5580
End With
End Sub
Private Sub Command2_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
If 1 = 245 Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
End If
End Sub
Private Sub Command3_Click()
On Error Resume Next
Form1.Tag = ""
Select Case Check1.Value
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1410
.Width = 4350
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1410
.Width = 4350
End With
End Select
Dim rtn As Long
Select Case Check2.Value
Case 1
rtn = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Form1.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes Form1.hwnd, 0, Me.HScroll1.Value, LWA_ALPHA
Case 0
rtn = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Form1.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes Form1.hwnd, 0, 255, LWA_ALPHA
End Select
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
Me.Hide
Form1.IsSet = False
Form1.Tag = ""
Form1.SetFocus
End Sub
Private Sub Command3_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
If 1 = 245 Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
End If
End Sub
Private Sub Command4_Click()
On Error Resume Next
Form1.Tag = ""
If 1 = 245 Then
With uFlags
.lpDisableColor = CLng(GetSetting("KeyLight", "GUISettings", "KeyDisableColor", RGB(0, 245, 245)))
.lpEnableColor = CLng(GetSetting("KeyLight", "GUISettings", "KeyEnableColor", RGB(0, 0, 0)))
.lpTopMost = CLng(GetSetting("KeyLight", "GUISettings", "TopMost", 1))
.lpTrans = CLng(GetSetting("KeyLight", "GUISettings", "Trans", 0))
.lpTransValue = CLng(GetSetting("KeyLight", "GUISettings", "TransValue", 245))
End With
End If
With uFlags
.lpDisableColor = CLng(GetSetting("KeyLight", "GUISettings", "KeyDisableColor", RGB(0, 0, 0)))
.lpEnableColor = CLng(GetSetting("KeyLight", "GUISettings", "KeyEnableColor", RGB(0, 245, 245)))
.lpTopMost = CLng(GetSetting("KeyLight", "GUISettings", "TopMost", 1))
.lpTrans = CLng(GetSetting("KeyLight", "GUISettings", "Trans", 0))
.lpTransValue = CLng(GetSetting("KeyLight", "GUISettings", "TransValue", 245))
End With
With uFlags
Me.Check1.Value = .lpTopMost
Check2.Value = .lpTrans
With Me.HScroll1
.Max = 255
.Min = 155
.LargeChange = 10
.SmallChange = 5
.Value = uFlags.lpTransValue
End With
Command1.BackColor = .lpEnableColor
Command2.BackColor = .lpDisableColor
End With
On Error Resume Next
Select Case Check2.Value
Case 0
With Frame1
.Enabled = False
End With
With HScroll1
.Max = 255
.Min = 155
.LargeChange = 10
.SmallChange = 5
.Enabled = False
End With
With Label1
.Alignment = 2
.Caption = HScroll1.Value
.Enabled = False
.BackStyle = 1
.BackColor = RGB(255, 255, 255)
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
Case 1
With Frame1
.Enabled = True
End With
With HScroll1
.Max = 255
.Min = 155
.LargeChange = 10
.SmallChange = 5
.Enabled = True
End With
With Label1
.Alignment = 2
.Caption = HScroll1.Value
.Enabled = True
.BackStyle = 1
.BackColor = RGB(255, 255, 255)
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
End Select
With Me.Command4
.Cancel = True
End With
With Me.Command3
.Default = True
End With
Me.Hide
Form1.IsSet = False
Form1.Tag = ""
Form1.SetFocus
End Sub
Private Sub Command4_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
If 1 = 245 Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
End If
End Sub
Private Sub Command5_Click()
On Error Resume Next
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
Form1.cSysTray1.InTray = False
Unload Me
Form1.cSysTray1.InTray = False
Unload Form1
End
Form1.cSysTray1.InTray = False
Unload Me
Form1.cSysTray1.InTray = False
Unload Form1
End
End Sub
Private Sub Command5_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
If 1 = 245 Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
End If
End Sub
Private Sub Command6_Click()
On Error Resume Next
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Load Form2
If 1 = 245 Then
Form2.Hide
Form1.IsSet = False
End If
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Form1.Width, Form1.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
Dim ans As Integer
ans = MsgBox("確定初始化程序設置嗎?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(1)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(0)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(245)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(RGB(0, 245, 245))
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(RGB(0, 0, 0))
If 1 = 245 Then
Unload Form2
End If
Load Form2
On Error Resume Next
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
If 1 = 245 Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
End If
With uFlags
.lpDisableColor = CLng(GetSetting("KeyLight", "GUISettings", "KeyDisableColor", RGB(0, 0, 0)))
.lpEnableColor = CLng(GetSetting("KeyLight", "GUISettings", "KeyEnableColor", RGB(0, 245, 245)))
.lpTopMost = CLng(GetSetting("KeyLight", "GUISettings", "TopMost", 1))
.lpTrans = CLng(GetSetting("KeyLight", "GUISettings", "Trans", 0))
.lpTransValue = CLng(GetSetting("KeyLight", "GUISettings", "TransValue", 245))
End With
With uFlags
Me.Check1.Value = .lpTopMost
Check2.Value = .lpTrans
With Me.HScroll1
.Max = 255
.Min = 155
.LargeChange = 10
.SmallChange = 5
.Value = uFlags.lpTransValue
End With
Command1.BackColor = .lpEnableColor
Command2.BackColor = .lpDisableColor
End With
On Error Resume Next
Select Case Check2.Value
Case 0
With Frame1
.Enabled = False
End With
With HScroll1
.Max = 255
.Min = 155
.LargeChange = 10
.SmallChange = 5
.Enabled = False
End With
With Label1
.Alignment = 2
.Caption = HScroll1.Value
.Enabled = False
.BackStyle = 1
.BackColor = RGB(255, 255, 255)
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
Case 1
With Frame1
.Enabled = True
End With
With HScroll1
.Max = 255
.Min = 155
.LargeChange = 10
.SmallChange = 5
.Enabled = True
End With
With Label1
.Alignment = 2
.Caption = HScroll1.Value
.Enabled = True
.BackStyle = 1
.BackColor = RGB(255, 255, 255)
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
End Select
With Me.Command4
.Cancel = True
End With
With Me.Command3
.Default = True
End With
With Me
.Height = 3840
.Width = 5580
End With
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Form1.Width, Form1.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Form1.Width, Form1.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Form1.Width, Form1.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
End Select
Dim rtn As Long
Select Case Form2.Check2.Value
Case 1
rtn = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Form1.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes Form1.hwnd, 0, Form2.HScroll1.Value, LWA_ALPHA
Case 0
rtn = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Form1.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes Form1.hwnd, 0, 255, LWA_ALPHA
End Select
On Error Resume Next
With lpKeyState
.dwCapsLock = GetKeyState(vbKeyCapital)
.dwNumLock = GetKeyState(vbKeyNumlock)
.dwScrollLock = GetKeyState(vbKeyScrollLock)
End With
Select Case lpKeyState.dwCapsLock
Case 0
With Form1.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Form1.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Form1.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Form1.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Form1.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Form1.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Form1.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Form1.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Form1.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
With Form1
.Left = .Left
.Top = .Top
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
With Form1.Timer1
.Interval = 245 * 2
.Enabled = True
End With
With Form1.cSysTray1
.InTray = True
.TrayTip = "Key Light - PC-DOS Workshop"
End With
Else
Load Form2
If 1 = 245 Then
Form2.Hide
Form1.IsSet = False
End If
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Form1.Width, Form1.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Form1.Width, Form1.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Form1.Width, Form1.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
End Select
Select Case Form2.Check2.Value
Case 1
rtn = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Form1.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes Form1.hwnd, 0, Form2.HScroll1.Value, LWA_ALPHA
Case 0
rtn = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Form1.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes Form1.hwnd, 0, 255, LWA_ALPHA
End Select
On Error Resume Next
With lpKeyState
.dwCapsLock = GetKeyState(vbKeyCapital)
.dwNumLock = GetKeyState(vbKeyNumlock)
.dwScrollLock = GetKeyState(vbKeyScrollLock)
End With
Select Case lpKeyState.dwCapsLock
Case 0
With Form1.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Form1.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Form1.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Form1.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Form1.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Form1.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Form1.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Form1.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Form1.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
With Form1
.Left = .Left
.Top = .Top
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
With Form1.Timer1
.Interval = 245 * 2
.Enabled = True
End With
With Form1.cSysTray1
.InTray = True
.TrayTip = "Key Light - PC-DOS Workshop"
End With
Exit Sub
End If
End Sub
Private Sub Command6_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
If 1 = 245 Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
End If
End Sub
Private Sub Command7_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
With Me
.Height = 3840
.Width = 5580
End With
If 1 = 245 Then
With uFlags
.lpDisableColor = CLng(GetSetting("KeyLight", "GUISettings", "KeyDisableColor", RGB(0, 245, 245)))
.lpEnableColor = CLng(GetSetting("KeyLight", "GUISettings", "KeyEnableColor", RGB(0, 0, 0)))
.lpTopMost = CLng(GetSetting("KeyLight", "GUISettings", "TopMost", 1))
.lpTrans = CLng(GetSetting("KeyLight", "GUISettings", "Trans", 0))
.lpTransValue = CLng(GetSetting("KeyLight", "GUISettings", "TransValue", 245))
End With
End If
With uFlags
.lpDisableColor = CLng(GetSetting("KeyLight", "GUISettings", "KeyDisableColor", RGB(0, 0, 0)))
.lpEnableColor = CLng(GetSetting("KeyLight", "GUISettings", "KeyEnableColor", RGB(0, 245, 245)))
.lpTopMost = CLng(GetSetting("KeyLight", "GUISettings", "TopMost", 1))
.lpTrans = CLng(GetSetting("KeyLight", "GUISettings", "Trans", 0))
.lpTransValue = CLng(GetSetting("KeyLight", "GUISettings", "TransValue", 245))
End With
Dim lpMessage As String
With uFlags
lpMessage = "以下是當前保存在系統數據庫中的設置信息清單:" & vbCrLf & vbCrLf
lpMessage = lpMessage & "程序保持在全部窗口前面: " & CBool(.lpTopMost) & vbCrLf
lpMessage = lpMessage & "程序窗口透明: " & CBool(.lpTrans) & vbCrLf
lpMessage = lpMessage & "如果窗口透明,透明度是: " & .lpTransValue & vbCrLf
lpMessage = lpMessage & "開關按鈕啟用時十六進制顏色: " & Hex(.lpEnableColor) & vbCrLf
lpMessage = lpMessage & "開關按鈕關閉時十六進制顏色: " & Hex(.lpDisableColor)
MsgBox lpMessage, vbInformation, "Info"
End With
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
With Me
.Height = 3840
.Width = 5580
End With
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
With Me
.Height = 3840
.Width = 5580
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
With Me
.Height = 3840
.Width = 5580
End With
End Select
Dim rtn As Long
Select Case Form2.Check2.Value
Case 1
rtn = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Form1.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes Form1.hwnd, 0, Form2.HScroll1.Value, LWA_ALPHA
Case 0
rtn = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Form1.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes Form1.hwnd, 0, 255, LWA_ALPHA
End Select
On Error Resume Next
With lpKeyState
.dwCapsLock = GetKeyState(vbKeyCapital)
.dwNumLock = GetKeyState(vbKeyNumlock)
.dwScrollLock = GetKeyState(vbKeyScrollLock)
End With
End Sub
Private Sub Command7_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
If 1 = 245 Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
End If
End Sub
Private Sub Form_Activate()
On Error Resume Next
Form1.Tag = "Suicune"
Exit Sub
Exit Sub
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
If 1 = 245 Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
End If
End Sub
Private Sub Form_Deactivate()
Form1.Tag = ""
End Sub
Private Sub Form_GotFocus()
On Error Resume Next
Form1.Tag = "Set"
Exit Sub
Exit Sub
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
If 1 = 245 Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
End If
End Sub
Private Sub Form_Initialize()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
If 1 = 245 Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Form1.Tag = "Suicune"
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
If 1 = 245 Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
End If
With uFlags
.lpDisableColor = CLng(GetSetting("KeyLight", "GUISettings", "KeyDisableColor", RGB(0, 0, 0)))
.lpEnableColor = CLng(GetSetting("KeyLight", "GUISettings", "KeyEnableColor", RGB(0, 245, 245)))
.lpTopMost = CLng(GetSetting("KeyLight", "GUISettings", "TopMost", 1))
.lpTrans = CLng(GetSetting("KeyLight", "GUISettings", "Trans", 0))
.lpTransValue = CLng(GetSetting("KeyLight", "GUISettings", "TransValue", 245))
End With
With uFlags
Me.Check1.Value = .lpTopMost
Check2.Value = .lpTrans
With Me.HScroll1
.Max = 255
.Min = 155
.LargeChange = 10
.SmallChange = 5
.Value = uFlags.lpTransValue
End With
Command1.BackColor = .lpEnableColor
Command2.BackColor = .lpDisableColor
End With
On Error Resume Next
Select Case Check2.Value
Case 0
With Frame1
.Enabled = False
End With
With HScroll1
.Max = 255
.Min = 155
.LargeChange = 10
.SmallChange = 5
.Enabled = False
End With
With Label1
.Alignment = 2
.Caption = HScroll1.Value
.Enabled = False
.BackStyle = 1
.BackColor = RGB(255, 255, 255)
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
Case 1
With Frame1
.Enabled = True
End With
With HScroll1
.Max = 255
.Min = 155
.LargeChange = 10
.SmallChange = 5
.Enabled = True
End With
With Label1
.Alignment = 2
.Caption = HScroll1.Value
.Enabled = True
.BackStyle = 1
.BackColor = RGB(255, 255, 255)
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
End Select
With Me.Command4
.Cancel = True
End With
With Me.Command3
.Default = True
End With
With Me
.Height = 3840
.Width = 5580
End With
End Sub
Private Sub HScroll1_Change()
On Error Resume Next
With Me.Label1
.Alignment = 2
.BorderStyle = 1
.BackStyle = 1
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Caption = CStr(HScroll1.Value)
End With
End Sub
Private Sub HScroll1_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 3840
.Width = 5580
End With
If 1 = 245 Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
End If
End Sub
Private Sub Label1_Click()
On Error Resume Next
Dim lpTransValue As String
lpTransValue = InputBox$("請輸入透明度數值" & vbCrLf & "範圍:155到255", "Input Value", Me.HScroll1.Value)
If CInt(lpTransValue) <= 255 Then
If CInt(lpTransValue) >= 155 Then
With Me.HScroll1
.Max = 255
.Min = 155
.LargeChange = 10
.SmallChange = 5
.Value = CInt(lpTransValue)
End With
Else
MsgBox "輸入的數值無效", vbCritical, "Error"
Exit Sub
End If
Else
MsgBox "輸入的數值無效", vbCritical, "Error"
Exit Sub
End If
End Sub

