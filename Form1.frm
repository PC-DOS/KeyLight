VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1230
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   3525
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Left            =   1155
      Top             =   375
   End
   Begin 工程1.cSysTray cSysTray1 
      Left            =   1920
      Top             =   1305
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "Form1.frx":030A
      TrayTip         =   "KeyLight - PC-DOS Workshop"
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1575
      Top             =   1320
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ScrollLock"
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   2310
      TabIndex        =   2
      Top             =   690
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MumLock"
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   1155
      TabIndex        =   1
      Top             =   690
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CapsLock"
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   690
      Width           =   1155
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   930
      Left            =   2565
      Shape           =   3  'Circle
      Top             =   -135
      Width           =   630
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   930
      Left            =   1395
      Shape           =   3  'Circle
      Top             =   -135
      Width           =   630
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   930
      Left            =   240
      Shape           =   3  'Circle
      Top             =   -135
      Width           =   630
   End
   Begin VB.Menu TrayMenu 
      Caption         =   "TrayMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuBringTop 
         Caption         =   "顯示程序主窗口(&S)"
      End
      Begin VB.Menu mnuSetting 
         Caption         =   "管理程序配置(&M)..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "關於Key Light(&A)..."
      End
      Begin VB.Menu b3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu b1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInit 
         Caption         =   "初始化程序设定(&I)"
      End
      Begin VB.Menu mnuCurrentSettings 
         Caption         =   "查看當前程序設置(&V)..."
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "退出(&E)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IsPopUpMenuShow As Boolean
Private Type SETTINGS
lpEnableColor As Long
lpDisableColor As Long
lpTopMost As Long
lpTrans As Long
lpTransValue As Long
End Type
Dim uFlags As SETTINGS
Public IsSet As Boolean
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
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
Private Sub cSysTray1_MouseDblClick(Button As Integer, Id As Long)
On Error Resume Next
IsPopUpMenuShow = False
BringWindowToTop Form1.hwnd
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
BringWindowToTop Form1.hwnd
End Sub
Private Sub cSysTray1_MouseDown(Button As Integer, Id As Long)
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
Dim rtn As Long
If Button = 2 Then
Timer2.Enabled = False
Timer1.Enabled = True
IsPopUpMenuShow = True
PopupMenu Me.TrayMenu
Timer1.Enabled = False
Timer2.Enabled = True
On Error Resume Next
BringWindowToTop Form1.hwnd
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
If 25 = 245 Then
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
End Select
End If
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
BringWindowToTop Form1.hwnd
Else
On Error Resume Next
BringWindowToTop Form1.hwnd
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
BringWindowToTop Form1.hwnd
End If
End Sub
Private Sub cSysTray1_MouseMove(Id As Long)
On Error Resume Next
Exit Sub
IsPopUpMenuShow = False
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_Activate()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
IsPopUpMenuShow = False
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_Click()
On Error Resume Next
IsPopUpMenuShow = False
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_DblClick()
On Error Resume Next
IsPopUpMenuShow = False
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_Deactivate()
On Error Resume Next
'IsPopUpMenuShow = False
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
IsPopUpMenuShow = False
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
IsPopUpMenuShow = False
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_GotFocus()
IsPopUpMenuShow = False
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_Initialize()
On Error Resume Next
IsPopUpMenuShow = False
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
If App.PrevInstance = True Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
With Form1
.Height = 1230
.Width = 3555
End With
MsgBox "本程序不允許同時運行2個及以上的實例" & vbCrLf & "請單擊'確定',終止應用程序", vbCritical, "Error"
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
Unload Me
On Error Resume Next
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
With Me.cSysTray1
.InTray = False
End With
Unload Me
Unload Form1
End
End
Else
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
IsPopUpMenuShow = False
If KeyCode = vbKeyF5 Then
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Load Form2
Form2.Hide
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
With Me.cSysTray1
.InTray = True
.TrayTip = "Key Light - PC-DOS Workshop"
End With
Exit Sub
End If
If KeyCode = vbKeyEscape Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
With Me.cSysTray1
.InTray = False
End With
Unload Me
Unload Form1
End
Exit Sub
End If
If KeyCode = vbKeyF4 And Shift = vbAltMask Then
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
With Me.cSysTray1
.InTray = False
End With
Unload Me
Unload Form1
End
Exit Sub
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
IsPopUpMenuShow = False
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
IsPopUpMenuShow = False
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_LinkClose()
On Error Resume Next
IsPopUpMenuShow = False
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_LinkError(LinkErr As Integer)
On Error Resume Next
IsPopUpMenuShow = False
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
On Error Resume Next
IsPopUpMenuShow = False
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_LinkOpen(Cancel As Integer)
On Error Resume Next
IsPopUpMenuShow = False
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_Load()
IsPopUpMenuShow = False
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Load Form2
Form2.Hide
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
With Me.cSysTray1
.InTray = True
.TrayTip = "Key Light - PC-DOS Workshop"
End With
End Sub
Private Sub Form_LostFocus()
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
ReleaseCapture
SendMessage Form1.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
Exit Sub
End If
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
Dim rtn As Long
If Button = 2 Then
On Error Resume Next
BringWindowToTop Form1.hwnd
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Timer2.Enabled = False
Timer1.Enabled = True
Timer2.Enabled = False
Timer1.Enabled = True
IsPopUpMenuShow = True
PopupMenu Me.TrayMenu
Timer1.Enabled = False
Timer2.Enabled = True
Timer1.Enabled = False
Timer2.Enabled = True
If 25 = 245 Then
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
End Select
End If
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
BringWindowToTop Form1.hwnd
Else
On Error Resume Next
BringWindowToTop Form1.hwnd
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
BringWindowToTop Form1.hwnd
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_Paint()
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_Resize()
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_Terminate()
On Error Resume Next
End
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
With Me.cSysTray1
.InTray = False
End With
Unload Me
Unload Form1
End
End Sub
Private Sub Label1_Change()
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label1_Click()
On Error Resume Next
IsPopUpMenuShow = False
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
End Select
With Form1
.Height = 1230
.Width = 3555
End With
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label1_DblClick()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label1_LinkClose()
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label1_LinkError(LinkErr As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label1_LinkNotify()
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label1_LinkOpen(Cancel As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
ReleaseCapture
SendMessage Form1.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
Exit Sub
End If
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
Dim rtn As Long
If Button = 2 Then
Timer2.Enabled = False
Timer1.Enabled = True
IsPopUpMenuShow = True
PopupMenu Me.TrayMenu
Timer1.Enabled = False
Timer2.Enabled = True
On Error Resume Next
BringWindowToTop Form1.hwnd
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
If 25 = 245 Then
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
End Select
End If
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
BringWindowToTop Form1.hwnd
Else
On Error Resume Next
BringWindowToTop Form1.hwnd
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
BringWindowToTop Form1.hwnd
End If
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label1_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label1_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label2_Change()
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label2_Click()
On Error Resume Next
IsPopUpMenuShow = False
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label2_DblClick()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label2_LinkClose()
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label2_LinkError(LinkErr As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label2_LinkNotify()
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label2_LinkOpen(Cancel As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
ReleaseCapture
SendMessage Form1.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
Exit Sub
End If
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
Dim rtn As Long
If Button = 2 Then
Timer2.Enabled = False
Timer1.Enabled = True
IsPopUpMenuShow = True
PopupMenu Me.TrayMenu
Timer1.Enabled = False
Timer2.Enabled = True
On Error Resume Next
BringWindowToTop Form1.hwnd
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
If 25 = 245 Then
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
End Select
End If
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
BringWindowToTop Form1.hwnd
Else
On Error Resume Next
BringWindowToTop Form1.hwnd
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
BringWindowToTop Form1.hwnd
End If
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label2_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label2_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label2_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label2_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label2_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label3_Change()
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label3_Click()
On Error Resume Next
IsPopUpMenuShow = False
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
BringWindowToTop Form1.hwnd
End Sub
Private Sub Label3_DblClick()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label3_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label3_LinkClose()
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label3_LinkError(LinkErr As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label3_LinkNotify()
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label3_LinkOpen(Cancel As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
ReleaseCapture
SendMessage Form1.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
Exit Sub
End If
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
Dim rtn As Long
If Button = 2 Then
Timer2.Enabled = False
Timer1.Enabled = True
IsPopUpMenuShow = True
PopupMenu Me.TrayMenu
Timer1.Enabled = False
Timer2.Enabled = True
On Error Resume Next
BringWindowToTop Form1.hwnd
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
If 25 = 245 Then
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
End Select
End If
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
BringWindowToTop Form1.hwnd
Else
On Error Resume Next
BringWindowToTop Form1.hwnd
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
BringWindowToTop Form1.hwnd
End If
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub mnuAbout_Click()
On Error Resume Next
IsPopUpMenuShow = False
frmAbout.Show 1
End Sub
Private Sub mnuBringTop_Click()
On Error Resume Next
IsPopUpMenuShow = False
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
BringWindowToTop Form1.hwnd
End Sub
Private Sub mnuCurrentSettings_Click()
On Error Resume Next
IsPopUpMenuShow = False
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
With Me
.Height = 1230
.Width = 3555
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
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
BringWindowToTop hwnd
End Sub
Private Sub mnuEnd_Click()
On Error Resume Next
IsPopUpMenuShow = False
SaveSetting "KeyLight", "GUISettings", "TopMost", CStr(Form2.Check1.Value)
SaveSetting "KeyLight", "GUISettings", "Trans", CStr(Form2.Check2.Value)
SaveSetting "KeyLight", "GUISettings", "TransValue", CStr(Form2.HScroll1.Value)
SaveSetting "KeyLight", "GUISettings", "KeyEnableColor", CStr(Form2.Command1.BackColor)
SaveSetting "KeyLight", "GUISettings", "KeyDisableColor", CStr(Form2.Command2.BackColor)
With Me.cSysTray1
.InTray = False
End With
Unload Me
Unload Form1
End
End Sub
Private Sub mnuInit_Click()
On Error Resume Next
On Error Resume Next
IsPopUpMenuShow = False
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Load Form2
Form2.Hide
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
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
Unload Form2
Load Form2
Form2.Hide
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
With Me
.Left = .Left
.Top = .Top
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
With Me.cSysTray1
.InTray = True
.TrayTip = "Key Light - PC-DOS Workshop"
End With
Else
Load Form2
Form2.Hide
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
With Me
.Left = .Left
.Top = .Top
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
With Me.cSysTray1
.InTray = True
.TrayTip = "Key Light - PC-DOS Workshop"
End With
Exit Sub
End If
End Sub
Private Sub mnuRefresh_Click()
On Error Resume Next
IsPopUpMenuShow = False
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Load Form2
Form2.Hide
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
With Me
.Left = .Left
.Top = .Top
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
With Me.cSysTray1
.InTray = True
.TrayTip = "Key Light - PC-DOS Workshop"
End With
End Sub
Private Sub mnuSetting_Click()
On Error Resume Next
IsPopUpMenuShow = False
Form2.Show 1
IsSet = True
Me.Tag = "Suicune"
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
If IsPopUpMenuShow = True Then
On Error Resume Next
With lpKeyState
.dwCapsLock = GetKeyState(vbKeyCapital)
.dwNumLock = GetKeyState(vbKeyNumlock)
.dwScrollLock = GetKeyState(vbKeyScrollLock)
End With
Select Case lpKeyState.dwCapsLock
Case 0
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
Exit Sub
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
Dim rtn As Long
With lpKeyState
.dwCapsLock = GetKeyState(vbKeyCapital)
.dwNumLock = GetKeyState(vbKeyNumlock)
.dwScrollLock = GetKeyState(vbKeyScrollLock)
End With
Select Case lpKeyState.dwCapsLock
Case 0
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
If Tag = "Suicune" Then
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
End If
If Tag <> "" Then
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
End If
If Tag = "Set" Then
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
End If
If Me.Tag = "" Then
On Error Resume Next
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End If
If FindWindow(vbNullString, Form2.Caption) = 0 Then
Exit Sub
End If
If IsSet = True Then
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
End If
If Tag = "Suicune" Then
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
Exit Sub
End If
Exit Sub
On Error Resume Next
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub label3_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub label3_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub label3_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub label3_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub label3_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub label3_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End Sub
Private Sub Timer2_Timer()
On Error Resume Next
On Error Resume Next
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
If IsPopUpMenuShow = True Then
On Error Resume Next
With lpKeyState
.dwCapsLock = GetKeyState(vbKeyCapital)
.dwNumLock = GetKeyState(vbKeyNumlock)
.dwScrollLock = GetKeyState(vbKeyScrollLock)
End With
Select Case lpKeyState.dwCapsLock
Case 0
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
Exit Sub
End If
Dim rtn As Long
If Tag = "" Then
SetWindowPos Form1.hwnd, HWND_TOP, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 1230
.Width = 3555
End With
Select Case Form2.Check1.Value
Case 1
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Form1
.Height = 1230
.Width = 3555
End With
With Form1
.Height = 1230
.Width = 3555
End With
Case 0
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
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
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape1
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwNumLock
Case 0
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape2
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
Select Case lpKeyState.dwScrollLock
Case 0
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command2.BackColor
.Visible = True
End With
Case 1
With Me.Shape3
.BorderStyle = 1
.BorderColor = RGB(127, 127, 127)
.BackStyle = 1
.BackColor = Form2.Command1.BackColor
.Visible = True
End With
End Select
With Me.Label1
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "CapsLock" & vbCrLf & "大寫鎖定"
End With
With Me.Label2
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "NumLock" & vbCrLf & "數碼鎖定"
End With
With Me.Label3
.Alignment = 2
.BackStyle = 0
.BorderStyle = 0
.BackColor = RGB(0, 0, 0)
.Caption = "ScrollLock" & vbCrLf & "滾動鎖定"
End With
If 1 = 245 Then
With Me
.Left = Screen.Width - .Width - 5
.Top = Screen.Height - .Height - 245 * 3
.BackColor = RGB(0, 0, 0)
.KeyPreview = True
.Enabled = True
End With
End If
With Timer1
.Interval = 245 * 2
.Enabled = True
End With
End If

End Sub
