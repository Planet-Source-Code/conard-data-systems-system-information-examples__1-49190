VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Computer Info Examples"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "System Info"
      Height          =   2415
      Left            =   3840
      TabIndex        =   9
      Top             =   2760
      Width           =   3735
      Begin VB.CommandButton Command6 
         Caption         =   "System Metrics"
         Height          =   375
         Left            =   840
         TabIndex        =   18
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1920
         TabIndex        =   12
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1920
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "High mem address:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Low mem address:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Processor type:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Number of processors:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "System Commands"
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   3615
      Begin VB.Frame Frame5 
         Caption         =   "For windows NT"
         Height          =   1575
         Left            =   0
         TabIndex        =   21
         Top             =   840
         Width           =   3615
         Begin VB.CommandButton Command11 
            Caption         =   "Log off"
            Height          =   375
            Left            =   840
            TabIndex        =   24
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Reboot"
            Height          =   375
            Left            =   840
            TabIndex        =   23
            Top             =   720
            Width           =   1815
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Shutdown"
            Height          =   375
            Left            =   840
            TabIndex        =   22
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Shutdown System"
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Other Stuff"
      Height          =   2535
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton Command12 
         Caption         =   "NTFS permissions"
         Height          =   375
         Left            =   1920
         TabIndex        =   26
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Run a program"
         Height          =   375
         Left            =   1920
         TabIndex        =   25
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Windows version"
         Height          =   375
         Left            =   1920
         TabIndex        =   20
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Lockout example"
         Height          =   495
         Left            =   1920
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Get sys directory"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Text effect"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Get network info"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Get/set computer name"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "All User Directories"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Stuff for user directories
Private Const TOKEN_QUERY = (&H8)
Private Declare Function GetAllUsersProfileDirectory Lib "userenv.dll" Alias "GetAllUsersProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetDefaultUserProfileDirectory Lib "userenv.dll" Alias "GetDefaultUserProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetProfilesDirectory Lib "userenv.dll" Alias "GetProfilesDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetUserProfileDirectory Lib "userenv.dll" Alias "GetUserProfileDirectoryA" (ByVal hToken As Long, ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long

'Stuff for getting and setting the computer name
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long

'Stuff for the shutdown process
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Const SE_PRIVILEGE_ENABLED = &H2
Const TokenPrivileges = 3
Const TOKEN_ASSIGN_PRIMARY = &H1
Const TOKEN_DUPLICATE = &H2
Const TOKEN_IMPERSONATE = &H4
Const TOKEN_QUERY_SOURCE = &H10
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_ADJUST_GROUPS = &H40
Const TOKEN_ADJUST_DEFAULT = &H80
Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Const ANYSIZE_ARRAY = 1
Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type
Private Type LUID
    LowPart As Long
    HighPart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    'pLuid As Luid
    pLuid As LARGE_INTEGER
    Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type
Private Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long

'This is for getting the system directory
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'This is used for getting system information
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

'Stuff for windows version
Private Declare Function GetVersion Lib "kernel32" () As Long

'This stuff is for the run dialog
Const shrdNoMRUString = &H2    '2nd bit is set
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long

Public Function GetWinVersion() As String
    'More stuff for windows version
    Dim Ver As Long, WinVer As Long
    Ver = GetVersion()
    WinVer = Ver And &HFFFF&
    'retrieve the windows version
    GetWinVersion = Format((WinVer Mod 256) + ((WinVer \ 256) / 100), "Fixed")
End Function

Private Sub Command1_Click()
    'Stuff to get the computer name
    Dim dwLen As Long
    Dim strString As String
    Dim YesNo As String
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    GetComputerName strString, dwLen
    strString = Left(strString, dwLen)
    MsgBox "Your computer name is: " & strString
    
    YesNo = MsgBox("Would you like to change your computer name?", vbYesNo, "Set")
    If YesNo = vbYes Then
        Dim sNewName As String
        sNewName = InputBox("Please enter a new computer name.")
        SetComputerName sNewName
        MsgBox "Computer name set to: " + sNewName
        MsgBox "Changes will only take effect after you restart your computer."
    Else
        Exit Sub
    End If
End Sub

Private Sub Command10_Click()
    'Reboot
    RebootNT True
End Sub

Private Sub Command11_Click()
    'Log off
    LogOffNT True
End Sub

Private Sub Command12_Click()
    frmNTFS.Show
End Sub

Private Sub Command13_Click()
    Dim sTitle As String, sPrompt As String
    sTitle = "Start a program ..."
    sPrompt = "Type the name of a program ..."
    If IsWinNT Then
        SHRunDialog Me.hwnd, 0, 0, StrConv(sTitle, vbUnicode), StrConv(sPrompt, vbUnicode), 0
    Else
        SHRunDialog Me.hwnd, 0, 0, sTitle, sPrompt, 0
    End If
End Sub

Private Sub Command2_Click()
    'Stuff for network info
    Dim YesNo As String
    YesNo = MsgBox("Notice: While this process is taking place your computer may appear to freeze as it collects data. Freeze time may be up to 1 minute. Do you wish to continue?", vbYesNo, "Warning")
    If YesNo = vbYes Then
        MsgBox "Again, your computer will appear to freeze. Do not forcefully close any application, wait till it's done."
        GetNetInfo
    Else
        Exit Sub
    End If
End Sub

Private Sub Command3_Click()
    'Stuff for system shutdown
    Dim YesNo As String
    YesNo = MsgBox("Are you sure you want to shutdown your computer?", vbYesNo, "Shutdown")
    If YesNo = vbYes Then
        InitiateShutdownMachine GetMyMachineName, True, True, True, 60, "You initiated a system shutdown..."
    Else
        Exit Sub
    End If
End Sub

Private Sub Command4_Click()
    frmTE.Show
End Sub

Private Sub Command5_Click()
    'Stuff for getting the system directory
    Dim sSave As String, ret As Long
    sSave = Space(255)
    ret = GetSystemDirectory(sSave, 255)
    sSave = Left$(sSave, ret)
    MsgBox "Windows System directory: " + sSave
End Sub

Private Sub Command6_Click()
    frmMetrics.Show
End Sub

Private Sub Command7_Click()
frmMenu.Show
End Sub

Private Sub Command8_Click()
    'Get the windows version
    MsgBox "Windows version: " + GetWinVersion
End Sub

Private Sub Command9_Click()
    'Shutdown
    ShutDownNT True
End Sub

Private Sub Form_Load()
    'This is for the all user directories
    Dim sBuffer As String, ret As Long, hToken As Long
    
    sBuffer = String(255, 0)
    GetAllUsersProfileDirectory sBuffer, 255
    List1.AddItem (StripTerminator(sBuffer))

    sBuffer = String(255, 0)
    GetDefaultUserProfileDirectory sBuffer, 255
    List1.AddItem (StripTerminator(sBuffer))

    sBuffer = String(255, 0)
    GetProfilesDirectory sBuffer, 255
    List1.AddItem (StripTerminator(sBuffer))
    
    sBuffer = String(255, 0)
    OpenProcessToken GetCurrentProcess, TOKEN_QUERY, hToken
    GetUserProfileDirectory hToken, sBuffer, 255
    List1.AddItem (StripTerminator(sBuffer))
    
    'This is for system information
    Dim SInfo As SYSTEM_INFO
    GetSystemInfo SInfo
    Text1.Text = Str$(SInfo.dwNumberOrfProcessors)
    Text2.Text = Str$(SInfo.dwProcessorType)
    Text3.Text = Str$(SInfo.lpMinimumApplicationAddress)
    Text4.Text = Str$(SInfo.lpMaximumApplicationAddress)
End Sub

Function StripTerminator(sInput As String) As String
    'This is for the all user directories
    Dim ZeroPos As Long
    ZeroPos = InStr(1, sInput, Chr$(0))
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function

Public Function InitiateShutdownMachine(ByVal Machine As String, Optional Force As Variant, Optional Restart As Variant, Optional AllowLocalShutdown As Variant, Optional Delay As Variant, Optional Message As Variant) As Boolean
    'Stuff for system shutdown
    Dim hProc As Long
    Dim OldTokenStuff As TOKEN_PRIVILEGES
    Dim OldTokenStuffLen As Long
    Dim NewTokenStuff As TOKEN_PRIVILEGES
    Dim NewTokenStuffLen As Long
    Dim pSize As Long
    If IsMissing(Force) Then Force = False
    If IsMissing(Restart) Then Restart = True
    If IsMissing(AllowLocalShutdown) Then AllowLocalShutdown = False
    If IsMissing(Delay) Then Delay = 0
    If IsMissing(Message) Then Message = ""
    'Make sure the Machine-name doesn't start with '\\'
    If InStr(Machine, "\\") = 1 Then
        Machine = Right(Machine, Len(Machine) - 2)
    End If
    'check if it's the local machine that's going to be shutdown
    If (LCase(GetMyMachineName) = LCase(Machine)) Then
        'may we shut this computer down?
        If AllowLocalShutdown = False Then Exit Function
        'open access token
        If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hProc) = 0 Then
            MsgBox "OpenProcessToken Error: " & GetLastError()
            Exit Function
        End If
        'retrieve the locally unique identifier to represent the Shutdown-privilege name
        If LookupPrivilegeValue(vbNullString, SE_SHUTDOWN_NAME, OldTokenStuff.Privileges(0).pLuid) = 0 Then
            MsgBox "LookupPrivilegeValue Error: " & GetLastError()
            Exit Function
        End If
        NewTokenStuff = OldTokenStuff
        NewTokenStuff.PrivilegeCount = 1
        NewTokenStuff.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
        NewTokenStuffLen = Len(NewTokenStuff)
        pSize = Len(NewTokenStuff)
        'Enable shutdown-privilege
        If AdjustTokenPrivileges(hProc, False, NewTokenStuff, NewTokenStuffLen, OldTokenStuff, OldTokenStuffLen) = 0 Then
            MsgBox "AdjustTokenPrivileges Error: " & GetLastError()
            Exit Function
        End If
        'initiate the system shutdown
        If InitiateSystemShutdown("\\" & Machine, Message, Delay, Force, Restart) = 0 Then
            Exit Function
        End If
        NewTokenStuff.Privileges(0).Attributes = 0
        'Disable shutdown-privilege
        If AdjustTokenPrivileges(hProc, False, NewTokenStuff, Len(NewTokenStuff), OldTokenStuff, Len(OldTokenStuff)) = 0 Then
            Exit Function
        End If
    Else
        'initiate the system shutdown
        If InitiateSystemShutdown("\\" & Machine, Message, Delay, Force, Restart) = 0 Then
            Exit Function
        End If
    End If
    InitiateShutdownMachine = True
End Function

Function GetMyMachineName() As String
    'Stuff for the shutdown, and can be used for other things
    Dim sLen As Long
    GetMyMachineName = Space(100)
    sLen = 100
    If GetComputerName(GetMyMachineName, sLen) Then
        GetMyMachineName = Left(GetMyMachineName, sLen)
    End If
End Function

Function IsWinNT() As Boolean
    'This is for the run function, but can be used for other things
    Dim OSInfo As OSVERSIONINFO
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    ret& = GetVersionEx(OSInfo)
    IsWinNT = (OSInfo.dwPlatformId = 2)
End Function
