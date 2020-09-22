VERSION 5.00
Begin VB.Form frmMetrics 
   Caption         =   "System Metrics"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3360
      TabIndex        =   13
      Text            =   "Text7"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3360
      TabIndex        =   12
      Text            =   "Text6"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3360
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3360
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3360
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3360
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Is your machine too slow to run windows?:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "Maximum width when resizing a window:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Width between desktop icons:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Height of windows caption:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Screen Y:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Screen X:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Number of mouse buttons:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmMetrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Const SM_CXSCREEN = 0 'X Size of screen
Const SM_CYSCREEN = 1 'Y Size of Screen
Const SM_CXVSCROLL = 2 'X Size of arrow in vertical scroll bar.
Const SM_CYHSCROLL = 3 'Y Size of arrow in horizontal scroll bar
Const SM_CYCAPTION = 4 'Height of windows caption
Const SM_CXBORDER = 5 'Width of no-sizable borders
Const SM_CYBORDER = 6 'Height of non-sizable borders
Const SM_CXDLGFRAME = 7 'Width of dialog box borders
Const SM_CYDLGFRAME = 8 'Height of dialog box borders
Const SM_CYVTHUMB = 9 'Height of scroll box on horizontal scroll bar
Const SM_CXHTHUMB = 10 ' Width of scroll box on horizontal scroll bar
Const SM_CXICON = 11 'Width of standard icon
Const SM_CYICON = 12 'Height of standard icon
Const SM_CXCURSOR = 13 'Width of standard cursor
Const SM_CYCURSOR = 14 'Height of standard cursor
Const SM_CYMENU = 15 'Height of menu
Const SM_CXFULLSCREEN = 16 'Width of client area of maximized window
Const SM_CYFULLSCREEN = 17 'Height of client area of maximized window
Const SM_CYKANJIWINDOW = 18 'Height of Kanji window
Const SM_MOUSEPRESENT = 19 'True is a mouse is present
Const SM_CYVSCROLL = 20 'Height of arrow in vertical scroll bar
Const SM_CXHSCROLL = 21 'Width of arrow in vertical scroll bar
Const SM_DEBUG = 22 'True if deugging version of windows is running
Const SM_SWAPBUTTON = 23 'True if left and right buttons are swapped.
Const SM_CXMIN = 28 'Minimum width of window
Const SM_CYMIN = 29 'Minimum height of window
Const SM_CXSIZE = 30 'Width of title bar bitmaps
Const SM_CYSIZE = 31 'height of title bar bitmaps
Const SM_CXMINTRACK = 34 'Minimum tracking width of window
Const SM_CYMINTRACK = 35 'Minimum tracking height of window
Const SM_CXDOUBLECLK = 36 'double click width
Const SM_CYDOUBLECLK = 37 'double click height
Const SM_CXICONSPACING = 38 'width between desktop icons
Const SM_CYICONSPACING = 39 'height between desktop icons
Const SM_MENUDROPALIGNMENT = 40 'Zero if popup menus are aligned to the left of the memu bar item. True if it is aligned to the right.
Const SM_PENWINDOWS = 41 'The handle of the pen windows DLL if loaded.
Const SM_DBCSENABLED = 42 'True if double byte characteds are enabled
Const SM_CMOUSEBUTTONS = 43 'Number of mouse buttons.
Const SM_CMETRICS = 44 'Number of system metrics
Const SM_CLEANBOOT = 67 'Windows 95 boot mode. 0 = normal, 1 = safe, 2 = safe with network
Const SM_CXMAXIMIZED = 61 'default width of win95 maximised window
Const SM_CXMAXTRACK = 59 'maximum width when resizing win95 windows
Const SM_CXMENUCHECK = 71 'width of menu checkmark bitmap
Const SM_CXMENUSIZE = 54 'width of button on menu bar
Const SM_CXMINIMIZED = 57 'width of rectangle into which minimised windows must fit.
Const SM_CYMAXIMIZED = 62 'default height of win95 maximised window
Const SM_CYMAXTRACK = 60 'maximum width when resizing win95 windows
Const SM_CYMENUCHECK = 72 'height of menu checkmark bitmap
Const SM_CYMENUSIZE = 55 'height of button on menu bar
Const SM_CYMINIMIZED = 58 'height of rectangle into which minimised windows must fit.
Const SM_CYSMCAPTION = 51 'height of windows 95 small caption
Const SM_MIDEASTENABLED = 74 'Hebrw and Arabic enabled for windows 95
Const SM_NETWORK = 63 'bit o is set if a network is present. Const SM_SECURE = 44 'True if security is present on windows 95 system
Const SM_SLOWMACHINE = 73 'true if machine is too slow to run win95.

Private Sub Form_Load()
    Text1.Text = Str$(GetSystemMetrics(SM_CMOUSEBUTTONS))
    Text2.Text = Str$(GetSystemMetrics(SM_CXSCREEN))
    Text3.Text = Str$(GetSystemMetrics(SM_CYSCREEN))
    Text4.Text = Str$(GetSystemMetrics(SM_CYCAPTION))
    Text5.Text = Str$(GetSystemMetrics(SM_CXICONSPACING))
    Text6.Text = Str$(GetSystemMetrics(SM_CYMAXTRACK))
    Text7.Text = Str$(GetSystemMetrics(SM_SLOWMACHINE)) & " (0 = No, 1 = Yes)"
End Sub

