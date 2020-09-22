VERSION 5.00
Begin VB.Form frmNTFS 
   Caption         =   "NTFS Permission"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   $"frmNTFS.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "Username to give permission to:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Folder to give permission on:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "frmNTFS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim sUserName As String
    Dim sFolderName As String
    sUserName = Trim$(CStr(Text2.Text))
    sFolderName = Trim$(CStr(Text1.Text))
    SetAccess sUserName, sFolderName, GENERIC_READ Or GENERIC_EXECUTE Or DELETE Or GENERIC_WRITE
End Sub

Private Sub Command2_Click()
    Dim sUserName As String
    Dim sFolderName As String
    sUserName = Trim$(Text2.Text)
    sFolderName = Trim$(Text1.Text)
    SetAccess sUserName, sFolderName, GENERIC_EXECUTE Or GENERIC_READ
End Sub

Private Sub Form_Load()
    Text1.Text = "enter folder name"
    Text2.Text = "enter username"
    Command1.Caption = "Change"
    Command2.Caption = "Read && Add"
End Sub
