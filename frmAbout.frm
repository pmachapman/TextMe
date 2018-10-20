VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3495
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2412.311
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2505
      Width           =   1260
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Web Update"
      Height          =   345
      Left            =   4260
      TabIndex        =   1
      Top             =   2955
      Width           =   1245
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":000C
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblBugs 
      Caption         =   "Send All Bug Reports To DukeSpukem@Hotmail.com"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   3855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1521.93
      Y2              =   1521.93
   End
   Begin VB.Label lblDescription 
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1532.283
      Y2              =   1532.283
   End
   Begin VB.Label lblVersion 
      Height          =   225
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   255
      TabIndex        =   3
      Top             =   2385
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next
Unload frmUpdate
frmUpdate.Show vbModal, Me
End Sub

Private Sub Form_Load()
Me.Caption = "About " + App.Title
lblTitle.Caption = App.Title
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
lblDisclaimer.Caption = App.LegalCopyright
lblDescription.Caption = App.Comments
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Trim$(UCase$(Command$)) = ":ABOUT" Then End
End Sub
