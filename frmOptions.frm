VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TextMe Settings"
   ClientHeight    =   1455
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   3750
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRead 
      Caption         =   "Read &Only"
      Height          =   375
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   10
      ToolTipText     =   "Select This To Make The Text In TextMe Read Only"
      Top             =   480
      Width           =   1935
   End
   Begin VB.CheckBox chkDelete 
      Caption         =   "Delete Original &File"
      Height          =   375
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   9
      ToolTipText     =   $"frmOptions.frx":000C
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      ToolTipText     =   "Apply Settings"
      Top             =   975
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Close This Dialog Without Applying New Settings"
      Top             =   975
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   90
      TabIndex        =   0
      ToolTipText     =   "Close This Dialog And Apply New Settings"
      Top             =   975
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    SaveSetting "Peter Chapman", "TextMe", "DeleteOriginal", chkDelete.Value
    frmMain.txtMain.Locked = chkRead.Value
    SaveSetting "Peter Chapman", "TextMe", "ReadOnly", chkRead.Value
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    cmdApply_Click
    Unload Me
End Sub

Private Sub Form_Load()
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    'Get Settings
    chkDelete.Value = GetSetting("Peter Chapman", "TextMe", "DeleteOriginal", "0")
    chkRead.Value = GetSetting("Peter Chapman", "TextMe", "ReadOnly", "0")
End Sub
