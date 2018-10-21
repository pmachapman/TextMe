VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmUpdate 
   Caption         =   "Updating...You Must Be Connected To The Internet"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6225
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wbUpdate 
      CausesValidation=   0   'False
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6135
      ExtentX         =   10821
      ExtentY         =   5530
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim Url As String
    If GetWindowsVersion < 6 Then
        Url = "http://textme.conglomo.co.nz/update.php?v="
    Else
        Url = "https://textme.conglomo.co.nz/update.php?v="
    End If
    wbUpdate.Navigate2 Url + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
    Form_Resize
    Dim Msg As Long
    Msg = MsgBox("Please Connect To The Internet Before Continuing." + vbCrLf + vbCrLf + "NOTE: No Personal Information Is Sent, Only The Version Number.", vbInformation + vbApplicationModal + vbOKCancel, "TextMe Web Update")
    If Msg = vbCancel Then
        Unload Me
        Exit Sub
    End If
    wbUpdate.Navigate2 Url + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
End Sub

Private Sub Form_Resize()
    wbUpdate.Height = ScaleHeight
    wbUpdate.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Trim$(UCase$(Command$)) = ":UPDATE" Then End
End Sub

