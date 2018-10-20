VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmUpdate 
   Caption         =   "Updating...You Must Be Connected To The Internet"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8310
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wbUpdate 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      ExtentX         =   13573
      ExtentY         =   11456
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
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
Private Sub Form_Load()
Form_Resize
Dim MSG As Long
MSG = MsgBox("Please Connect To The Internet Before Continuing." + vbCrLf + vbCrLf + "NOTE: No Personal Information Is Sent, Only The Version Number.", vbInformation + vbApplicationModal + vbOKCancel, "TextMe Web Update")
If MSG = vbCancel Then
Unload Me
Exit Sub
End If
wbUpdate.Navigate "http://www15.brinkster.com/pmchapman/update.asp?v=" + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
End Sub

Private Sub Form_Resize()
wbUpdate.Height = ScaleHeight
wbUpdate.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Trim$(UCase$(Command$)) = ":UPDATE" Then End
End Sub
