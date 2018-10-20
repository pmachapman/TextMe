VERSION 5.00
Begin VB.Form frmBackup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup TextMe Database"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   Icon            =   "frmBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox drvPath 
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   3375
   End
   Begin VB.DirListBox folPath 
      Height          =   2565
      Left            =   1200
      TabIndex        =   7
      Top             =   480
      Width           =   3375
   End
   Begin VB.CheckBox chkDelete 
      Caption         =   "Delete All Existing Notes"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Delete All Notes In The Backup Folder Before Backing Up"
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      ToolTipText     =   "Cancel Backup of Note Database"
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "Backup Note Database"
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      Caption         =   "WARNING: Notes In The Backup Folder Will Be Replaced"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3960
      Width           =   4695
   End
   Begin VB.Label lblDescription 
      Caption         =   "Backup To"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCopy_Click()
Dim sFiles() As String
Dim lCtr As Long
On Error GoTo CopyErr
If frmBackup.Caption = "Restore TextMe Database" Then
    If Right$(txtPath.Text, 1) = "/" Then txtPath.Text = Left$(txtPath.Text, Len(txtPath.Text) - 1) + "\"
    If Right$(txtPath.Text, 1) <> "\" Then txtPath.Text = txtPath.Text + "\"
    sFiles = AllFiles(txtPath.Text)
    If sFiles(0) = "" Then
    MsgBox "Restore Directory Empty Or Does Not Exist, Please Select Another Folder To Restore From.", vbApplicationModal + vbExclamation + vbOKOnly, "TextMe Notes Database Restore"
    Exit Sub
    End If
    If chkDelete.Value = 1 Then
        sFiles = AllFiles(App.Path + "\TextFiles\")
        If sFiles(0) <> "" Then Kill App.Path + "\TextFiles\" & "*.*"
    End If
    sFiles = AllFiles(txtPath.Text)
    
    For lCtr = 0 To UBound(sFiles)
        FileCopy txtPath.Text & sFiles(lCtr), App.Path + "\textfiles\" & sFiles(lCtr)
    Next
    frmMain.filMain.Refresh
    frmMain.filRefresh.Refresh
    frmMain.filMain.Selected(0) = True
Else
    If Right$(txtPath.Text, 1) = "/" Then txtPath.Text = Left$(txtPath.Text, Len(txtPath.Text) - 1) + "\"
    If Right$(txtPath.Text, 1) <> "\" Then txtPath.Text = txtPath.Text + "\"
    If chkDelete.Value = 1 Then
        sFiles = AllFiles(txtPath.Text)
        If sFiles(0) <> "" Then Kill txtPath.Text & "*.*"
    End If
    sFiles = AllFiles(App.Path + "\textfiles\")
    For lCtr = 0 To UBound(sFiles)
        FileCopy App.Path + "\textfiles\" & sFiles(lCtr), txtPath.Text & sFiles(lCtr)
    Next
End If
Unload Me
Exit Sub
CopyErr:
If Me.Caption = "Restore TextMe Database" Then
    MsgBox "The TextMe Database Could Not Be Restored. Please Check The Restore Folder Exists And Contains Valid Notes.", vbApplicationModal + vbExclamation + vbOKOnly, "TextMe Notes Database Restore"
    If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmBackup.cmdCopy_Click", Err.Number, Err.Description + " while restoring"
Else
    MsgBox "The TextMe Database Could Not Be Backed Up. Please Check The Backup Folder Exists And Is Empty.", vbApplicationModal + vbExclamation + vbOKOnly, "TextMe Notes Database Backup"
    If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmBackup.cmdCopy_Click", Err.Number, Err.Description + " while backing up"
End If
End Sub

Private Sub drvPath_Change()
On Error GoTo NoDrive
ChDrive drvPath.Drive
folPath.Path = CurDir
drvPath.Refresh
folPath.Refresh
Exit Sub
NoDrive:
MsgBox "Drive Inaccessible, Please Select Another", vbApplicationModal + vbExclamation + vbOKOnly, "Cannot Select Drive"
drvPath.SetFocus
End Sub

Private Sub folPath_Change()
drvPath.Drive = folPath.Path
If Right$(folPath.Path, 1) = "\" Then txtPath.Text = folPath.Path Else txtPath.Text = folPath.Path + "\"
txtPath.SelStart = Len(txtPath.Text)
End Sub

Private Sub Form_Activate()
frmMain.timMain.Interval = 0
If frmMain.recBorder(6).FillStyle = 1 Then
    Me.Caption = "Restore TextMe Database"
    lblDescription.Caption = "Restore From"
    lblWarning.Caption = "WARNING: Notes In The TextMe Database Will Be Replaced"
    cmdCopy.ToolTipText = "Restore Note Database"
    cmdCancel.ToolTipText = "Cancel Restoration of Note Database"
    chkDelete.ToolTipText = "Clear Database Before Restoration"
End If
drvPath.Drive = App.Path
folPath.Path = App.Path
txtPath.Text = folPath.Path + "\"
txtPath.SelStart = Len(txtPath.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.timMain.Interval = 5000
End Sub
