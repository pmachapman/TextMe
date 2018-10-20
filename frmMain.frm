VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00DEEDEF&
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7365
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   277
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   491
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox filRefresh 
      Height          =   2625
      Left            =   2280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Timer timMain 
      Interval        =   5000
      Left            =   2400
      Top             =   3600
   End
   Begin VB.TextBox txtMain 
      DataSource      =   "dtaMain"
      Height          =   3795
      Left            =   2160
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   5160
   End
   Begin VB.PictureBox Toolbar 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEEDEF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   375
      ScaleWidth      =   7335
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   7335
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   8
         Left            =   3000
         OLEDropMode     =   1  'Manual
         ToolTipText     =   "TextMe Settings"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   9
         Left            =   3360
         OLEDropMode     =   1  'Manual
         ToolTipText     =   "Check For Updates To TextMe (Requires Internet Connection)"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   11
         Left            =   4080
         OLEDropMode     =   1  'Manual
         ToolTipText     =   "Exit The TextMe Notes Organiser"
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   10
         Left            =   3720
         OLEDropMode     =   1  'Manual
         ToolTipText     =   "About The TextMe Notes Organiser"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgDivider 
         Height          =   360
         Index           =   3
         Left            =   2900
         OLEDropMode     =   1  'Manual
         Top             =   0
         Width           =   45
      End
      Begin VB.Image imgDivider 
         Height          =   360
         Index           =   2
         Left            =   2100
         OLEDropMode     =   1  'Manual
         Top             =   0
         Width           =   45
      End
      Begin VB.Image imgDivider 
         Height          =   360
         Index           =   1
         Left            =   500
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":0442
         Top             =   0
         Width           =   45
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   7
         Left            =   2520
         OLEDropMode     =   1  'Manual
         ToolTipText     =   "Restore Your Notes"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   6
         Left            =   2160
         OLEDropMode     =   1  'Manual
         ToolTipText     =   "Backup Your Notes"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   5
         Left            =   1680
         OLEDropMode     =   1  'Manual
         ToolTipText     =   "Delete The Open Note (Hold Down Shift To Delete All Notes)"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   4
         Left            =   1320
         OLEDropMode     =   1  'Manual
         ToolTipText     =   "Rename The Open Note"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   3
         Left            =   960
         OLEDropMode     =   1  'Manual
         ToolTipText     =   "Copy The Open Note (Including Unsaved Changes) To A New Note"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgDivider 
         Height          =   360
         Index           =   0
         Left            =   30
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":05A4
         Top             =   0
         Width           =   75
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   2
         Left            =   600
         OLEDropMode     =   1  'Manual
         ToolTipText     =   "Create A New, Blank Note"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   1
         Left            =   120
         OLEDropMode     =   1  'Manual
         ToolTipText     =   "Save Changes To The Open Note"
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.FileListBox filMain 
      Height          =   3795
      Left            =   0
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Choose A Note To Open. You Can Also Drag A Text File In Here To Create A Note From The File"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Image imgButtonN 
      Height          =   360
      Index           =   8
      Left            =   3000
      Picture         =   "frmMain.frx":0766
      Top             =   5280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonD 
      Height          =   360
      Index           =   8
      Left            =   3000
      Picture         =   "frmMain.frx":0E68
      Top             =   4800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonO 
      Height          =   360
      Index           =   8
      Left            =   3000
      Picture         =   "frmMain.frx":156A
      Top             =   4320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonN 
      Height          =   360
      Index           =   9
      Left            =   3360
      Picture         =   "frmMain.frx":1C6C
      Top             =   5280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonN 
      Height          =   360
      Index           =   1
      Left            =   120
      Picture         =   "frmMain.frx":236E
      Top             =   5280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonN 
      Height          =   360
      Index           =   2
      Left            =   600
      Picture         =   "frmMain.frx":2A70
      Top             =   5280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonN 
      Height          =   360
      Index           =   3
      Left            =   960
      Picture         =   "frmMain.frx":3172
      Top             =   5280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonN 
      Height          =   360
      Index           =   4
      Left            =   1320
      Picture         =   "frmMain.frx":3874
      Top             =   5280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonN 
      Height          =   360
      Index           =   5
      Left            =   1680
      Picture         =   "frmMain.frx":3F76
      Top             =   5280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonN 
      Height          =   360
      Index           =   6
      Left            =   2160
      Picture         =   "frmMain.frx":4678
      Top             =   5280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonN 
      Height          =   360
      Index           =   7
      Left            =   2520
      Picture         =   "frmMain.frx":4D7A
      Top             =   5280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonN 
      Height          =   360
      Index           =   10
      Left            =   3720
      Picture         =   "frmMain.frx":547C
      Top             =   5280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonN 
      Height          =   360
      Index           =   11
      Left            =   4080
      Picture         =   "frmMain.frx":5B7E
      Top             =   5280
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgButtonD 
      Height          =   360
      Index           =   9
      Left            =   3360
      Picture         =   "frmMain.frx":6400
      Top             =   4800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonD 
      Height          =   360
      Index           =   1
      Left            =   120
      Picture         =   "frmMain.frx":6B02
      Top             =   4800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonD 
      Height          =   360
      Index           =   7
      Left            =   2520
      Picture         =   "frmMain.frx":7204
      Top             =   4800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonD 
      Height          =   360
      Index           =   4
      Left            =   1320
      Picture         =   "frmMain.frx":7906
      Top             =   4800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonD 
      Height          =   360
      Index           =   2
      Left            =   600
      Picture         =   "frmMain.frx":8008
      Top             =   4800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonD 
      Height          =   360
      Index           =   11
      Left            =   4080
      Picture         =   "frmMain.frx":870A
      Top             =   4800
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgButtonD 
      Height          =   360
      Index           =   5
      Left            =   1680
      Picture         =   "frmMain.frx":8F8C
      Top             =   4800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonD 
      Height          =   360
      Index           =   3
      Left            =   960
      Picture         =   "frmMain.frx":968E
      Top             =   4800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonD 
      Height          =   360
      Index           =   6
      Left            =   2160
      Picture         =   "frmMain.frx":9D90
      Top             =   4800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonD 
      Height          =   360
      Index           =   10
      Left            =   3720
      Picture         =   "frmMain.frx":A492
      Top             =   4800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonO 
      Height          =   360
      Index           =   1
      Left            =   120
      Picture         =   "frmMain.frx":AB94
      Top             =   4320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonO 
      Height          =   360
      Index           =   7
      Left            =   2520
      Picture         =   "frmMain.frx":B296
      Top             =   4320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonO 
      Height          =   360
      Index           =   4
      Left            =   1320
      Picture         =   "frmMain.frx":B998
      Top             =   4320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonO 
      Height          =   360
      Index           =   2
      Left            =   600
      Picture         =   "frmMain.frx":C09A
      Top             =   4320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonO 
      Height          =   360
      Index           =   11
      Left            =   4080
      Picture         =   "frmMain.frx":C79C
      Top             =   4320
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgButtonO 
      Height          =   360
      Index           =   5
      Left            =   1680
      Picture         =   "frmMain.frx":D01E
      Top             =   4320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonO 
      Height          =   360
      Index           =   3
      Left            =   960
      Picture         =   "frmMain.frx":D720
      Top             =   4320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonO 
      Height          =   360
      Index           =   6
      Left            =   2160
      Picture         =   "frmMain.frx":DE22
      Top             =   4320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonO 
      Height          =   360
      Index           =   10
      Left            =   3720
      Picture         =   "frmMain.frx":E524
      Top             =   4320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgButtonO 
      Height          =   360
      Index           =   9
      Left            =   3360
      Picture         =   "frmMain.frx":EC26
      Top             =   4320
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim TextChange As Boolean
Dim CurrentNote As String
Dim ExecuteClick As Boolean
Dim DragSelStart As Long
Dim DragSelLength As Long

'See if Delete Pressed In filMain
Dim filMainKeydown As Boolean

Private Sub filMain_Click()
On Error GoTo FileOpenErr
Dim Successful As Boolean
If Len(txtMain.Text) <> FileLen(filMain.Path + "\" + filMain.FileName) Then ExecuteClick = True
If ExecuteClick = False Then
    ExecuteClick = True
Else
    Successful = ClickCode 'Call ClickCode Function - Returns True or False
    If Successful = False Then
    filMain.ListIndex = filMain.ListIndex - 1 'If it Don't Exist, Go To One Before
    End If
    txtMain.SelStart = 0
    txtMain.SelLength = 0
End If
Exit Sub
FileOpenErr:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.filMain_Click", Err.Number, Err.Description
RestoreDatabase
Resume
End Sub

Private Sub filMain_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
filMainKeydown = True
DeleteNote
ElseIf KeyCode = vbKeyHome Then
    If SalamanderExists = True Then Salamander.ReportError "TextMe", "Easter Egg!", 2003, "Hey Gabie, Do You Have Your Car Here Today?"
End If
End Sub

Private Sub filMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseRollover
End Sub

Private Sub filMain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.GetFormat(vbCFFiles) = True Then 'It Is A File
    If CheckValid(Data.Files.Item(1)) = True Then 'It Is A Valid Text File
        DragCode (Data.Files.Item(1)) 'Add It To Database
        txtMain.SetFocus
    Else
        InvalidDrag 'Invoke Invalid Drag Error Message
    End If
ElseIf Data.GetFormat(vbCFText) = True Then 'It Is A String
    CreateDrag (Data.GetData(vbCFText)) 'Create A New Note With It
Else
    InvalidDrag 'Invoke Invalid Drag Error Message
End If
End Sub

Private Sub Form_Activate()
On Error GoTo ActivateErr
filMain.SetFocus
DragSelStart = -1
ActivateErr:
End Sub

Private Sub Form_GotFocus()
filMain.SetFocus
End Sub

Private Sub Form_Load()
Dim cmd As String
Dim i As Long
On Error GoTo StartErr
MouseRollover 'Prepare Toolbar
Me.Caption = App.Title 'Set Window Caption
GetRegistrySettings 'Load Registry Settings
txtMain.Locked = ReadOnly 'Set Readonly Setting
'Apply SubDivider Image to all subdividers (saves program size when compiled)
For i = 2 To imgDivider.UBound
    imgDivider(i).Picture = imgDivider(1).Picture
Next i
ExecuteClick = True 'Next Click In filMain Will Be Parsed
filMain.Path = App.Path + "\textfiles\" 'Load TextFiles Directory For filMain
filRefresh.Path = App.Path + "\TextFiles\" 'Load TextFiles Directory For filRefresh
TextChange = False 'TextBox Has Not Changed
If filMain.ListCount = 0 Then RestoreDatabase 'If Empty, Restore Database
filMain.ListIndex = 0
GetStartFile 'Load File Specified In COMMAND$
If App.PrevInstance = True Then End 'If Already Running, File Has Been Added To Database, Quit.
Exit Sub 'Finish Loading
StartErr:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.Form_Load", Err.Number, Err.Description
If Err.Number = 76 Then
    MkDir App.Path + "\TextFiles"
    Resume
Else
MsgBox "TextMe Could Not Start. Please Check That The Textfiles Folder Exists In The Same Folder As TextMe and Registry Access Is Enabled", vbCritical + vbApplicationModal + vbOKOnly, "Cannot Start TextMe"
End If
End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseRollover
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.GetFormat(vbCFFiles) = True Then 'It Is A File
    If CheckValid(Data.Files.Item(1)) = True Then 'It Is A Valid Text File
        DragCode (Data.Files.Item(1)) 'Add It To Database
        txtMain.SetFocus
    Else
        InvalidDrag 'Invoke Invalid Drag Error Message
    End If
ElseIf Data.GetFormat(vbCFText) = True Then 'It Is A String
    CreateDrag (Data.GetData(vbCFText)) 'Create A New Note With It
Else
    InvalidDrag 'Invoke Invalid Drag Error Message
End If
End Sub

Private Sub Form_Resize()
On Error GoTo ResizeError
If filMain.Height < 19 Then
    filMain.Height = 19
ElseIf filMain.Height > 18 Then
    If txtMain.Height + Toolbar.Height <> ScaleHeight Then filMain.Height = ScaleHeight - Toolbar.Height + 1 'The +1 is to make up for the moving of filMain up one to make it level with txtMain
End If
If txtMain.Height + Toolbar.Height <> ScaleHeight Then txtMain.Height = ScaleHeight - Toolbar.Height
If Toolbar.Width <> ScaleWidth Then txtMain.Width = ScaleWidth - filMain.Width
If Toolbar.Width <> ScaleWidth Then Toolbar.Width = ScaleWidth
Exit Sub
ResizeError:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.Form_Resize", Err.Number, Err.Description
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim MSG As Long
Dim F As Long
On Error GoTo NoUnloadSaveCode
If TextChange = True Then
    MSG = MsgBox("Save Changes To " + CurrentNote + "?", vbExclamation + vbYesNoCancel + vbApplicationModal, "Save Note")
    If MSG = vbCancel Then
        Cancel = 1
        Exit Sub
    End If
    If MSG = vbYes Then
        F = FreeFile
        If Right$(txtMain.Text, 2) = vbCrLf Then txtMain.Text = Left$(txtMain.Text, Len(txtMain.Text) - 2)
        Open filMain.Path + "\" + CurrentNote For Output As F
            Print #F, txtMain.Text
        Close F
        TextChange = False
        txtMain.SetFocus
        GoTo ExitTextMe
NoUnloadSaveCode:
        If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.Form_Unload", Err.Number, Err.Description
        MsgBox "Cannot Save " + CurrentNote, vbExclamation + vbOKOnly + vbApplicationModal, "Cannot Save Note"
    End If
End If
ExitTextMe:
SaveRegistrySettings
End
End Sub

Private Sub CopyNote()
Dim copyname As String
Dim F As Long
Dim E As Long
Dim MSG As Long
Dim notestatus As Long
On Error GoTo CopyErr
copyname = InputBox$("Copy " + filMain.FileName + " To", "Copy To A New Note")
If copyname = "" Then Exit Sub
If copyname = filMain.FileName Then Exit Sub
SaveCode
For E = 0 To filMain.ListCount
    If LCase$(filMain.List(E)) = LCase$(copyname) Then
        MSG = MsgBox("Replace Note " + filMain.List(E) + "?", vbExclamation + vbYesNo + vbApplicationModal, "Copy To A New Note")
        If MSG = vbNo Then
            Exit Sub
        Else
            notestatus = 1
            Exit For
        End If
        Exit Sub
    End If
Next E
F = FreeFile
If Right$(txtMain.Text, 2) = vbCrLf Then txtMain.Text = Left$(txtMain.Text, Len(txtMain.Text) - 2)
TextChange = False
Open filMain.Path + "\" + copyname For Output As F
    Print #F, txtMain.Text
Close F
filMain.Refresh
If notestatus = 1 Then
    filMain.ListIndex = E
Else
    For F = 0 To filMain.ListCount
        If filMain.List(F) = copyname Then filMain.ListIndex = F
    Next F
End If
filMain.SetFocus
CurrentNote = filMain.FileName
Exit Sub
CopyErr:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.imgCopy_MouseUp", Err.Number, Err.Description
MsgBox "Note " + filMain.FileName + " Could Not Be Copied", vbExclamation + vbOKOnly + vbApplicationModal, "Cannot Copy Note"
End Sub

Private Sub DeleteNote(Optional Shift As Integer)
DeleteFile:
On Error GoTo NoDelete
'Delete All
Dim MSG As Long
If Shift = 1 Then
    Dim sFile() As String
    sFile = AllFiles(App.Path + "\TextFiles\")
    If sFile(0) = "" Then
        RestoreDatabase
    End If
    MSG = MsgBox("Are You Sure You Want To Delete All Notes In The Database?", vbExclamation + vbApplicationModal + vbYesNo, "Delete All Notes")
    If MSG = vbYes Then
        sFile = AllFiles(App.Path + "\TextFiles\")
        If sFile(0) <> "" Then Kill App.Path + "\TextFiles\*.*"
        RestoreDatabase
    End If
    Exit Sub
End If
'Delete Selected Note
If filMain.ListIndex = -1 Then Exit Sub
Dim i As Long
Dim j As Long
MSG = MsgBox("Delete Note " + filMain.FileName + "?", vbExclamation + vbYesNo + vbApplicationModal, "Delete Note")
If MSG = vbYes Then
    i = filMain.ListIndex
    j = filMain.ListCount
    Kill filMain.Path + "\" + filMain.FileName
    RefreshDatabase
    If filMain.ListCount = 0 Then RestoreDatabase
    If filMain.ListCount = 0 Then
        Do
            CreateNewNote
            filMain.Refresh
        Loop Until filMain.ListCount > 0
        Exit Sub
    End If
    TextChange = False
    If i = 0 Then
        filMain.ListIndex = 0
    ElseIf i = j Then
        filMain.ListIndex = filMain.ListCount
    Else
        filMain.ListIndex = i - 1
    End If
End If
CurrentNote = filMain.FileName
Exit Sub
NoDelete:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.imgDelete_MouseUp", Err.Number, Err.Description
MsgBox "Cannot Delete Note", vbInformation + vbOKOnly, "Cannot Delete Note"
filMain.Refresh
End Sub

Private Sub imgButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgButton(Index).Picture = imgButtonD(Index).Picture
End Sub

Private Sub imgButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case imgButton(Index).Picture
Case imgButtonN(Index).Picture
    MouseRollover (Index)
    imgButton(Index).Picture = imgButtonO(Index).Picture
Case imgButtonD(Index).Picture
    If X < 0 Or X > imgButton(Index).Width Or Y < 0 Or Y > imgButton(Index).Height Then imgButton(Index).Picture = imgButtonO(Index).Picture
Case imgButtonO(Index).Picture
    If Button > 0 And ((X > 0 And X < imgButton(Index).Width) And (Y > 0 And Y < imgButton(Index).Height)) Then imgButton(Index).Picture = imgButtonD(Index).Picture
End Select
End Sub

Private Sub imgButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If imgButton(Index).Picture <> imgButtonO(Index).Picture Then
    Select Case Index
    Case 1
        SaveNote
    Case 2
        CreateNewNote
    Case 3
        CopyNote
    Case 4
        RenameNote
    Case 5
        DeleteNote (Shift)
    Case 6
        SaveCode
        frmBackup.Show vbModal, Me
    Case 7
        SaveCode
        frmBackup.Show vbModal, Me
    Case 8
        frmOptions.Show vbModal, Me
    Case 9
        On Error Resume Next
        Unload frmUpdate
        frmUpdate.Show vbModal, Me
    Case 10
        frmAbout.Show vbModal, Me
    Case 11
        Unload Me
    End Select
    If imgButton(Index).Picture = imgButtonD(Index).Picture Then imgButton(Index).Picture = imgButtonO(Index).Picture
Else
    imgButton(Index).Picture = imgButtonN(Index).Picture
End If
End Sub

Private Sub imgButton_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.GetFormat(vbCFFiles) = True Then 'It Is A File
    If CheckValid(Data.Files.Item(1)) = True Then 'It Is A Valid Text File
        DragCode (Data.Files.Item(1)) 'Add It To Database
        txtMain.SetFocus
    Else
        InvalidDrag 'Invoke Invalid Drag Error Message
    End If
ElseIf Data.GetFormat(vbCFText) = True Then 'It Is A String
    CreateDrag (Data.GetData(vbCFText)) 'Create A New Note With It
Else
    InvalidDrag 'Invoke Invalid Drag Error Message
End If
End Sub

Private Sub imgDivider_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then Me.MousePointer = 5
End Sub

Private Sub imgDivider_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseRollover
End Sub

Private Sub imgDivider_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then Me.MousePointer = 0
End Sub

Private Sub imgDivider_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.GetFormat(vbCFFiles) = True Then 'It Is A File
    If CheckValid(Data.Files.Item(1)) = True Then 'It Is A Valid Text File
        DragCode (Data.Files.Item(1)) 'Add It To Database
        txtMain.SetFocus
    Else
        InvalidDrag 'Invoke Invalid Drag Error Message
    End If
ElseIf Data.GetFormat(vbCFText) = True Then 'It Is A String
    CreateDrag (Data.GetData(vbCFText)) 'Create A New Note With It
Else
    InvalidDrag 'Invoke Invalid Drag Error Message
End If
End Sub

Private Sub CreateNewNote()
BeginNew:
Dim createname As String
Dim notestatus As Long
Dim F As Long
Dim E As Long
Dim MSG As Long
On Error GoTo CreateErr
filRefresh.Refresh
If filRefresh.ListCount > 0 And filMain.ListCount = 0 Then
    filMain.Refresh
    GoTo ExitCreate
End If
If filMain.ListCount = 0 Then txtMain.Text = ""
createname = InputBox$("Create Note", "Create A Blank Note")
If createname = "" Then GoTo ExitCreate
If filMain.ListCount <> 0 Then SaveCode
For E = 0 To filMain.ListCount
    If LCase$(filMain.List(E)) = LCase$(createname) Then
        MSG = MsgBox("Replace Note " + filMain.List(E) + "?", vbExclamation + vbYesNo + vbApplicationModal, "Create A Blank Note")
        If MSG = vbNo Then
            GoTo ExitCreate
        Else
            notestatus = 1
            Exit For
        End If
        GoTo ExitCreate
    End If
Next E
F = FreeFile
'Next Two Lines Check file Was Created
Open filMain.Path + "\" + createname For Output As F
Close F
filMain.Refresh
TextChange = False
If notestatus = 1 Then
    filMain.ListIndex = E
Else
    For F = 0 To filMain.ListCount
        If filMain.List(F) = createname Then filMain.ListIndex = F
    Next F
End If
CurrentNote = filMain.FileName
If Me.Visible = True Then filMain.SetFocus
GoTo ExitCreate
CreateErr:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.imgCreate_MouseUp", Err.Number, Err.Description
If filMain.ListCount = 0 Then
    MSG = MsgBox("Note " + createname + " Could Not Be Created and The Note Database Is Empty. This Could Mean There Database Corruption." + vbCrLf + "Exit TextMe?", vbCritical + vbYesNo + vbApplicationModal, "Cannot Create Blank Note")
    If MSG = vbYes Then End
End If
MsgBox "Note " + createname + " Could Not Be Created.", vbExclamation + vbOKOnly + vbApplicationModal, "Cannot Create Blank Note"
ExitCreate:
If filMain.ListCount = 0 Then GoTo BeginNew
End Sub
Private Sub RenameNote()
Dim newname As String
Dim notestatus As Long
Dim F As Long
Dim E As Long
Dim MSG As Long
If filMain.ListIndex = -1 Then Exit Sub
On Error GoTo nameerr
newname = InputBox$("Please enter a new name for " + filMain.FileName, "Rename A Note", filMain.FileName)
If newname = "" Then Exit Sub
If newname = filMain.FileName Then Exit Sub
SaveCode
For E = 0 To filMain.ListCount
    If LCase$(filMain.List(E)) = LCase$(newname) And LCase(filMain.FileName) <> LCase(newname) Then
        MSG = MsgBox("Replace Note " + filMain.List(E) + "?", vbExclamation + vbYesNo + vbApplicationModal, "Rename A Note")
        If MSG = vbNo Then
            Exit Sub
        Else
            notestatus = 1
            Kill filMain.Path + "\" + newname
            Exit For
        End If
        Exit Sub
    End If
Next E
Name filMain.Path + "\" + filMain.FileName As filMain.Path + "\" + newname
filMain.Refresh
TextChange = False
If notestatus = 1 Then
    filMain.ListIndex = E
Else
    For F = 0 To filMain.ListCount
        If filMain.List(F) = newname Then filMain.ListIndex = F
    Next F
End If
filMain.SetFocus
CurrentNote = filMain.FileName
Exit Sub
nameerr:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.imgRename_MouseUp", Err.Number, Err.Description
MsgBox "Cannot Rename " + filMain.FileName + " to " + newname, vbInformation + vbApplicationModal + vbOKOnly, "Cannot Rename Note"
End Sub

Private Sub SaveNote()
Dim F As Long
If filMain.ListIndex = -1 Then Exit Sub
On Error GoTo NoSave
F = FreeFile
If Right$(txtMain.Text, 2) = vbCrLf Then txtMain.Text = Left$(txtMain.Text, Len(txtMain.Text) - 2)
Open filMain.Path + "\" + CurrentNote For Output As F
    Print #F, txtMain.Text
Close F
TextChange = False
txtMain.SetFocus
Exit Sub
NoSave:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.imgSave_MouseUp", Err.Number, Err.Description
MsgBox "Cannot Save " + CurrentNote, vbExclamation + vbOKOnly + vbApplicationModal, "Cannot Save Note"
End Sub

Private Sub InvalidDrag()
MsgBox "The File Was Not A Valid Text File. TextMe Can Only Read Plain Text Files, Not Microsoft Word Or Rich Text Files.", vbExclamation + vbOKOnly + vbApplicationModal, "Cannot Import Note"
End Sub

Private Sub timMain_Timer()
If frmBackup.Visible = True And filMain.ListCount = 0 Then Exit Sub
RefreshDatabase 'Because Of A Debugging Exercise It Got Moved To A SUB
End Sub

Private Sub Toolbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseRollover
End Sub

Private Sub Toolbar_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.GetFormat(vbCFFiles) = True Then 'It Is A File
    If CheckValid(Data.Files.Item(1)) = True Then 'It Is A Valid Text File
        DragCode (Data.Files.Item(1)) 'Add It To Database
        txtMain.SetFocus
    Else
        InvalidDrag 'Invoke Invalid Drag Error Message
    End If
ElseIf Data.GetFormat(vbCFText) = True Then 'It Is A String
    CreateDrag (Data.GetData(vbCFText)) 'Create A New Note With It
Else
    InvalidDrag 'Invoke Invalid Drag Error Message
End If
End Sub

Private Sub txtMain_Change()
TextChange = True
End Sub

Private Sub txtMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseRollover
End Sub

Private Sub txtMain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim SelPlace As Long
Dim foo As Variant
Dim SelPlaceLength As Long
If Data.GetFormat(vbCFFiles) = True Then
    If CheckValid(Data.Files.Item(1)) = False Then
        InvalidDrag
        Exit Sub
    End If
    DragCode (Data.Files.Item(1))
    txtMain.SetFocus
ElseIf Data.GetFormat(vbCFText) = True Then
    If txtMain.Locked = True Then Exit Sub
    If DragSelStart = -1 Then
        txtMain.SelText = Data.GetData(vbCFText)
    Else
        SelPlace = txtMain.SelStart
        txtMain.SelStart = DragSelStart
        txtMain.SelLength = DragSelLength
        foo = Data.GetData(vbCFText)
        If txtMain.SelText = Data.GetData(vbCFText) Then
            txtMain.SelStart = SelPlace
            txtMain.SelLength = SelPlaceLength
            txtMain.SelText = Data.GetData(vbCFText)
            txtMain.SelStart = DragSelStart
            txtMain.SelLength = DragSelLength
            If txtMain.SelText = Data.GetData(vbCFText) Then
                txtMain.SelText = ""
            Else
               txtMain.SelStart = DragSelStart + DragSelLength
                txtMain.SelLength = DragSelLength
            txtMain.SelText = ""
            End If
            txtMain.SelStart = SelPlace + DragSelLength
        Else
            txtMain.SelText = txtMain.SelText + Data.GetData(vbCFText)
        End If
    End If
Else
    InvalidDrag
    Exit Sub
End If
DragSelStart = -1
End Sub

Private Sub txtMain_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    DragSelStart = txtMain.SelStart
    DragSelLength = txtMain.SelLength
End Sub

Function CheckValid(file As String) As Boolean
On Error GoTo NotValid
Dim F As Long
Dim s As String
F = FreeFile
Open file For Input As F
s = Input$(LOF(F), F)
Close F
CheckValid = True
Exit Function
NotValid:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.CheckValid()", Err.Number, Err.Description
CheckValid = False
End Function

Function ClickCode() As Boolean
On Error GoTo loaderr
Dim F As Long
Dim s As String
Dim nRet As Long
Dim OpeningFile As Boolean
Const WM_SETTEXT = &HC
If TextChange = True Then SaveCode
If Len(txtMain.Text) = FileLen(filMain.Path + "\" + filMain.FileName) Then
    ClickCode = True
    Exit Function
End If
F = FreeFile
Open filMain.Path + "\" + filMain.FileName For Input As F
OpeningFile = True
s = Input$(LOF(F), F)
Close F
OpeningFile = False
nRet = SendMessage(txtMain.hWnd, WM_SETTEXT, 0&, ByVal s)
nRet = SetWindowText(txtMain.hWnd, s)
CurrentNote = filMain.FileName
TextChange = False
ClickCode = True
Exit Function
loaderr:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.ClickCode()", Err.Number, Err.Description
Dim MSG As Long
MSG = MsgBox("Cannot Open Note " + filMain.FileName + vbCrLf + vbCrLf + "It May Be Corrupted, Do You Want To Delete It?", vbExclamation + vbYesNo + vbApplicationModal, "Cannot Open Note")
    If MSG = vbYes Then
        If OpeningFile = True Then Close F
        If Dir(App.Path + "\textfiles\" + filMain.FileName) = filMain.FileName Then Kill App.Path + "\TextFiles\" + filMain.FileName
        RefreshDatabase
    End If
If OpeningFile = True Then Close F
ClickCode = False
End Function

Sub CreateDrag(Data As String)
Dim createname As String
Dim notestatus As Long
Dim F As Long
Dim E As Long
Dim MSG As Long
On Error GoTo CreateDragErr
createname = InputBox$("Create Note", "Create A Note")
If createname = "" Then Exit Sub
SaveCode
For E = 0 To filMain.ListCount
    If LCase$(filMain.List(E)) = LCase$(createname) Then
        MSG = MsgBox("Replace Note " + filMain.List(E) + "?", vbExclamation + vbYesNo + vbApplicationModal, "Create Note")
        If MSG = vbNo Then
            Exit Sub
        Else
            notestatus = 1
            Exit For
        End If
        Exit Sub
    End If
Next E
F = FreeFile
Open filMain.Path + "\" + createname For Output As F
Print #F, Data
Close F
filMain.Refresh
TextChange = False
If notestatus = 1 Then
    filMain.ListIndex = E
Else
    For F = 0 To filMain.ListCount
        If filMain.List(F) = createname Then filMain.ListIndex = F
    Next F
End If
filMain.SetFocus
Exit Sub
CreateDragErr:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.CreateDrag()", Err.Number, Err.Description
MsgBox "Note " + createname + " Could Not Be Created", vbExclamation + vbOKOnly + vbApplicationModal, "Cannot Create Note"
End Sub

Private Sub DragCode(Data As String)
Dim file As Variant
Dim i As Long
Dim E As Long
Dim F As Long
Dim MSG As Long
Dim notestatus As Long
On Error GoTo DragErr
notestatus = 0
file = Split(Data, "\", -1, vbTextCompare)
i = UBound(file)
For E = 0 To filMain.ListCount
    If LCase$(filMain.List(E)) = LCase$(file(i)) Then
        If Data = App.Path + "\textfiles\" + file(i) Then Exit Sub
        MSG = MsgBox("Replace Note " + filMain.List(E) + "?", vbExclamation + vbYesNo + vbApplicationModal, "Import Into TextMe")
        If MSG = vbNo Then
            Exit Sub
        Else
            notestatus = 1
            Exit For
        End If
        Exit Sub
    End If
Next E
FileCopy Data, App.Path + "\textfiles\" + file(i)
If DeleteOriginal = 1 Then Kill Data
filMain.Refresh
TextChange = False
If notestatus = 1 Then
    filMain.Selected(E) = True
Else
    For F = 0 To filMain.ListCount
        If filMain.List(F) = file(i) Then filMain.Selected(F) = True
    Next F
End If
Exit Sub
DragErr:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.DragCode()", Err.Number, Err.Description
MsgBox "Invalid File, File Already Open or File Already Exists In Database", vbApplicationModal + vbOKOnly + vbCritical, "Cannot Import Into TextMe"
End Sub

Private Sub GetStartFile()
On Error GoTo StartFileErr
Dim Valid As Boolean
Dim cmd As String
'Remove quotation marks, check command$ <> ""
cmd = Command$
If cmd = "" Then Exit Sub
cmd = Trim$(cmd)
If Left$(cmd, 1) = """" Then cmd = Right$(cmd, Len(cmd) - 1)
If Right$(cmd, 1) = """" Then cmd = Left$(cmd, Len(cmd) - 1)
'Check if valid file
Valid = CheckValid(cmd)
If Valid = False Then
    InvalidDrag
Exit Sub
End If
DragCode (cmd)
StartFileErr:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.GetStartFile()", Err.Number, Err.Description
End Sub

Sub MouseRollover(Optional Index As Integer)
Dim i As Integer
For i = 1 To 11
    If i <> Index Then If imgButton(i).Picture <> imgButtonN(i).Picture Then imgButton(i).Picture = imgButtonN(i).Picture
Next i
End Sub
Private Sub RefreshDatabase()
On Error GoTo CorruptErr
Dim CurSel As String
Dim CurSelPos As Long
Dim F As Long
Dim sFiles() As String
Dim N As Long
'Get list of files right now
filRefresh.Refresh
If filRefresh.List(0) = "" Then GoTo RefreshCode
If filRefresh.ListCount = filMain.ListCount Then
    For N = 0 To filRefresh.ListCount
        If filRefresh.List(N) <> filMain.List(N) Then GoTo RefreshCode
    Next
    Exit Sub
End If
RefreshCode:
'Refresh Filelistbox
ExecuteClick = False
CurSel = filMain.FileName
CurSelPos = filMain.ListIndex
filMain.Refresh
If filMain.ListCount = 0 Then RestoreDatabase
filMain.Selected(0) = True
ExecuteClick = False
If Me.Visible = True Then
    For F = 0 To filMain.ListCount
        If filMain.List(F) = CurSel Then
            filMain.Selected(F) = True
            Exit For
        Else
            If F = filMain.ListCount Then
                If filMain.List(F) <> CurSel Then
                    If CurSelPos >= filMain.ListCount Then
                        filMain.Selected(filMain.ListCount - 1) = True
                    Else
                        filMain.Selected(CurSelPos) = True
                    End If
                End If
            End If
        End If
    Next F
End If
If CurrentNote <> filMain.FileName Then SaveCode
ExecuteClick = True
Exit Sub
CorruptErr:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.RefreshDatabase()", Err.Number, Err.Description
If filMain.ListCount = 0 Then RestoreDatabase
If filMain.ListCount > 0 Then Exit Sub
'This Should Never, Ever Happen (It Might!!)
Dim MSG As Long
MSG = MsgBox("WARNING: The TextMe Database has been corrupted!" + vbCrLf + "Would you like to wipe the TextMe Database?" + vbCrLf + vbCrLf + "This Operation Will Result In All Your Notes Being Destroyed!" + vbCrLf + vbCrLf + "Otherwise, you may open the TextFiles folder in the TextMe Program folder to find the file causing the corruption", vbCritical + vbApplicationModal + vbYesNo, "WARNING: Database Corruption")
If MSG = vbNo Then
    End
Else
    Kill App.Path + "\textfiles\*.*"
    filMain.Refresh
    GoTo CorruptErr
End If
End Sub
Sub RestoreDatabase()
On Error GoTo NoRestore
'Check If TextFiles Folder Deleted
ChDir App.Path + "\textfiles"
ChDir App.Path
'Ask for Refresh, etc
AskForRefresh:
Dim MSG As Long
filMain.Refresh
txtMain.Text = ""
MSG = MsgBox("WARNING: The TextMe Database is empty!" + vbCrLf + "Would you like to restore from a previous backup?" + vbCrLf + vbCrLf + "Otherwise, click No to create a blank note or Cancel to exit TextMe", vbCritical + vbApplicationModal + vbYesNoCancel, "Restore TextMe Database")
If MSG = vbYes Then
    'imgRestore.Picture = imgRestoreD.Picture
    frmBackup.Show vbModal, Me
    Do Until filMain.ListCount > 0
        CreateNewNote
        filMain.Refresh
    Loop
    filMain.Refresh
    filRefresh.Refresh
    TextChange = False
    filMain.Selected(0) = True
ElseIf MSG = vbNo Then
    Do Until filMain.ListCount > 0
        CreateNewNote
        filMain.Refresh
    Loop
Else
    Unload Me
End If
Exit Sub
NoRestore:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.RestoreDatabase()", Err.Number, Err.Description
If Err.Number = 76 Then
    MkDir App.Path + "\TextFiles"
    Resume AskForRefresh
Else
    MsgBox "Unknown Error Number " + CStr(Err.Number) + "Has Occured." + vbCrLf + "Please Report This To Peter Chapman, Or Update TextMe To See If This Problem has Been Fixed.", vbApplicationModal + vbOKOnly + vbCritical, App.Title
End If
End Sub
Sub SaveCode()
Dim MSG As Long
Dim F As Long
On Error GoTo NoSaveCode
If frmBackup.Visible = True Then Exit Sub
If TextChange = True Then
    MSG = MsgBox("Save Changes To " + CurrentNote + "?", vbExclamation + vbYesNo + vbApplicationModal, "Save Note")
    If MSG = vbYes Then
    F = FreeFile
    If Right$(txtMain.Text, 2) = vbCrLf Then txtMain.Text = Left$(txtMain.Text, Len(txtMain.Text) - 2)
        Open filMain.Path + "\" + CurrentNote For Output As F
            Print #F, txtMain.Text
        Close F
        TextChange = False
        txtMain.SetFocus
        Exit Sub
NoSaveCode:
        If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.SaveCode()", Err.Number, Err.Description
        MsgBox "Cannot Save " + CurrentNote, vbExclamation + vbOKOnly + vbApplicationModal, "Cannot Save Note"
    End If
End If
End Sub
Public Sub GetRegistrySettings()
Dim iWindowState As Integer
Dim MSG As Integer
On Error GoTo NoGet
DeleteOriginal = GetSetting("Peter Chapman", "TextMe", "DeleteOriginal", "0")
ReadOnly = GetSetting("Peter Chapman", "TextMe", "ReadOnly", "0")
iWindowState = GetSetting("Peter Chapman", "TextMe", "WindowState", 0)
If iWindowState = 0 Then
    Me.Top = GetSetting("Peter Chapman", "TextMe", "Top", Me.Top)
    Me.Left = GetSetting("Peter Chapman", "TextMe", "Left", Me.Left)
    Me.Height = GetSetting("Peter Chapman", "TextMe", "Height", Me.Height)
    Me.Width = GetSetting("Peter Chapman", "TextMe", "Width", Me.Width)
Else
    Me.WindowState = iWindowState
End If
Exit Sub
NoGet:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.GetRegistrySettings()", Err.Number, Err.Description
ResetRegistrySettings
End Sub
Public Sub SaveRegistrySettings()
On Error GoTo NoSaveRegistry
SaveSetting "Peter Chapman", "TextMe", "DeleteOriginal", Val(DeleteOriginal)
SaveSetting "Peter Chapman", "TextMe", "ReadOnly", Val(ReadOnly)
If Me.WindowState = vbMaximized Or Me.WindowState = vbMinimized Then
    SaveSetting "Peter Chapman", "TextMe", "WindowState", Me.WindowState
Else
    SaveSetting "Peter Chapman", "TextMe", "Left", Me.Left
    SaveSetting "Peter Chapman", "TextMe", "Top", Me.Top
    SaveSetting "Peter Chapman", "TextMe", "Height", Me.Height
    SaveSetting "Peter Chapman", "TextMe", "Width", Me.Width
    SaveSetting "Peter Chapman", "TextMe", "Windowstate", Me.WindowState
End If
Exit Sub
NoSaveRegistry:
MsgBox "Cannot Save Registry Settings, Please Contact Your Network Administrator", vbCritical + vbDefaultButton1 + vbApplicationModal + vbOKOnly, App.Title
End Sub
Public Sub ResetRegistrySettings()
On Error GoTo NoResetRegistry
SaveSetting "Peter Chapman", "TextMe", "DeleteOriginal", 0
SaveSetting "Peter Chapman", "TextMe", "ReadOnly", 0
txtMain.Locked = False
If Me.WindowState = vbMaximized Or Me.WindowState = vbMinimized Then
    SaveSetting "Peter Chapman", "TextMe", "Windowstate", Me.WindowState
Else
    SaveSetting "Peter Chapman", "TextMe", "Left", Me.Left
    SaveSetting "Peter Chapman", "TextMe", "Top", Me.Top
    SaveSetting "Peter Chapman", "TextMe", "Height", Me.Height
    SaveSetting "Peter Chapman", "TextMe", "Width", Me.Width
    SaveSetting "Peter Chapman", "TextMe", "Windowstate", Me.WindowState
End If
NoResetRegistry:
MsgBox "Cannot Reset Registry Settings, Please Contact Your Network Administrator", vbCritical + vbDefaultButton1 + vbApplicationModal + vbOKOnly, App.Title
End Sub
