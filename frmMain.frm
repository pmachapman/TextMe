VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   4140
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6885
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   276
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   459
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
      Height          =   3765
      Left            =   2160
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   375
      Width           =   3900
   End
   Begin VB.PictureBox Toolbar 
      Appearance      =   0  'Flat
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
         Picture         =   "frmMain.frx":0442
         ToolTipText     =   "Insert The Time & Date"
         Top             =   0
         Width           =   360
      End
      Begin VB.Shape recBorder 
         BackColor       =   &H00E8E6E1&
         BorderColor     =   &H00C56A31&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E2B598&
         Height          =   360
         Index           =   8
         Left            =   3000
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgDivider 
         Height          =   360
         Index           =   4
         Left            =   3400
         OLEDropMode     =   1  'Manual
         Top             =   0
         Width           =   45
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   9
         Left            =   3480
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":07FD
         ToolTipText     =   "TextMe Settings"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   10
         Left            =   3840
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":0BA3
         ToolTipText     =   "Check For Updates To TextMe (Requires Internet Connection)"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   12
         Left            =   4560
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":0F76
         ToolTipText     =   "Exit The TextMe Notes Organiser"
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   11
         Left            =   4200
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":12FB
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
         Picture         =   "frmMain.frx":16AE
         Top             =   0
         Width           =   45
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   7
         Left            =   2520
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":19F7
         ToolTipText     =   "Restore Your Notes"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   6
         Left            =   2160
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":1D8B
         ToolTipText     =   "Backup Your Notes"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   5
         Left            =   1680
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":2120
         ToolTipText     =   "Delete The Open Note (Hold Down Shift To Delete All Notes)"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   4
         Left            =   1320
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":24D4
         ToolTipText     =   "Rename The Open Note"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   3
         Left            =   960
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":285D
         ToolTipText     =   "Copy The Open Note (Including Unsaved Changes) To A New Note"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgDivider 
         Height          =   360
         Index           =   0
         Left            =   30
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":2C0D
         Top             =   0
         Width           =   75
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   2
         Left            =   600
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":2F5E
         ToolTipText     =   "Create A New, Blank Note"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   1
         Left            =   120
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":3318
         ToolTipText     =   "Save Changes To The Open Note"
         Top             =   0
         Width           =   360
      End
      Begin VB.Shape recBorder 
         BackColor       =   &H00E8E6E1&
         BorderColor     =   &H00C56A31&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E2B598&
         Height          =   360
         Index           =   1
         Left            =   120
         Top             =   0
         Width           =   360
      End
      Begin VB.Shape recBorder 
         BackColor       =   &H00E8E6E1&
         BorderColor     =   &H00C56A31&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E2B598&
         Height          =   360
         Index           =   2
         Left            =   600
         Top             =   0
         Width           =   360
      End
      Begin VB.Shape recBorder 
         BackColor       =   &H00E8E6E1&
         BorderColor     =   &H00C56A31&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E2B598&
         Height          =   360
         Index           =   3
         Left            =   960
         Top             =   0
         Width           =   360
      End
      Begin VB.Shape recBorder 
         BackColor       =   &H00E8E6E1&
         BorderColor     =   &H00C56A31&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E2B598&
         Height          =   360
         Index           =   4
         Left            =   1320
         Top             =   0
         Width           =   360
      End
      Begin VB.Shape recBorder 
         BackColor       =   &H00E8E6E1&
         BorderColor     =   &H00C56A31&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E2B598&
         Height          =   360
         Index           =   5
         Left            =   1680
         Top             =   0
         Width           =   360
      End
      Begin VB.Shape recBorder 
         BackColor       =   &H00E8E6E1&
         BorderColor     =   &H00C56A31&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E2B598&
         Height          =   360
         Index           =   6
         Left            =   2160
         Top             =   0
         Width           =   360
      End
      Begin VB.Shape recBorder 
         BackColor       =   &H00E8E6E1&
         BorderColor     =   &H00C56A31&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E2B598&
         Height          =   360
         Index           =   7
         Left            =   2520
         Top             =   0
         Width           =   360
      End
      Begin VB.Shape recBorder 
         BackColor       =   &H00E8E6E1&
         BorderColor     =   &H00C56A31&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E2B598&
         Height          =   360
         Index           =   9
         Left            =   3480
         Top             =   0
         Width           =   360
      End
      Begin VB.Shape recBorder 
         BackColor       =   &H00E8E6E1&
         BorderColor     =   &H00C56A31&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E2B598&
         Height          =   360
         Index           =   10
         Left            =   3840
         Top             =   0
         Width           =   360
      End
      Begin VB.Shape recBorder 
         BackColor       =   &H00E8E6E1&
         BorderColor     =   &H00C56A31&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E2B598&
         Height          =   360
         Index           =   11
         Left            =   4200
         Top             =   0
         Width           =   360
      End
      Begin VB.Shape recBorder 
         BackColor       =   &H00E8E6E1&
         BorderColor     =   &H00C56A31&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E2B598&
         Height          =   360
         Index           =   12
         Left            =   4560
         Top             =   0
         Width           =   435
      End
   End
   Begin VB.FileListBox filMain 
      Height          =   3795
      Left            =   -15
      MultiSelect     =   2  'Extended
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Choose A Note To Open. You Can Also Drag A Text File In Here To Create A Note From The File"
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_ITEMFROMPOINT = &H1A9
Private Const LB_GETTEXT = &H189
Private Const LB_GETTEXTLEN = &H18A

Dim TextChange As Boolean
Dim CurrentNote As String
Dim ExecuteClick As Boolean
Dim DragSelStart As Long
Dim DragSelLength As Long

Private Sub filMain_Click()
On Error GoTo FileOpenErr
Dim Successful As Boolean
If CurrentNote = filMain.List(filMain.ListIndex) Then Exit Sub
If Len(txtMain.Text) <> FileLen(filMain.Path + "\" + filMain.FileName) Then ExecuteClick = True
If ExecuteClick = False Then
    ExecuteClick = True
Else
    Successful = ClickCode 'Call ClickCode Function - Returns True or False
    If Successful = False Then
        filMain.Selected(filMain.ListIndex - 1) = True 'If it Don't Exist, Go To One Before
        filMain.Selected(filMain.ListIndex) = False 'Deselect Invalid Note
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
Select Case KeyCode
Case vbKeyDelete, vbKeyBack
    DeleteNote
Case vbKeyHome
    If SalamanderExists = True Then Salamander.ReportError "TextMe", "Easter Egg!", 2003, "Hey Gabie, Do You Have Your Car Here Today?"
End Select
End Sub

Private Sub filMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseRollover
'Nice Tooltip Code...
'Purpose: Assign the ToolTipText property based on the mouse position
On Error GoTo Err
    Dim lResult As Long
    Dim lParam As Long
    Dim sText As String
    Dim nIndex As Integer

    'Determine the ListBox index of the item beneath the mouse
    lParam = (CInt(Y / Screen.TwipsPerPixelY) * 2 ^ 16) + CInt(X / Screen.TwipsPerPixelX)
    lResult = SendMessage(filMain.hwnd, LB_ITEMFROMPOINT, 0, ByVal lParam)
    
    'The high-order word contains a success/failure flag
    'If (lResult \ 2 ^ 16) <> 0 Then Exit Sub
    
    nIndex = CInt(lResult)
        'Determine the size of the buffer required for the item text
        lResult = SendMessage(filMain.hwnd, LB_GETTEXTLEN, nIndex, ByVal 0)
        
        If (lResult = -1) Then GoTo Err
        
        'Retrieve the item text
        sText = Space(lResult + 1)
        lResult = SendMessage(filMain.hwnd, LB_GETTEXT, nIndex, ByVal sText)
        
        If (lResult = -1) Then GoTo Err
        
        filMain.ToolTipText = Left(sText, lResult)
    Exit Sub
    
Err:
filMain.ToolTipText = ""
End Sub

Private Sub filMain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.GetFormat(vbCFFiles) = True Then 'It Is A File
    Dim i As Integer
    For i = 1 To Data.Files.Count
    If CheckValid(Data.Files.Item(i)) = False Then
        InvalidDrag
    Else
        DragCode Data.Files.Item(i), Shift
    End If
    Next i
    txtMain.SetFocus
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
' Initialize the list of custom colors to white.
Dim C As Integer  ' counter variable
For C = 0 To 15  ' loop through each element
  CustomColours(C) = RGB(255, 255, 255) ' set each element to RGB color 0 (black)
Next C
MouseRollover 'Prepare Toolbar
Me.Caption = App.Title 'Set Window Caption
GetRegistrySettings 'Load Registry Settings
txtMain.Locked = ReadOnly 'Set Readonly Setting
'Apply SubDivider Image to all subdividers (saves program size when compiled)
For i = 2 To imgDivider.UBound
    imgDivider(i).Picture = imgDivider(1).Picture
Next i
ExecuteClick = True 'Next Click In filMain Will Be Parsed
filMain.Path = App.Path + "\TextFiles\" 'Load TextFiles Directory For filMain
filRefresh.Path = App.Path + "\TextFiles\" 'Load TextFiles Directory For filRefresh
TextChange = False 'TextBox Has Not Changed
If filMain.ListCount = 0 Then RestoreDatabase 'If Empty, Restore Database
filMain.Selected(0) = True
filMain_Click
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
    Dim i As Integer
    For i = 1 To Data.Files.Count
    If CheckValid(Data.Files.Item(i)) = False Then
        InvalidDrag
    Else
        DragCode Data.Files.Item(i), Shift
    End If
    Next i
    txtMain.SetFocus
ElseIf Data.GetFormat(vbCFText) = True Then 'It Is A String
    CreateDrag (Data.GetData(vbCFText)) 'Create A New Note With It
Else
    InvalidDrag 'Invoke Invalid Drag Error Message
End If
End Sub

Private Sub Form_Resize()
On Error GoTo ResizeError
If ScaleHeight = 0 Then Exit Sub 'Minimised
If filMain.Height < 19 Then filMain.Height = 19
If filMain.Height > 18 Then
    If Me.ScaleHeight - filMain.Height > 36 Then
        If txtMain.Height + Toolbar.Height <> ScaleHeight Then filMain.Height = ScaleHeight - Toolbar.Height + 2  'The +1 is to make up for the moving of filMain up one to make it level with txtMain
    ElseIf ScaleHeight - filMain.Height - Toolbar.Height < 0 Then
        If ScaleHeight - Toolbar.Height + 2 > 0 Then
            filMain.Height = ScaleHeight - Toolbar.Height + 2
        Else
            Exit Sub
        End If
    End If
End If
If txtMain.Height + Toolbar.Height <> ScaleHeight Then txtMain.Height = ScaleHeight - Toolbar.Height
If Toolbar.Width <> ScaleWidth Then
    If ScaleWidth > filMain.Width Then
        txtMain.Width = ScaleWidth - filMain.Width + 2
    End If
End If
If Toolbar.Width <> ScaleWidth Then Toolbar.Width = ScaleWidth
Exit Sub
ResizeError:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.Form_Resize", Err.Number, Err.Description
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Msg As Long
Dim F As Long
On Error GoTo NoUnloadSaveCode
If TextChange = True Then
    Msg = MsgBox("Save Changes To " + CurrentNote + "?", vbExclamation + vbYesNoCancel + vbApplicationModal, "Save Note")
    If Msg = vbCancel Then
        Cancel = 1
        Exit Sub
    End If
    If Msg = vbYes Then
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
Dim Msg As Long
Dim notestatus As Long
On Error GoTo CopyErr
copyname = InputBox$("Copy " + filMain.FileName + " To", "Copy To A New Note")
If copyname = "" Then Exit Sub
If copyname = filMain.FileName Then Exit Sub
SaveCode
For E = 0 To filMain.ListCount
    If LCase$(filMain.List(E)) = LCase$(copyname) Then
        Msg = MsgBox("Replace Note " + filMain.List(E) + "?", vbExclamation + vbYesNo + vbApplicationModal, "Copy To A New Note")
        If Msg = vbNo Then
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
    filMain.Selected(E) = True
Else
    For F = 0 To filMain.ListCount
        If filMain.List(F) = copyname Then filMain.Selected(F) = True
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
Dim Msg As Long
If Shift = 1 Then
    Dim sFile() As String
    sFile = AllFiles(App.Path + "\TextFiles\")
    If sFile(0) = "" Then
        RestoreDatabase
    End If
    Msg = MsgBox("Are You Sure You Want To Delete All Notes In The Database?", vbExclamation + vbApplicationModal + vbYesNo, "Delete All Notes")
    If Msg = vbYes Then
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
Msg = MsgBox("Delete Note " + filMain.FileName + "?", vbExclamation + vbYesNo + vbApplicationModal, "Delete Note")
If Msg = vbYes Then
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
        filMain.Selected(0) = True
    ElseIf i = j Then
        DeselectAll
        filMain.Selected(filMain.ListCount) = True
    Else
        DeselectAll
        filMain.Selected(i - 1) = True
    End If
End If
CurrentNote = filMain.FileName
Exit Sub
NoDelete:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.imgDelete_MouseUp", Err.Number, Err.Description
MsgBox "Cannot Delete Note", vbInformation + vbOKOnly, "Cannot Delete Note"
filMain.Refresh
End Sub
Private Sub DeselectAll()
Dim i As Integer
For i = 0 To filMain.ListCount - 1
filMain.Selected(i) = False
Next i
End Sub

Private Sub imgButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And X > 0 And X < imgButton(Index).Width And Y > 0 And Y < imgButton(Index).Height Then
    recBorder(Index).FillStyle = 0
    recBorder(Index).BorderStyle = 1
ElseIf X < 0 Or X > imgButton(Index).Width Or Y < 0 Or Y > imgButton(Index).Height Then
    recBorder(Index).BorderStyle = 1
    recBorder(Index).BackStyle = 1
End If
End Sub

Private Sub imgButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Or Button = 0 Or Button = 3 Or Button = 5 Or Button = 7 Then
If recBorder(Index).BorderStyle = 0 Then
    MouseRollover (Index)
    recBorder(Index).BorderStyle = 1
    recBorder(Index).BackStyle = 1
ElseIf recBorder(Index).FillStyle = 0 Then
    If X < 0 Or X > imgButton(Index).Width Or Y < 0 Or Y > imgButton(Index).Height Then
        recBorder(Index).FillStyle = 1
    End If
Else
        If Button > 0 And ((X > 0 And X < imgButton(Index).Width) And (Y > 0 And Y < imgButton(Index).Height)) Then
            recBorder(Index).FillStyle = 0
        End If
End If
End If
End Sub

Private Sub imgButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If recBorder(Index).FillStyle = 0 Then
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
        On Error Resume Next
        frmBackup.Show vbModal, Me
    Case 7
        SaveCode
        On Error Resume Next
        frmBackup.Show vbModal, Me
    Case 8
        txtMain.SelText = Time & " " & Date
    Case 9
        On Error Resume Next
        frmOptions.Show vbModal, Me
    Case 10
        On Error Resume Next
        Unload frmUpdate
        frmUpdate.Show vbModal, Me
    Case 11
        On Error Resume Next
        frmAbout.Show vbModal, Me
    Case 12
        Unload Me
    End Select
    recBorder(Index).FillStyle = 1
'Else
    'recBorder(Index).FillStyle = 1
    'recBorder(Index).BorderStyle = 0
    'recBorder(Index).BackStyle = 0
End If
End Sub

Private Sub imgButton_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.GetFormat(vbCFFiles) = True Then 'It Is A File
    Dim i As Integer
    For i = 1 To Data.Files.Count
    If CheckValid(Data.Files.Item(i)) = False Then
        InvalidDrag
    Else
        DragCode Data.Files.Item(i), Shift
    End If
    Next i
    txtMain.SetFocus
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
    Dim i As Integer
    For i = 1 To Data.Files.Count
    If CheckValid(Data.Files.Item(i)) = False Then
        InvalidDrag
    Else
        DragCode Data.Files.Item(i), Shift
    End If
    Next i
    txtMain.SetFocus
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
Dim Msg As Long
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
        Msg = MsgBox("Replace Note " + filMain.List(E) + "?", vbExclamation + vbYesNo + vbApplicationModal, "Create A Blank Note")
        If Msg = vbNo Then
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
    filMain.Selected(E) = True
Else
    For F = 0 To filMain.ListCount
        If filMain.List(F) = createname Then filMain.Selected(F) = True
    Next F
End If
CurrentNote = filMain.FileName
If Me.Visible = True Then filMain.SetFocus
GoTo ExitCreate
CreateErr:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.imgCreate_MouseUp", Err.Number, Err.Description
If filMain.ListCount = 0 Then
    Msg = MsgBox("Note " + createname + " Could Not Be Created and The Note Database Is Empty. This Could Mean There Database Corruption." + vbCrLf + "Exit TextMe?", vbCritical + vbYesNo + vbApplicationModal, "Cannot Create Blank Note")
    If Msg = vbYes Then End
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
Dim Msg As Long
If filMain.ListIndex = -1 Then Exit Sub
On Error GoTo nameerr
newname = InputBox$("Please enter a new name for " + filMain.FileName, "Rename A Note", filMain.FileName)
If newname = "" Then Exit Sub
If newname = filMain.FileName Then Exit Sub
SaveCode
For E = 0 To filMain.ListCount
    If LCase$(filMain.List(E)) = LCase$(newname) And LCase(filMain.FileName) <> LCase(newname) Then
        Msg = MsgBox("Replace Note " + filMain.List(E) + "?", vbExclamation + vbYesNo + vbApplicationModal, "Rename A Note")
        If Msg = vbNo Then
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
    filMain.Selected(E) = True
Else
    For F = 0 To filMain.ListCount
        If filMain.List(F) = newname Then filMain.Selected(F) = True
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
    Dim i As Integer
    For i = 1 To Data.Files.Count
    If CheckValid(Data.Files.Item(i)) = False Then
        InvalidDrag
    Else
        DragCode Data.Files.Item(i), Shift
    End If
    Next i
    txtMain.SetFocus
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
    Dim i As Integer
    For i = 1 To Data.Files.Count
    If CheckValid(Data.Files.Item(i)) = False Then
        InvalidDrag
    Else
        DragCode Data.Files.Item(i), Shift
    End If
    Next i
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
    CurrentNote = filMain.FileName
    ClickCode = True
    Exit Function
End If
F = FreeFile
Open filMain.Path + "\" + filMain.FileName For Input As F
OpeningFile = True
s = Input$(LOF(F), F)
Close F
OpeningFile = False
nRet = SendMessage(txtMain.hwnd, WM_SETTEXT, 0&, ByVal s)
nRet = SetWindowText(txtMain.hwnd, s)
CurrentNote = filMain.FileName
TextChange = False
ClickCode = True
Exit Function
loaderr:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.ClickCode()", Err.Number, Err.Description
Dim Msg As Long
Msg = MsgBox("Cannot Open Note " + filMain.FileName + vbCrLf + vbCrLf + "It May Be Corrupted, Do You Want To Delete It?", vbExclamation + vbYesNo + vbApplicationModal, "Cannot Open Note")
    If Msg = vbYes Then
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
Dim Msg As Long
On Error GoTo CreateDragErr
createname = InputBox$("Create Note", "Create A Note")
If createname = "" Then Exit Sub
SaveCode
For E = 0 To filMain.ListCount
    If LCase$(filMain.List(E)) = LCase$(createname) Then
        Msg = MsgBox("Replace Note " + filMain.List(E) + "?", vbExclamation + vbYesNo + vbApplicationModal, "Create Note")
        If Msg = vbNo Then
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
    filMain.Selected(E) = True
Else
    For F = 0 To filMain.ListCount
        If filMain.List(F) = createname Then filMain.Selected(F) = True
    Next F
End If
filMain.SetFocus
Exit Sub
CreateDragErr:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.CreateDrag()", Err.Number, Err.Description
MsgBox "Note " + createname + " Could Not Be Created", vbExclamation + vbOKOnly + vbApplicationModal, "Cannot Create Note"
End Sub

Private Sub DragCode(Data As String, Shift As Integer)
Dim file As Variant
Dim i As Long
Dim E As Long
Dim F As Long
Dim Msg As Long
Dim notestatus As Long
On Error GoTo DragErr
notestatus = 0
file = Split(Data, "\", -1, vbTextCompare)
i = UBound(file)
For E = 0 To filMain.ListCount
    If LCase$(filMain.List(E)) = LCase$(file(i)) Then
        If LCase$(Data) = LCase$(App.Path + "\textfiles\" + file(i)) Then Exit Sub
        Msg = MsgBox("Replace Note " + filMain.List(E) + "?", vbExclamation + vbYesNo + vbApplicationModal, "Import Into TextMe")
        If Msg = vbNo Then
            Exit Sub
        Else
            notestatus = 1
            Exit For
        End If
        Exit Sub
    End If
Next E
FileCopy Data, App.Path + "\textfiles\" + file(i)
If DeleteOriginal = 1 And Shift <> 1 Or DeleteOriginal = 0 And Shift = 1 Then Kill Data
If file(i) = filMain.FileName Then TextChange = False
filMain.Refresh
ClickCode
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
DragCode cmd, 0
StartFileErr:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.GetStartFile()", Err.Number, Err.Description
End Sub

Sub MouseRollover(Optional Index As Integer)
Dim i As Integer
For i = imgButton.LBound To imgButton.UBound
    If i <> Index Then
        recBorder(i).FillStyle = 1
        recBorder(i).BackStyle = 0
        recBorder(i).BorderStyle = 0
    End If
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
Dim Msg As Long
Msg = MsgBox("WARNING: The TextMe Database has been corrupted!" + vbCrLf + "Would you like to wipe the TextMe Database?" + vbCrLf + vbCrLf + "This Operation Will Result In All Your Notes Being Destroyed!" + vbCrLf + vbCrLf + "Otherwise, you may open the TextFiles folder in the TextMe Program folder to find the file causing the corruption", vbCritical + vbApplicationModal + vbYesNo, "WARNING: Database Corruption")
If Msg = vbNo Then
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
Dim Msg As Long
filMain.Refresh
txtMain.Text = ""
Msg = MsgBox("WARNING: The TextMe Database is empty!" + vbCrLf + "Would you like to restore from a previous backup?" + vbCrLf + vbCrLf + "Otherwise, click No to create a blank note or Cancel to exit TextMe", vbCritical + vbApplicationModal + vbYesNoCancel, "Restore TextMe Database")
If Msg = vbYes Then
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
ElseIf Msg = vbNo Then
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
Dim Msg As Long
Dim F As Long
On Error GoTo NoSaveCode
If frmBackup.Visible = True Then Exit Sub
If TextChange = True Then
    Msg = MsgBox("Save Changes To " + CurrentNote + "?", vbExclamation + vbYesNo + vbApplicationModal, "Save Note")
    If Msg = vbYes Then
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
Dim Msg As Integer
Dim C As Integer
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
'Font Settings
txtMain.fontname = GetSetting("Peter Chapman", "TextMe", "FontName", txtMain.fontname)
txtMain.FontBold = CBool(GetSetting("Peter Chapman", "TextMe", "FontBold", CStr(txtMain.FontBold)))
txtMain.FontItalic = CBool(GetSetting("Peter Chapman", "TextMe", "FontItalic", CStr(txtMain.FontItalic)))
txtMain.FontSize = CInt(GetSetting("Peter Chapman", "TextMe", "FontSize", CStr(txtMain.FontSize)))
txtMain.FontStrikethru = CBool(GetSetting("Peter Chapman", "TextMe", "FontStrikeThru", CStr(txtMain.FontStrikethru)))
txtMain.FontUnderline = CBool(GetSetting("Peter Chapman", "TextMe", "FontUnderline", CStr(txtMain.FontUnderline)))
txtMain.ForeColor = GetSetting("Peter Chapman", "TextMe", "FontColour", CStr(txtMain.ForeColor))
txtMain.Font.Charset = CInt(GetSetting("Peter Chapman", "TextMe", "FontCharset", CStr(txtMain.Font.Charset)))
txtMain.BackColor = GetSetting("Peter Chapman", "TextMe", "FontBackground", CStr(txtMain.BackColor))
filMain.fontname = txtMain.fontname
filMain.FontSize = txtMain.FontSize
filMain.Font.Charset = txtMain.Font.Charset
filMain.ForeColor = txtMain.ForeColor
filMain.FontBold = txtMain.FontBold
filMain.FontItalic = txtMain.FontItalic
filMain.FontUnderline = txtMain.FontUnderline
filMain.FontStrikethru = txtMain.FontStrikethru
filMain.BackColor = txtMain.BackColor
For C = 0 To 15
    CustomColours(C) = CLng(GetSetting("Peter Chapman", "TextMe", "CustomColour" + CStr(C), CStr(RGB(255, 255, 255))))
Next C
Exit Sub
NoGet:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.GetRegistrySettings()", Err.Number, Err.Description
ResetRegistrySettings
End Sub
Public Sub SaveRegistrySettings()
Dim C As Integer
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
SaveSetting "Peter Chapman", "TextMe", "FontName", txtMain.fontname
SaveSetting "Peter Chapman", "TextMe", "FontBold", CStr(txtMain.FontBold)
SaveSetting "Peter Chapman", "TextMe", "FontItalic", CStr(txtMain.FontItalic)
SaveSetting "Peter Chapman", "TextMe", "FontSize", CStr(txtMain.FontSize)
SaveSetting "Peter Chapman", "TextMe", "FontStrikeThru", CStr(txtMain.FontStrikethru)
SaveSetting "Peter Chapman", "TextMe", "FontUnderline", CStr(txtMain.FontUnderline)
SaveSetting "Peter Chapman", "TextMe", "FontColour", CStr(txtMain.ForeColor)
SaveSetting "Peter Chapman", "TextMe", "FontCharset", CStr(txtMain.Font.Charset)
SaveSetting "Peter Chapman", "TextMe", "FontBackground", CStr(txtMain.BackColor)
For C = 0 To 15
    SaveSetting "Peter Chapman", "TextMe", "CustomColour" + CStr(C), CStr(CustomColours(C))
Next C
Exit Sub
NoSaveRegistry:
MsgBox "Cannot Save Registry Settings, Please Contact Your Network Administrator", vbCritical + vbDefaultButton1 + vbApplicationModal + vbOKOnly, App.Title
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.SaveRegistrySettings()", Err.Number, Err.Description
End Sub
Public Sub ResetRegistrySettings()
Dim C As Integer
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
'Font Settings
SaveSetting "Peter Chapman", "TextMe", "FontName", txtMain.fontname
SaveSetting "Peter Chapman", "TextMe", "FontBold", CStr(txtMain.FontBold)
SaveSetting "Peter Chapman", "TextMe", "FontItalic", CStr(txtMain.FontItalic)
SaveSetting "Peter Chapman", "TextMe", "FontSize", CStr(txtMain.FontSize)
SaveSetting "Peter Chapman", "TextMe", "FontStrikeThru", CStr(txtMain.FontStrikethru)
SaveSetting "Peter Chapman", "TextMe", "FontUnderline", CStr(txtMain.FontUnderline)
SaveSetting "Peter Chapman", "TextMe", "FontColour", CStr(txtMain.ForeColor)
SaveSetting "Peter Chapman", "TextMe", "FontCharset", CStr(txtMain.Font.Charset)
SaveSetting "Peter Chapman", "TextMe", "FontBackground", CStr(txtMain.BackColor)
For C = 0 To 15
    SaveSetting "Peter Chapman", "TextMe", "CustomColour" + CStr(C), CStr(RGB(255, 255, 255))
Next C
Exit Sub
NoResetRegistry:
MsgBox "Cannot Reset Registry Settings, Please Contact Your Network Administrator", vbCritical + vbDefaultButton1 + vbApplicationModal + vbOKOnly, App.Title
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmMain.ResetRegistrySettings()", Err.Number, Err.Description
End Sub
