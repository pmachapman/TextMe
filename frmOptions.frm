VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TextMe Settings"
   ClientHeight    =   1575
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   4455
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Reset Settings To Defaults (Click To Turn On Or Off)"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "Change Background Colour"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      ToolTipText     =   "Change TextMe's Background Colour"
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "Change Font"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Change TextMe's Font"
      Top             =   120
      Width           =   2175
   End
   Begin VB.CheckBox chkRead 
      Caption         =   "Read Only"
      Height          =   375
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      ToolTipText     =   "Select This To Make The Text In TextMe Read Only"
      Top             =   480
      Width           =   1095
   End
   Begin VB.CheckBox chkDelete 
      Caption         =   "Delete Original File"
      Height          =   375
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      ToolTipText     =   $"frmOptions.frx":000C
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      ToolTipText     =   "Apply New Settings"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      ToolTipText     =   "Close This Dialog Without Applying New Settings"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   90
      TabIndex        =   4
      ToolTipText     =   "Close This Dialog And Apply New Settings"
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FontSettings(0 To 8) As Variant

Private Sub chkDelete_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkRead_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkReset_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()
If chkReset.Value = 1 Then
    chkDelete.Value = 0
    chkRead.Value = 0
    FontSettings(0) = "MS Sans Serif"
    FontSettings(1) = 8
    FontSettings(2) = 0
    FontSettings(3) = &H80000008
    FontSettings(4) = False
    FontSettings(5) = False
    FontSettings(6) = False
    FontSettings(7) = False
    FontSettings(8) = &H80000005
    
End If
    SaveSetting "Peter Chapman", "TextMe", "DeleteOriginal", chkDelete.Value
    frmMain.txtMain.Locked = chkRead.Value
    SaveSetting "Peter Chapman", "TextMe", "ReadOnly", chkRead.Value
    frmMain.txtMain.fontname = FontSettings(0)
    frmMain.txtMain.FontSize = FontSettings(1)
    frmMain.txtMain.Font.Charset = FontSettings(2)
    frmMain.txtMain.ForeColor = FontSettings(3)
    If FontSettings(4) >= 700 Then frmMain.txtMain.FontBold = True Else frmMain.txtMain.FontBold = False
    frmMain.txtMain.FontItalic = FontSettings(5)
    frmMain.txtMain.FontUnderline = FontSettings(6)
    frmMain.txtMain.FontStrikethru = FontSettings(7)
    frmMain.txtMain.BackColor = FontSettings(8)
    
    frmMain.filMain.fontname = FontSettings(0)
    frmMain.filMain.FontSize = FontSettings(1)
    frmMain.filMain.Font.Charset = FontSettings(2)
    frmMain.filMain.ForeColor = FontSettings(3)
    If FontSettings(4) >= 700 Then frmMain.filMain.FontBold = True Else frmMain.filMain.FontBold = False
    frmMain.filMain.FontItalic = FontSettings(5)
    frmMain.filMain.FontUnderline = FontSettings(6)
    frmMain.filMain.FontStrikethru = FontSettings(7)
    frmMain.filMain.BackColor = FontSettings(8)
cmdApply.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdColour_Click()
    cmdApply.Enabled = True
    ChangeColour
End Sub

Private Sub cmdFont_Click()
    cmdApply.Enabled = True
    ChangeFont
End Sub

Private Sub cmdOK_Click()
    cmdApply.Enabled = True
    cmdApply_Click
    Unload Me
End Sub

Private Sub Form_Load()
    'center the form
    'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    'Get Settings
    chkDelete.Value = GetSetting("Peter Chapman", "TextMe", "DeleteOriginal", "0")
    chkRead.Value = GetSetting("Peter Chapman", "TextMe", "ReadOnly", "0")
    FontSettings(0) = frmMain.txtMain.fontname
    FontSettings(1) = frmMain.txtMain.FontSize
    FontSettings(2) = frmMain.txtMain.Font.Charset
    FontSettings(3) = frmMain.txtMain.ForeColor
    FontSettings(4) = frmMain.txtMain.FontBold
    FontSettings(5) = frmMain.txtMain.FontItalic
    FontSettings(6) = frmMain.txtMain.FontUnderline
    FontSettings(7) = frmMain.txtMain.FontStrikethru
    FontSettings(8) = frmMain.txtMain.BackColor
End Sub

Public Sub ChangeFont()
' Display a Choose Font dialog box.  Print out the typeface name, point size,
' and style of the selected font.  More detail about topics in this example can be found in
' the pages for CHOOSEFONT_TYPE and LOGFONT.
Dim cf As ChooseFont ' data structure needed for function
Dim lfont As LOGFONT  ' receives information about the chosen font
Dim hMem As Long, pMem As Long  ' handle and pointer to memory buffer
Dim fontname As String  ' receives name of font selected
Dim retval As Long  ' return value
On Error GoTo FontError
' Initialize frmMain.txtMain's font
' (Note that some of that information is in the CHOOSEFONT_TYPE structure instead.)
lfont.lfHeight = -MulDiv(FontSettings(1), GetDeviceCaps(hdc, LOGPIXELSY), 72)  ' determine default height
lfont.lfWidth = 0  ' determine default width
lfont.lfEscapement = 0  ' angle between baseline and escapement vector
lfont.lfOrientation = 0  ' angle between baseline and orientation vector
If FontSettings(4) = False Then
    lfont.lfWeight = 400 'normal
Else
    lfont.lfWeight = 700 'bold
End If
lfont.lfItalic = FontSettings(5)
lfont.lfUnderline = FontSettings(6)
lfont.lfStrikeOut = FontSettings(7)
lfont.lfCharSet = FontSettings(2)
lfont.lfOutPrecision = 0  ' default precision mapping
lfont.lfClipPrecision = 0  ' default clipping precision
lfont.lfQuality = 0  ' default quality setting
lfont.lfPitchAndFamily = 0 Or 16  ' default pitch, proportional with serifs
lfont.lfFaceName = FontSettings(0) & vbNullChar  ' string must be null-terminated

' Create the memory block which will act as the LOGFONT structure buffer.
hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lfont))
pMem = GlobalLock(hMem)  ' lock and get pointer
CopyMemory ByVal pMem, lfont, Len(lfont)  ' copy structure's contents into block

' Initialize dialog box: Screen and printer fonts, point size between 10 and 72.
cf.lStructSize = Len(cf)  ' size of structure
cf.hwndOwner = Me.hwnd  ' window Form1 is opening this dialog box
'below only needed if displaying printer fonts
cf.hdc = Printer.hdc  ' device context of default printer (using VB's mechanism)
NoPrinter:
cf.lpLogFont = pMem  ' pointer to LOGFONT memory block buffer
cf.iPointSize = 120  ' 12 point font (in units of 1/10 point)
cf.flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
cf.rgbColors = FontSettings(3)
cf.lCustData = 0  ' we don't use this here...
cf.lpfnHook = 0  ' ...or this...
cf.lpTemplateName = ""  ' ...or this...
cf.hInstance = 0  ' ...or this...
cf.lpszStyle = ""  ' ...or this
cf.nFontType = REGULAR_FONTTYPE  ' regular font type i.e. not bold or anything
cf.nSizeMin = 8  ' minimum point size
cf.nSizeMax = 72  ' maximum point size
' Now, call the function.  If successful, copy the LOGFONT structure back into the structure
' and then print out the attributes we mentioned earlier that the user selected.
retval = ChooseFont(cf)  ' open the dialog box
If retval <> 0 Then  ' success
  CopyMemory lfont, ByVal pMem, Len(lfont)  ' copy memory back
  ' Now make the fixed-length string holding the font name into a "normal" string.
  fontname = Left(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
  ' Display font name and a few attributes.
  FontSettings(0) = fontname
  FontSettings(1) = cf.iPointSize / 10   ' in units of 1/10 point!
  FontSettings(2) = lfont.lfCharSet
  FontSettings(3) = cf.rgbColors
  If lfont.lfWeight >= 700 Then FontSettings(4) = True Else FontSettings(4) = False
  FontSettings(5) = lfont.lfItalic
  FontSettings(6) = lfont.lfUnderline
  FontSettings(7) = lfont.lfStrikeOut
End If

' Deallocate the memory block we created earlier.  Note that this must
' be done whether the function succeeded or not.
retval = GlobalUnlock(hMem)  ' destroy pointer, unlock block
retval = GlobalFree(hMem)  ' free the allocated memory
Exit Sub
FontError:
If Err.Number = 482 Then Resume NoPrinter
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmOptions.ChangeFont()", Err.Number, Err.Description
MsgBox "Error Changing Font", vbExclamation + vbOKOnly + vbApplicationModal, "Change Font"
End Sub
Public Sub ChangeColour()
' Display a Choose Color common dialog box.  The background
' color of Form1 will be set to the color the user selects.  Although the entire
' list of custom colors is initialized to black, this example stores the
' colors into an array which can be used again to save the user's custom
' color selections.
On Error GoTo ColourError
Dim cc As CHOOSECOLOR ' structure to pass data
Dim hMem As Long  ' handle to the memory block to store the custom color list
Dim pMem As Long  ' pointer to the memory block to store the custom color list
Dim retval As Long  ' return value

' Create a memory block and get a pointer to it.
hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, 64)  ' allocate sufficient memory block
pMem = GlobalLock(hMem)  ' get a pointer to the block
' Copy the data inside the array into the memory block.
CopyMemory ByVal pMem, CustomColours(0), 64  ' 16 elements * 4 bytes

' Store the initial settings of the Choose Color box.
cc.lStructSize = Len(cc)  ' size of the structure
cc.hwndOwner = Me.hwnd
cc.hInstance = 0  ' not needed
cc.rgbResult = FontSettings(8)  ' set default selected color to Form1's background color
cc.lpCustColors = pMem  ' pointer to list of custom colors
cc.flags = CC_ANYCOLOR Or CC_RGBINIT  ' allow any color, use rgbResult as default selection
cc.lCustData = 0  ' not needed
cc.lpfnHook = 0  ' not needed
cc.lpTemplateName = ""  ' not needed

' Open the Choose Color box.  If the user chooses a color, set Form1's
' background color to that color.
retval = CHOOSECOLOR(cc)
If retval <> 0 Then  ' success
  ' Copy the possibly altered contents of the custom color list
  ' back into the array.
  CopyMemory CustomColours(0), ByVal pMem, 64
  ' Set background color.
  FontSettings(8) = cc.rgbResult
End If

' Deallocate the memory blocks to free up resources.
retval = GlobalUnlock(hMem)
retval = GlobalFree(pMem)
Exit Sub
ColourError:
If SalamanderExists = True Then Salamander.ReportError "TextMe", "frmOptions.ChangeColour()", Err.Number, Err.Description
MsgBox "Error Changing Colour", vbExclamation + vbOKOnly + vbApplicationModal, "Change Colour"
End Sub
