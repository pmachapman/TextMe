Attribute VB_Name = "modMain"
Option Explicit
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Colour dialog
Public Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Public Type CHOOSECOLOR
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type
Public Const CC_ANYCOLOR = &H100
Public Const CC_RGBINIT = &H1
'Font Dialog
Public Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Const LOGPIXELSY = 90
Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 32 'lfFaceName(1 To 32) As Byte
End Type
Public Type ChooseFont
        lStructSize As Long
        hwndOwner As Long          '  caller's window handle
        hdc As Long                '  printer DC/IC or NULL
        lpLogFont As Long
        iPointSize As Long         '  10 * size in points of selected font
        flags As Long              '  enum. type flags
        rgbColors As Long          '  returned text color
        lCustData As Long          '  data passed to hook fn.
        lpfnHook As Long           '  ptr. to hook function
        lpTemplateName As String     '  custom template name
        hInstance As Long          '  instance handle of.EXE that
                                       '    contains cust. dlg. template
        lpszStyle As String          '  return the style field here
                                       '  must be LF_FACESIZE or bigger
        nFontType As Integer          '  same value reported to the EnumFonts
                                       '    call back with the extra FONTTYPE_
                                       '    bits added
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long           '  minimum pt size allowed &
        nSizeMax As Long           '  max pt size allowed if
                                       '    CF_LIMITSIZE is used
End Type
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_ZEROINIT = &H40
Public Const CF_SCREENFONTS = &H1
Public Const CF_PRINTERFONTS = &H2
Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Public Const CF_EFFECTS = &H100&
Public Const CF_FORCEFONTEXIST = &H10000
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Public Const CF_LIMITSIZE = &H2000&
Public Const REGULAR_FONTTYPE = &H400
'Options
Global DeleteOriginal As Boolean
Global ReadOnly As Boolean
'Salamander
Global Salamander As Object
Global SalamanderExists As Boolean
'Custom colours
Global CustomColours(0 To 15) As Long  ' holds list of the 16 custom colors

Sub Main()
If Trim$(UCase$(Command$)) = ":UPDATE" Then
    frmUpdate.Show
ElseIf Trim$(UCase$(Command$)) = ":ABOUT" Then
    frmAbout.Show vbModal
Else
    If GetSystemMetrics(19) Then 'Mouse Exists
        CreateSalamander 'Start Error Reporting Module If Avaliable
        frmMain.Show
    Else
        MsgBox "TextMe Requires A Mouse To Run." + vbCrLf + "Please Download An Older Version From http://pcos.cjb.net/ For Use Without A Mouse.", vbApplicationModal + vbCritical + vbOKOnly, "Cannot Run TextMe"
        If SalamanderExists = True Then Salamander.ReportError "TextMe", "modMain.Main()", 0, "No Mouse"
        End
    End If
End If
End Sub
Public Sub CreateSalamander()
On Error GoTo NoCreateSalamander
Set Salamander = CreateObject("Salamander.ErrorReporting")
SalamanderExists = True
Exit Sub
NoCreateSalamander:
SalamanderExists = False
End Sub

Function AllFiles(ByVal DirPath As String) As String()
'EXAMPLE
'Dim sFiles() As String
'Dim lCtr As Long

'sFiles = AllFiles("C:\windows\")
'For lCtr = 0 To UBound(sFiles)
'    Debug.Print sFiles(lCtr)
'Next

Dim sFile As String
Dim lElement As Long
Dim sAns() As String
ReDim sAns(0) As String

sFile = Dir(DirPath, vbNormal + vbHidden + vbReadOnly + vbSystem + vbArchive)
If sFile <> "" Then
sAns(0) = sFile
    Do
        sFile = Dir
        If sFile = "" Then Exit Do
        lElement = IIf(sAns(0) = "", 0, UBound(sAns) + 1)
        ReDim Preserve sAns(lElement) As String
        sAns(lElement) = sFile
    Loop
End If
AllFiles = sAns
End Function
