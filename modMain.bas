Attribute VB_Name = "modMain"
Option Explicit
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Options
Global DeleteOriginal As Boolean
Global ReadOnly As Boolean
'Salamander
Global Salamander As Object
Global SalamanderExists As Boolean
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
