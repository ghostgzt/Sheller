VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "Sheller"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
  
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
  
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
  
  
  
Const PROCESS_QUERY_INFORMATION = &H400
  
Const STILL_ALIVE = &H103


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim ff
Function xshell(sty, sec, wat, shl)
    Dim pid, hProcess
    Sleep (sec)
    pid = Shell(shl, sty)
   If wat = "w" Then
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)
    Do
    Call GetExitCodeProcess(hProcess, ExitCode)
    DoEvents
    Loop While ExitCode = STILL_ALIVE
    Call CloseHandle(hProcess)
   End If
End Function
Function dshell(str As String)
If Mid(str, 1, 1) <> "#" Then
Dim zx
zx = Split(str, "|")
Select Case UBound(zx)
Case 0
xshell 1, 0, "n", zx(0)
Case 1
xshell zx(0), 0, "n", zx(1)
Case 2
xshell zx(0), zx(1), "n", zx(2)
Case 3
xshell zx(0), zx(1), zx(2), zx(3)
End Select
End If
End Function
Private Sub Form_Load()
On Error Resume Next
ChDrive (Mid(App.Path, 1, 2))
ChDir (App.Path)
On Error GoTo x
If Command$ <> "" Then
    ff = Replace(Command$, """", "")
    If ff = "/?" Then
    MsgBox App.EXEName + ".exe " + App.EXEName + ".ini" + vbCrLf + " " + App.EXEName + ".ini" + vbCrLf + "  1|3000|w|notpad" + vbCrLf + "  cmd" + vbCrLf + "  2|calc" + vbCrLf + "  0|1000|mspaint", vbInformation
    End
    End If
Else
    ff = App.EXEName + ".ini"
End If

If Mid(ff, 2, 1) <> ":" Then
ff = App.Path & "\" & ff
If Dir(ff) = "" Then
GoTo x

End If


End If

On Error GoTo y
Dim str As String
Open ff For Input As #1
Do While Not EOF(1)
Line Input #1, str
If str <> "" Then
dshell (str)
End If
Loop
Close
End

Exit Sub
x:
MsgBox "Shell Is Not Found!", vbExclamation, "Sheller"
Open ff For Output As #1
Close
End
y:
MsgBox "Shell Is Error!", vbExclamation, "Sheller"
End
End Sub
