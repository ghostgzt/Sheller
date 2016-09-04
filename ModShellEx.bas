Attribute VB_Name = "ModShellEx"
'*************************************************************************
'**模 块 名：ModShellEx
'**说    明：增强SHELL函数
'**创 建 人：马大哈
'**描    述：紫水晶工作室 http://www.m5home.com/
'**日    期：2007年4月24日
'**版    本：V1.0
'*************************************************************************
Option Explicit

Private Declare Function GetProcessVersion Lib "kernel32" (ByVal ProcessId As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function ShellEx(ByVal FileName As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus, Optional ByVal DelayTime As Long = -1)
    '与SHELL函数一样的参数,不过是阻塞执行.(同步)
    'FileName - 目标文件名
    'WindowStyle - 程序运行时窗口的样式
    'DelayTime - 等待的时间,单位为ms
    '备注:
    '       DelayTime设置为-1时表示一直等待,直到目标程序运行结束
    Dim I As Long, J As Long
    
    I = Shell(FileName, WindowStyle)
    Do
        If GetProcessVersion(I) = 0 Then Exit Do
        Sleep 10
        J = J + 1
        If DelayTime <> -1 And J > DelayTime \ 10 Then Exit Do
    Loop
End Function

Public Function ShellOnce(ByVal FileName As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus)
    '与SHELL函数一样的参数,但只将目标执行一次
    'FileName - 目标文件名
    'WindowStyle - 程序运行时窗口的样式
    Static I As Long
    
    If I <> 0 Then          '如果有PID值就判断其是否正在执行
        If GetProcessVersion(I) <> 0 Then Exit Function       '如果正在执行,函数返回
    End If
    I = Shell(FileName, WindowStyle)
End Function

