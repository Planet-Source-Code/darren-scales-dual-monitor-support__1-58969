VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function EnumDisplayMonitors Lib "user32" (ByVal hdc As Long, ByVal lprcClip As Long, ByVal lpfnEnum As Long, dwData As Any) As Long
Private Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Const SM_XVIRTUALSCREEN = 76      'Virtual Left
Const SM_YVIRTUALSCREEN = 77      'Virtual Top
Const SM_CXVIRTUALSCREEN = 78     'Virtual Width
Const SM_CYVIRTUALSCREEN = 79     'Virtual Height
Const SM_CMONITORS = 80           'Get number of monitors
Const SM_SAMEDISPLAYFORMAT = 81

Property Get Count(Optional EnabledOnly As Boolean) As Integer
Dim I As Integer

    Count = GetSystemMetrics(SM_CMONITORS) + 1

    If EnabledOnly = False Then Exit Property
    
    For I = 1 To Count
        If Width(I) = 0 And Height(I) = 0 Then
            Count = Count - 1
        End If
    Next I

End Property

Property Get AllSameColourDepth() As Boolean

    AllSameColourDepth = GetSystemMetrics(SM_SAMEDISPLAYFORMAT)

End Property

Property Get Enabled(Screen As Integer) As Boolean

    If Width(Screen) = 0 And Height(Screen) = 0 Then
        Enabled = False
    Else
        Enabled = True
    End If

End Property

Property Get PrimaryMon() As Integer
Dim I As Integer

    For I = 1 To Count
        If Left(I) = 0 And Top(I) = 0 Then
            PrimaryMon = I
            Exit For
        End If
    Next I

End Property

Public Function Height(Screen As Integer) As Long
Dim MonInf As MONITORINFO

    MonInf.cbSize = Len(MonInf)
    GetMonitorInfo Screen, MonInf
    Height = MonInf.rcMonitor.Bottom - MonInf.rcMonitor.Top

End Function
Public Function Width(Screen As Integer) As Long
Dim MonInf As MONITORINFO

    MonInf.cbSize = Len(MonInf)
    GetMonitorInfo Screen, MonInf
    Width = MonInf.rcMonitor.Right - MonInf.rcMonitor.Left

End Function
Public Function Top(Screen As Integer) As Long
Dim MonInf As MONITORINFO

    MonInf.cbSize = Len(MonInf)
    GetMonitorInfo Screen, MonInf
    Top = MonInf.rcMonitor.Top

End Function

Public Function Left(Screen As Integer) As Long
Dim MonInf As MONITORINFO

    MonInf.cbSize = Len(MonInf)
    GetMonitorInfo Screen, MonInf
    Left = MonInf.rcMonitor.Left

End Function

Public Function WorkHeight(Screen As Integer) As Long
Dim MonInf As MONITORINFO

    MonInf.cbSize = Len(MonInf)
    GetMonitorInfo Screen, MonInf
    WorkHeight = MonInf.rcWork.Bottom '- MonInf.rcWork.Top

End Function
Public Function WorkWidth(Screen As Integer) As Long
Dim MonInf As MONITORINFO

    MonInf.cbSize = Len(MonInf)
    GetMonitorInfo Screen, MonInf
    WorkWidth = MonInf.rcWork.Right - MonInf.rcWork.Left

End Function
Public Function WorkTop(Screen As Integer) As Long
Dim MonInf As MONITORINFO

    MonInf.cbSize = Len(MonInf)
    GetMonitorInfo Screen, MonInf
    WorkTop = MonInf.rcWork.Top

End Function

Public Function WorkLeft(Screen As Integer) As Long
Dim MonInf As MONITORINFO

    MonInf.cbSize = Len(MonInf)
    GetMonitorInfo Screen, MonInf
    WorkLeft = MonInf.rcWork.Left

End Function

