Attribute VB_Name = "modMain"
Option Explicit

Public Mons As clsScreens

Public D_Mon As Integer
Public Set_BarAlign As Integer

Sub Main()

    LoadSettings
    Set Mons = New clsScreens
    
    frmMain.Show

End Sub

Public Function LoadSettings()

    D_Mon = 1
    Set_BarAlign = 1
    
End Function
