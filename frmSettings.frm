VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   180
      TabIndex        =   8
      Top             =   2040
      Width           =   4635
      Begin VB.Label Label1 
         Caption         =   "Clicking around each monitor above(the boxes) and clicking 'Apply' will position the bar to that location on each monitor"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   180
         Width           =   4395
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4980
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.CheckBox chkShowDis 
      Caption         =   "Show disabled"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1395
   End
   Begin VB.PictureBox picMonitors 
      BackColor       =   &H8000000B&
      Height          =   1575
      Left            =   60
      ScaleHeight     =   1515
      ScaleWidth      =   6015
      TabIndex        =   1
      Top             =   60
      Width           =   6075
      Begin VB.PictureBox picBar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         ScaleHeight     =   165
         ScaleWidth      =   1245
         TabIndex        =   3
         Top             =   1740
         Width           =   1275
      End
      Begin VB.PictureBox Monitor 
         BackColor       =   &H00E0E0E0&
         Height          =   1095
         Index           =   0
         Left            =   2820
         ScaleHeight     =   1035
         ScaleWidth      =   1215
         TabIndex        =   2
         Top             =   1740
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblNotEnabled 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Not Enabled"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1320
         TabIndex        =   4
         Top             =   1740
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4980
      TabIndex        =   0
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblDetail 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   1680
      Width           =   2775
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WhichMon As Integer
Private BarAlign As Integer

Private Sub chkShowDis_Click()

    SetMonitorDisp

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

'set some settings
    D_Mon = WhichMon          'the monitor to be shown on
    Set_BarAlign = BarAlign   'top or bottom of screen
    
    frmMain.SetDisplay        'position the bar
    
End Sub

Private Function SetMonitorDisp()
Dim I As Integer
Dim TotWidth As Long
Dim MonCent As Long
Dim TakeI As Integer

    ClearMons
'detects No. screens, creates boxes to resemble them.

    For I = 1 To Mons.Count   'loop through all available monitors
        If chkShowDis = False And Mons.Enabled(I) = False Then  'is the monitor disabled? do we want disabled?
        
            'monitor not enabled and user doesnt want to see them
            TakeI = TakeI + 1
        Else
            Load Monitor(I - TakeI)
            Monitor(I - TakeI).Tag = I  'uses the tag to remember which monitor it resembles
            If Mons.Enabled(I) = False Then
                'the monitors disabled, sp put a label on the box to let us know
                Monitor(I - TakeI).Tag = Monitor(I - TakeI).Tag & "'dis'"  'make a note its disabled in the tag
                Load lblNotEnabled(lblNotEnabled.ubound + 1)
                With lblNotEnabled(lblNotEnabled.ubound)
                    Set .Container = Monitor(I - TakeI)
                    .Top = (Monitor(I - TakeI).Height / 2) - (.Height / 2)
                    .Left = (Monitor(I - TakeI).Width / 2) - (.Width / 2)
                    .Visible = True
                End With
            End If
                'position the box and make it look 'pretty'
            Monitor(I - TakeI).BackColor = &HE0E0E0
            Monitor(I - TakeI).Left = Monitor(I - 1 - TakeI).Left + Monitor(I - TakeI).Width + 50
            Monitor(I - TakeI).Top = (Monitor(I - TakeI).Container.Height / 2) - (Monitor(I - TakeI).Height / 2)
        End If
    Next I
    
        'positions the boxes in the center
    TotWidth = Monitor(Monitor.ubound).Width + Monitor(Monitor.ubound).Left
    TotWidth = TotWidth - Monitor(1).Left
    MonCent = TotWidth / (Monitor.ubound)
        'still positioning.....
    For I = 1 To Monitor.ubound
        Monitor(I).Left = (MonCent * I) - (Monitor(I).Width / 2)
        Monitor(I).Left = Monitor(I).Left + ((Monitor(I).Container.Width - TotWidth) / ((Monitor.ubound) * 2))
        Monitor(I).Visible = True
    Next I

        'creates a little bar in the correct place
    If Set_BarAlign = 0 Then
        Monitor_MouseUp D_Mon, 1, 0, 1, 1
    ElseIf Set_BarAlign = 1 Then
        Monitor_MouseUp D_Mon, 1, 0, 1, Monitor(D_Mon).Height / 2 + 1
    End If
    
End Function

Private Function ClearMons()
Dim I As Integer

    For I = lblNotEnabled.ubound To 1 Step -1
        Unload lblNotEnabled(I)
    Next I
    
    picBar.Visible = False
    Set picBar.Container = picMonitors
    
    For I = Monitor.ubound To 1 Step -1
        Unload Monitor(I)
    Next I
    
End Function

Private Sub Form_Load()

    SetMonitorDisp

End Sub

Private Sub Monitor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If InStr(1, Monitor(Index).Tag, "'dis'") > 0 Then Exit Sub  'is the monitor disabled?
       
    Set picBar.Container = Monitor(Index)  'puts the little bar into the selected box
    
        'is the bar to be at the top or bottom?
    If Y >= Monitor(Index).Height / 2 Then
        picBar.Top = Monitor(Index).Height - picBar.Height
        BarAlign = 1
    Else
        picBar.Top = -10
        BarAlign = 0
    End If
    
    picBar.Width = Monitor(Index).Width
    picBar.Left = -10
    picBar.Visible = True
    
    WhichMon = Monitor(Index).Tag
        'retrieve resolution for the selected screen to show us
    lblDetail.Caption = "Monitor " & WhichMon & ". " & Mons.Width(WhichMon) & "x" & Mons.Height(WhichMon)
    
End Sub
