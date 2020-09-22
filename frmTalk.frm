VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTray 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6300
      ScaleHeight     =   285
      ScaleWidth      =   2745
      TabIndex        =   1
      Top             =   0
      Width           =   2775
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   1
         Left            =   1800
         ScaleHeight     =   795
         ScaleWidth      =   135
         TabIndex        =   5
         Top             =   -300
         Width           =   135
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   0
         Left            =   840
         ScaleHeight     =   795
         ScaleWidth      =   135
         TabIndex        =   4
         Top             =   -240
         Width           =   135
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "About"
         Height          =   495
         Left            =   -60
         TabIndex        =   2
         Top             =   -120
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Caption         =   "Settings"
         Height          =   495
         Left            =   900
         TabIndex        =   3
         Top             =   -120
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Caption         =   "Exit"
         Height          =   495
         Left            =   1860
         TabIndex        =   6
         Top             =   -120
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    frmAbout.Show

End Sub

Private Sub Command2_Click()

    frmSettings.Show

End Sub

Private Sub Command3_Click()

    Set Mons = Nothing
    End

End Sub

Private Sub Form_Load()

    SetDisplay
    Text1.Text = Mons.Count(True) & " monitors, "
    Text1 = Text1 & Mons.Width(Mons.PrimaryMon) & "x" & Mons.Height(Mons.PrimaryMon)
    Text1 = Text1 & " being the resolution of your primary monitor(" & Mons.PrimaryMon & ")"
    
    Text1 = Text1 & "  Any problems? Found an error? : darren@sirdaz.com"
    
    frmSettings.Show
    
End Sub

Public Function SetDisplay()

    Me.Width = Mons.WorkWidth(D_Mon) * Screen.TwipsPerPixelX
    Me.Left = Mons.WorkLeft(D_Mon) * Screen.TwipsPerPixelX
    
    Me.Height = Text1.Height
    
'move bar to the top or bottom
    If Set_BarAlign = 0 Then
        'mons.worktop return pixels, we need twips so it works it out
        Me.Top = Mons.WorkTop(D_Mon) * Screen.TwipsPerPixelY
    ElseIf Set_BarAlign = 1 Then
        Me.Top = (Mons.WorkHeight(D_Mon) * Screen.TwipsPerPixelY) - Me.Height
    End If
    
    Text1.Width = Me.Width - picTray.Width
    Text1.Left = 0
    Text1.Top = 0
    
    picTray.Left = Text1.Width
    picTray.Top = 0

End Function
