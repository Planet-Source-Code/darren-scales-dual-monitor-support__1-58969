VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About the screen class"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   3945
      TabIndex        =   4
      Top             =   3660
      Width           =   3975
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   1
         Left            =   2520
         ScaleHeight     =   435
         ScaleWidth      =   255
         TabIndex        =   9
         Top             =   -60
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   0
         Left            =   1140
         ScaleHeight     =   435
         ScaleWidth      =   255
         TabIndex        =   8
         Top             =   -60
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "To Do"
         Height          =   435
         Left            =   2640
         TabIndex        =   7
         Top             =   -120
         Width           =   1395
      End
      Begin VB.CommandButton Command2 
         Caption         =   "About"
         Height          =   435
         Left            =   1260
         TabIndex        =   6
         Top             =   -120
         Width           =   1395
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Class Usage"
         Height          =   435
         Left            =   -120
         TabIndex        =   5
         Top             =   -120
         Width           =   1395
      End
   End
   Begin VB.PictureBox picToDo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3075
      Left            =   60
      ScaleHeight     =   3075
      ScaleWidth      =   4095
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Label Label10 
         Caption         =   "- Anything else that would look good"
         Height          =   255
         Left            =   60
         TabIndex        =   17
         Top             =   2040
         Width           =   3915
      End
      Begin VB.Label Label9 
         Caption         =   "- Rearrange monitors"
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   1680
         Width           =   3915
      End
      Begin VB.Label Label8 
         Caption         =   "- View monitors as layout in XP display properties"
         Height          =   255
         Left            =   60
         TabIndex        =   15
         Top             =   1440
         Width           =   3915
      End
      Begin VB.Label Label7 
         Caption         =   "- Retrieve colour information, ie. colour depth"
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   1200
         Width           =   3915
      End
      Begin VB.Label Label6 
         Caption         =   "- Change resolution on a particular screen"
         Height          =   255
         Left            =   60
         TabIndex        =   13
         Top             =   960
         Width           =   3915
      End
      Begin VB.Label Label5 
         Caption         =   "These are the things that Im hoping to add if i eventually get some time;"
         Height          =   435
         Left            =   60
         TabIndex        =   12
         Top             =   240
         Width           =   3915
      End
   End
   Begin VB.PictureBox picUsage 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3075
      Left            =   60
      ScaleHeight     =   3075
      ScaleWidth      =   4095
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "See the accompanied usage txt file"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   1380
         Width           =   3255
      End
   End
   Begin VB.Label Label4 
      Caption         =   $"frmAbout.frx":0000
      Height          =   1035
      Left            =   60
      TabIndex        =   3
      Top             =   2400
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   $"frmAbout.frx":0109
      Height          =   675
      Left            =   60
      TabIndex        =   2
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Screen Class"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "This simple app demonstrates the use of the clsScreen class, making it easy to add dual monitor support to your own apps."
      Height          =   675
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   4095
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    picUsage.Visible = True
    picToDo.Visible = False

End Sub

Private Sub Command2_Click()

    picUsage.Visible = False
    picToDo.Visible = False

End Sub

Private Sub Command3_Click()

    picUsage.Visible = False
    picToDo.Visible = True

End Sub
