VERSION 5.00
Begin VB.Form frmCredits 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  :.  Football Management System"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5070
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.desButton cmdclose 
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   570
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   3240
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "Credits"
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   4890
      Begin VB.PictureBox picCredits 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         Height          =   930
         Left            =   120
         ScaleHeight     =   58
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   311
         TabIndex        =   1
         Top             =   240
         Width           =   4725
         Begin VB.TextBox txtCredits 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2655
            Left            =   225
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   2
            Text            =   "frmCredits.frx":0000
            Top             =   1170
            Width           =   4170
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   120
      Picture         =   "frmCredits.frx":00C6
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "For your queries, comments and suggestions, please E-mail me at the given E-mail Address."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1845
      TabIndex        =   3
      Top             =   945
      Width           =   3165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CRASHPOINT'O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   0
      Left            =   1800
      TabIndex        =   5
      Top             =   240
      Width           =   1470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "snithgyer@yahoo.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   0
      Left            =   2880
      TabIndex        =   4
      Top             =   480
      Width           =   2115
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
If txtCredits.Top > 0 - (txtCredits.Height) Then
    txtCredits.Top = txtCredits.Top - 1
Else
    txtCredits.Visible = False
    Timer2.Enabled = True
    Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
    txtCredits.Top = 78
    txtCredits.Visible = True
    Timer1.Enabled = True
    Timer2.Enabled = False
End Sub
