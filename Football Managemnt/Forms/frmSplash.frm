VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4710
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7170
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4710
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   2500
      Left            =   1080
      Top             =   4440
   End
   Begin VB.PictureBox picCredits 
      BackColor       =   &H00400000&
      Enabled         =   0   'False
      Height          =   1050
      Left            =   1320
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   311
      TabIndex        =   1
      Top             =   2400
      Width           =   4725
      Begin VB.TextBox txtcredits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
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
         Text            =   "frmSplash.frx":6F4E
         Top             =   1170
         Width           =   4170
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   4440
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   8
      Top             =   4440
   End
   Begin VB.Label lblIni 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    lblIni.Caption = "Initializing System"
   
End Sub

Private Sub Timer1_Timer()

If txtcredits.Top > 0 - (txtcredits.Height) Then
    txtcredits.Top = txtcredits.Top - 1
Else
    txtcredits.Visible = False
    Timer2.Enabled = True
    Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()

    txtcredits.Top = 78
    txtcredits.Visible = True
    Timer1.Enabled = True
    Timer2.Enabled = False
    
End Sub

Private Sub Timer3_Timer()

    FrmLogin.Show
    Unload frmSplash
    Timer3.Enabled = False

End Sub
