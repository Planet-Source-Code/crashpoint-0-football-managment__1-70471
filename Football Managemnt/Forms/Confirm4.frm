VERSION 5.00
Begin VB.Form Confirm4 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6045
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   2835
         Top             =   3420
      End
      Begin VB.TextBox txtPass 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   285
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   2070
         Width           =   5565
      End
      Begin Project1.desButton cmdUnlock 
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   2760
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         Caption         =   "Unlock"
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
      Begin Project1.desButton cmdClose 
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   2760
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DENIED"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   555
         Left            =   1920
         TabIndex        =   6
         Top             =   600
         Width           =   1905
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   285
         Picture         =   "Confirm4.frx":0000
         Stretch         =   -1  'True
         Top             =   405
         Width           =   5565
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER PASSWORD TO UNLOCK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007D3F2D&
         Height          =   240
         Left            =   1305
         TabIndex        =   5
         Top             =   1725
         Width           =   3525
      End
      Begin VB.Shape Shape1 
         Height          =   915
         Left            =   225
         Top             =   1665
         Width           =   5685
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DENIED"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   555
         Left            =   2040
         TabIndex        =   4
         Top             =   600
         Width           =   1905
      End
   End
End
Attribute VB_Name = "Confirm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdUnlock_Click()
    If txtPass.Text = APass Then
        Unload Me
        DelPlayer.Show vbModal
    Else
        MsgBox "Wrong password supplied. Attempt to unlock failed.", vbOKOnly + vbExclamation, "System Manager"
        SendKeys HiLyt
        Exit Sub
    End If
End Sub

Private Sub desButton1_Click()

End Sub

Private Sub Timer1_Timer()
    If Trim(txtPass.Text) = "" Then
        cmdUnlock.Enabled = False
    Else
        cmdUnlock.Enabled = True
    End If
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdUnlock_Click
    End If
End Sub




