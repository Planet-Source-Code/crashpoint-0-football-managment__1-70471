VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form AdminDel 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  :.  Delete Admin Profile"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   1440
      Top             =   5040
   End
   Begin MSAdodcLib.Adodc adoAdDel 
      Height          =   330
      Left            =   120
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Admin Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000009&
      Height          =   1935
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   3855
      Begin VB.PictureBox picAns 
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1515
         ScaleWidth      =   3555
         TabIndex        =   7
         Top             =   240
         Width           =   3615
         Begin VB.Label lblErr 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   615
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.Label lblInfo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "FOUND !!!"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   480
            TabIndex        =   11
            Top             =   0
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000FFFF&
            Visible         =   0   'False
            X1              =   3720
            X2              =   -120
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Label lbladCon 
            BackStyle       =   0  'Transparent
            Caption         =   "------"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   1200
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.Label lbladDes 
            BackStyle       =   0  'Transparent
            Caption         =   "------"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   840
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.Label lbladname 
            BackStyle       =   0  'Transparent
            Caption         =   "------"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   480
            Visible         =   0   'False
            Width           =   3015
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000009&
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   4200
      Width           =   3855
      Begin Project1.desButton cmdClose 
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
      Begin Project1.desButton cmdDelete 
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Delete"
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Caption         =   "Coresponding Password"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   3480
      Width           =   3855
      Begin VB.TextBox txtPass 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   360
         Locked          =   -1  'True
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   330
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Enter Admin Id"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3855
      Begin Project1.desButton cmdSearch 
         Height          =   375
         Left            =   2880
         TabIndex        =   13
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "Search"
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
      Begin VB.TextBox txtadminId 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "Admin Id :-"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Picture         =   "AdminDel.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "AdminDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdDelete_Click()

    On Error Resume Next
        
    AUsername = adoAdDel.Recordset.Fields("Username")
    APass = adoAdDel.Recordset.Fields("Password")

    If txtPass.Text = tempPass Then
        
        If adoAdDel.Recordset.RecordCount = 0 Then
            Me.Hide
            MsgBox "There Are No Admin Profiles Available. " & vbCrLf & _
                    " The System Will Be Logout." & vbCrLf + vbCrLf & _
                    " Please Create A New Admin Profile To Access The System", _
                    vbOKOnly + vbInformation, "SysMan"
            'MDIMain.Hide
            AdminNew.Show vbModal
            AdminLogin.Show vbModal
            Unload Me
            Exit Sub
        End If
        
        If UCase(txtAdminId.Text) = UCase(tempUser) Then
            Me.Hide
            MsgBox "The Admin Profile To Be Deleted In Currently Being Used." & _
                    vbCrLf + vbCrLf & "Automatic Logout Suggetsed. " & vbCrLf & _
                    " Please Log-In With A Different Admin Profile.", _
                    vbOKOnly + vbInformation, "SysMan"
            Unload MDIMain
            Unload Me
            'Unload MainScr
            AdminLogin.Show
            Exit Sub
        Else
            If MsgBox("The Admin Profile Will Be Deleted !!!" & vbCrLf & _
                    " Proceed?", vbYesNo + vbQuestion, "SysMan") = vbYes Then
                adoAdDel.Recordset.delete
                adoAdDel.Refresh
                txtPass.Text = ""
                txtPass.Enabled = False
                txtPass.BackColor = &H808080
                txtAdminId.Text = ""
                cmdDelete.Enabled = False
                Unload Me
            Else
                Exit Sub
            End If
            Exit Sub
        End If
    
    Else
        MsgBox "The Password Doesn't Match With The Current Log-In Admin Profile. " & vbCrLf & _
               " Access denied.", vbCritical + vbOKOnly, "SysMan"
        txtPass.SetFocus
        SendKeys highLig
        
        Exit Sub
    End If
End Sub

Private Sub cmdSearch_Click()

    On Error GoTo noData
    
        adoAdDel.Refresh
        adoAdDel.Recordset.Find ("Username = '" & UCase(txtAdminId.Text) & "'")
        
        AName = adoAdDel.Recordset.Fields("Name")
        AId = adoAdDel.Recordset.Fields("ID_No")
        ATel = adoAdDel.Recordset.Fields("Telphone")
        
        If lbladname.Caption <> "" Or lbladDes.Caption <> "" _
                    Or lbladCon.Caption <> "" Or lblInfo.Caption <> "" Then
                    
            lbladname.Caption = ""
            lbladDes.Caption = ""
            lbladCon.Caption = ""
            lblInfo.Caption = ""
            
        End If
        
        lblInfo.Caption = "FOUND !!!"
        lblInfo.Visible = True
        Line1.Visible = True
        lbladname.Visible = True
        lbladDes.Visible = True
        lbladCon.Visible = True
        
        lbladname.Caption = AName
        lbladDes.Caption = AId
        lbladCon.Caption = ATel
        
        
        cmdDelete.Enabled = True
        txtPass.BackColor = &HFFFFFF
        txtPass.Enabled = True
        txtPass.Locked = False
        txtPass.SetFocus
        Exit Sub
    
noData:
        'MsgBox "The Admin Profile No Exist !!! " & vbCrLf & _
            " Please enter a valid Admin Id.", vbOKOnly, "SysMan"
        lblInfo.Caption = "NOT FOUND !!!"
        lblErr.Visible = True
        lblErr.Caption = "The Admin Profile No Exist !!!" & _
                            vbCrLf & " Please Enter A Valid Admin Id... "
        Timer1.Enabled = True
        lbladname.Caption = ""
        lbladDes.Caption = ""
        lbladCon.Caption = ""
        lblInfo.Caption = ""
        Line1.Visible = False
        
        txtPass.BackColor = &HC0C0C0
        cmdDelete.Enabled = False
        txtAdminId.SetFocus
        SendKeys highLig
        adoAdDel.Refresh

End Sub

Private Sub Form_Load()

    Call DataConn(adoAdDel, "Admin")
    
End Sub

Private Sub Timer1_Timer()

    If lblErr.Visible = True Then
        lblErr.Caption = ""
    End If
    
End Sub
