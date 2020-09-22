VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ClerkEdit 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  :. Edit Clerk Profile"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Search Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   22
      Top             =   720
      Width           =   6015
      Begin VB.TextBox txtAdminId 
         Height          =   375
         Left            =   1440
         TabIndex        =   24
         Top             =   360
         Width           =   2895
      End
      Begin Project1.desButton cmdSearch 
         Height          =   375
         Left            =   4560
         TabIndex        =   23
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "Clerk Id :"
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
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Caption         =   "Clerk Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   6015
      Begin VB.TextBox txtAdName 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtadDes 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtAdCont 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   16
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000009&
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   1680
         TabIndex        =   11
         Top             =   2160
         Width           =   4215
         Begin VB.TextBox txtNPass 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2040
            Locked          =   -1  'True
            MaxLength       =   6
            PasswordChar    =   "*"
            TabIndex        =   13
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtCNpass 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2040
            Locked          =   -1  'True
            MaxLength       =   6
            PasswordChar    =   "*"
            TabIndex        =   12
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000009&
            Caption         =   "New Password"
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
            TabIndex        =   15
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000009&
            Caption         =   "Confirm Password"
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
            TabIndex        =   14
            Top             =   960
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000009&
         Caption         =   "Authorization"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1680
         TabIndex        =   8
         Top             =   3720
         Width           =   4215
         Begin VB.TextBox txtOPass 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2040
            Locked          =   -1  'True
            MaxLength       =   6
            PasswordChar    =   "*"
            TabIndex        =   9
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000009&
            Caption         =   "Old Password"
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
            TabIndex        =   10
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1695
         ScaleWidth      =   1455
         TabIndex        =   6
         Top             =   960
         Width           =   1455
         Begin VB.Shape Shape1 
            BorderColor     =   &H0000FFFF&
            Height          =   1695
            Left            =   0
            Top             =   0
            Width           =   1455
         End
         Begin VB.Image imgAdmin 
            BorderStyle     =   1  'Fixed Single
            Height          =   1695
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Select Photo"
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
            Height          =   615
            Left            =   240
            TabIndex        =   7
            Top             =   480
            Width           =   975
         End
      End
      Begin Project1.desButton cmdBrowse 
         Height          =   735
         Left            =   240
         TabIndex        =   5
         Top             =   3000
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1296
         Caption         =   "Browse"
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
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "Clerk Name"
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
         Left            =   1800
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "ID Number "
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
         Left            =   1800
         TabIndex        =   20
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "Contact No #"
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
         Left            =   1800
         TabIndex        =   19
         Top             =   1560
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000009&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   6120
      Width           =   6015
      Begin Project1.desButton cmdClose 
         Height          =   375
         Left            =   3600
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
      Begin Project1.desButton cmdUpdate 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Update"
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
      Begin Project1.desButton cmdReload 
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Reload"
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
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   1440
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoEdit 
      Height          =   330
      Left            =   120
      Top             =   7200
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
      Caption         =   "Admin Edit"
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
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "EDIT CLERK PROFILE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   600
      TabIndex        =   26
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "ClerkEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdBrowse_Click()

    Dim sPic As String
    Dim ssPic As String
    
      dlgCommon.Filter = "Picture Files(*.jpg; *.bmp; *.gif)|*.jpg;*.bmp;*.gif"
      dlgCommon.ShowOpen
      
        sPic = dlgCommon.FileName
        ssPic = Mid$(sPic, InStrRev(sPic, "/") + 1)
        
            If sPic <> "" Then
                imgAdmin.Picture = LoadPicture(sPic)
            End If

End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdReload_Click()
    
    If MsgBox("This will reload the current clerk profile !!!." & vbCrLf & _
            " Any unsaved data will be lost. Do You Want To Proceed?", _
                            vbYesNo + vbQuestion, "SysMan") = vbYes Then
        
        imgAdmin.Picture = LoadPicture("")
        dlgCommon.FileName = ""
        
        adoEdit.Refresh
        
        adoEdit.Recordset.Find ("Username = '" & UCase(txtAdminId.Text) & "'")
         
         txtAdName.Text = adoEdit.Recordset.Fields("Name")
         txtadDes.Text = adoEdit.Recordset.Fields("ID_No")
         txtAdCont.Text = adoEdit.Recordset.Fields("Telphone")
        
        On Error Resume Next
            imgAdmin.Picture = LoadPicture("")
            
         txtAdName.Enabled = True
         txtadDes.Enabled = True
         txtAdCont.Enabled = True
         
         txtNPass.BackColor = &HFFFFFF
         txtCNpass.BackColor = &HFFFFFF
         txtOPass.BackColor = &HFFFFFF
         
         txtAdName.Locked = False
         txtadDes.Locked = False
         txtAdCont.Locked = False
         txtNPass.Locked = False
         txtCNpass.Locked = False
         txtOPass.Locked = False
                                
            cmdBrowse.Enabled = True
            
            cmdUpdate.Enabled = True
            cmdReload.Enabled = True
            
            txtAdName.SetFocus
            SendKeys highLig
    Else
        Exit Sub
    
    End If

End Sub

Private Sub cmdSearch_Click()
    
    Dim sPic As String
    
    On Error GoTo noRecord
    
    
        adoEdit.Refresh
        adoEdit.Recordset.Find ("Username = '" & Trim(txtAdminId.Text) & "'")
         
        txtAdName.Text = adoEdit.Recordset.Fields("Name")
        txtadDes.Text = adoEdit.Recordset.Fields("ID_No")
        txtAdCont.Text = adoEdit.Recordset.Fields("Telphone")
        
        On Error Resume Next
        sPic = adoEdit.Recordset.Fields("Picture")
        Label5.Visible = False
        imgAdmin.Picture = LoadPicture(sPic)
                
        txtAdName.BackColor = &HFFFFFF
        txtadDes.BackColor = &HFFFFFF
        txtAdCont.BackColor = &HFFFFFF
        
        txtAdName.Enabled = True
        txtadDes.Enabled = True
        txtAdCont.Enabled = True
        txtNPass.Enabled = True
        txtCNpass.Enabled = True
        txtOPass.Enabled = True
        
        cmdReload.Enabled = True
        cmdUpdate.Enabled = True
        
        txtAdName.SetFocus
        SendKeys highLig
        
        Exit Sub
noRecord:
        MsgBox "The clerk profile could not be found.", vbOKOnly + vbExclamation, "SysMan"
        
        Call clearData
        SendKeys highLig
        
End Sub

Private Sub cmdSearch_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtAdName.SetFocus
        SendKeys highLig
    End If
    
End Sub

Private Sub cmdUpdate_Click()
    
    On Error GoTo errorhandle
    
    If txtAdminId.Text = "" Then
        Call missing
        txtAdminId.SetFocus
        Exit Sub
    End If
    
    If txtAdName.Text = "" Then
        Call missing
        txtAdName.SetFocus
        Exit Sub
    End If
        
    If txtadDes.Text = "" Then
        Call missing
        txtadDes.SetFocus
        Exit Sub
    End If
    
    If txtNPass.Text = "" Then
        Call missing
        txtNPass.SetFocus
        Exit Sub
    End If
    
    If txtCNpass.Text = "" Then
        Call missing
        txtCNpass.SetFocus
        Exit Sub
    End If
    
    If txtOPass.Text = "" Then
        Call missing
        txtOPass.SetFocus
        Exit Sub
    End If
    
    'If imgAdmin.Picture = 0 Then
        'MsgBox "Please provide a photo for identification", vbOKOnly + vbExclamation, "SysMan"
        'Exit Sub
    'End If
    
    If txtCNpass.Text <> txtNPass.Text Then
        MsgBox "Password Must Be Same!!!", vbInformation, "SysMan"
        txtCNpass.Text = ""
        txtCNpass.SetFocus
        Exit Sub
    End If
    
    If txtOPass.Text <> adoEdit.Recordset.Fields("Password") Then
        MsgBox "Password No Match !!!", vbInformation, "SysMan"
        txtOPass.Text = ""
        txtOPass.SetFocus
        Exit Sub
    End If
    
    If MsgBox("The Clerk Profile Will Be Updated.. Proceed ?", vbOKCancel + vbQuestion, "SysMan") = vbOK Then
             
        adoEdit.Recordset.Fields("Name") = txtAdName.Text
        adoEdit.Recordset.Fields("ID_No") = txtadDes.Text
        adoEdit.Recordset.Fields("Telphone") = txtAdCont.Text
        adoEdit.Recordset.Fields("Password") = txtNPass.Text
        
        On Error Resume Next
        If dlgCommon.FileName <> "" Then
            adoEdit.Recordset.Fields("Picture") = dlgCommon.FileName
        End If
        
        adoEdit.Recordset.Update
        adoEdit.Refresh
            
        If MsgBox("Record successfully updated. Continue editing record?", vbYesNo + vbQuestion, "SysMan") = vbYes Then
            txtAdminId.SetFocus
            SendKeys highLig
            Exit Sub
        Else
            MsgBox "The System Will Be Log-Out Automatically" & vbCrLf & " Please Log-In Again ", vbInformation, "SysMan"
            Unload Me
            'Unload MainScr
            AdminLogin.Show
            AdminLogin.txtAdminId.Text = ""
            AdminLogin.txtAdminPass.Text = ""
        End If
           
    Else
            Exit Sub

    End If
    
    Exit Sub
    
errorhandle:
    
End Sub

Private Sub Form_Load()

    Call DataConn(adoEdit, "Clerk")
    cmdReload.Enabled = False
    cmdUpdate.Enabled = False
    
End Sub

Public Sub clearData()

    'clear the data fields
        txtAdName.Text = ""
        txtadDes.Text = ""
        txtAdCont.Text = ""
    
    'disables the buttons
    
        cmdReload.Enabled = False
        cmdUpdate.Enabled = False
        
End Sub



Private Sub txtAdCont_GotFocus()

    SendKeys highLig

End Sub

Private Sub txtAdCont_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtNPass.SetFocus
        SendKeys highLig
    End If
    
End Sub

Private Sub txtadDes_GotFocus()
    
    SendKeys highLig

End Sub

Private Sub txtadDes_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtAdCont.SetFocus
        SendKeys highLig
    End If
    
End Sub

Private Sub txtAdminId_GotFocus()

    SendKeys highLig
    
End Sub

Private Sub txtAdminId_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call cmdSearch_Click
    End If
    
End Sub

Private Sub txtAdName_GotFocus()

    SendKeys highLig

End Sub

Private Sub txtAdName_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtadDes.SetFocus
        SendKeys highLig
    End If
    
End Sub

Private Sub txtCNpass_GotFocus()

    SendKeys highLig

End Sub

Private Sub txtCNpass_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtOPass.SetFocus
        SendKeys highLig
    End If
    
End Sub

Private Sub txtNPass_GotFocus()

    SendKeys highLig

End Sub

Private Sub txtNPass_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtCNpass.SetFocus
        SendKeys highLig
    End If
    
End Sub

Private Sub txtOPass_GotFocus()

    SendKeys highLig

End Sub

Private Sub txtOPass_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdReload.SetFocus
    End If
End Sub


