VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form AdminNew 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  :. Admin Setup - Football  Management System..."
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   1680
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoAdmin 
      Height          =   330
      Left            =   360
      Top             =   7440
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
      Caption         =   "Admin Login"
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
   Begin VB.Frame Frame3 
      BackColor       =   &H80000009&
      Height          =   735
      Left            =   0
      TabIndex        =   14
      Top             =   6360
      Width           =   5055
      Begin Project1.desButton cmdClose 
         Height          =   375
         Left            =   3360
         TabIndex        =   21
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
      Begin Project1.desButton cmdReset 
         Height          =   375
         Left            =   1800
         TabIndex        =   20
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Reset"
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
      Begin Project1.desButton cmdSave 
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Save"
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
      Caption         =   "Admin Security Settings"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   7
      Top             =   4560
      Width           =   5055
      Begin VB.TextBox txtPassConfirm 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtAdminPass 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtAdminId 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1800
         TabIndex        =   11
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000009&
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         Caption         =   "Admin Password"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         Caption         =   "Admin Id"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Admin Info"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   5055
      Begin Project1.desButton cmdBrowse 
         Height          =   735
         Left            =   3600
         TabIndex        =   18
         Top             =   2160
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
      Begin VB.PictureBox picAdmin 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   1560
         ScaleHeight     =   1815
         ScaleWidth      =   1815
         TabIndex        =   16
         Top             =   1800
         Width           =   1815
         Begin VB.Label Label8 
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
            Left            =   360
            TabIndex        =   17
            Top             =   600
            Width           =   1095
         End
         Begin VB.Image imgAdmin 
            BorderStyle     =   1  'Fixed Single
            Height          =   1815
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1815
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H0000FFFF&
            Height          =   1815
            Left            =   0
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.TextBox txtContact 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   6
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtDesign 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   5
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtAdName 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "ID Number (IC)"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "Picture"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "Contact No #"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "Admin Name"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
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
         Width           =   1335
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "NEW ADMIN PROFILE"
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
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "AdminNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim picName As String
Dim spicname As String

Private Sub cmdBrowse_Click()

    dlgCommon.Filter = "Picture Files(*.jpg; *.bmp; *.gif)|*.jpg;*.bmp;*.gif"
    dlgCommon.ShowOpen
    
        picName = dlgCommon.FileName
        
            spicname = Mid$(picName, InStrRev(picName, "/") + 1)
        
            If picName <> "" Then
                imgAdmin.Picture = LoadPicture(picName)
                Label8.Caption = ""
            End If
            
End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdReset_Click()

    'reset all the fields
    
    txtAdName.Text = ""
    txtDesign.Text = ""
    txtAdminId.Text = ""
    txtAdminPass.Text = ""
    txtPassConfirm.Text = ""
    txtContact.Text = ""
    
    imgAdmin.Picture = LoadPicture("")
    
    txtAdName.SetFocus
    
End Sub

Private Sub cmdSave_Click()

On Error GoTo ErrHandler 'handles the unwanted / unexpected errors

    If Trim(txtAdName.Text) = "" Or Trim(txtDesign.Text) = "" _
        Or Trim(txtContact.Text) = "" Or Trim(txtAdminId.Text) = "" _
            Or Trim(txtAdminPass.Text) = "" Or Trim(txtPassConfirm.Text) = "" Then
        
        Call missing
        
        'checks the data and focus on it
        If Trim(txtAdName.Text) = "" Then
            txtAdName.Text = ""
            txtAdName.SetFocus
        ElseIf Trim(txtDesign.Text) = "" Then
            txtDesign.Text = ""
            txtDesign.SetFocus
        ElseIf Trim(txtContact.Text) = "" Then
            txtContact.Text = ""
            txtContact.SetFocus
        ElseIf Trim(txtAdminId.Text) = "" Then
            txtAdminId.Text = ""
            txtAdminId.SetFocus
        ElseIf Trim(txtAdminPass.Text) = "" Then
            txtAdminPass.Text = ""
            txtAdminPass.SetFocus
        Else
            Trim(txtPassConfirm.Text) = ""
            txtPassConfirm.Text = ""
            txtPassConfirm.SetFocus
        
        End If
    
        Exit Sub
    End If
    
    If IsNumeric(txtContact.Text) = False Then
        MsgBox "Please Enter Numeric Value Only.", vbOKOnly + vbExclamation, "SysMan"
        txtContact.SetFocus
        SendKeys highLig
        Exit Sub
    End If
    
    'validate all data and checkin into database
    
    If Trim(txtAdminPass.Text) = Trim(txtPassConfirm.Text) Then
        
        adoAdmin.Refresh
        adoAdmin.Recordset.AddNew
    
        With adoAdmin.Recordset
            .Fields(0) = txtAdName.Text
            .Fields(1) = txtDesign.Text
            .Fields(2) = txtAdminId.Text
            .Fields(3) = txtAdminPass.Text
            .Fields(4) = txtContact.Text
            .Fields(5) = picName
            
        End With
        adoAdmin.Recordset.Update
        
        adoAdmin.Refresh
     
        
        adoAdmin.Recordset.MoveFirst 'show error if no data entered
        
        txtAdName.Text = ""
        txtDesign.Text = ""
        txtContact.Text = ""
        txtAdminId.Text = ""
        txtAdminPass.Text = ""
        txtPassConfirm.Text = ""
        
        'Confirmation of data entry
        If MsgBox("Administrator Profile Succesfully Entered." & vbCrLf & _
            " Enter Another Admin Profile ?", vbYesNo + vbQuestion, "SysMan") = vbYes Then
            adoAdmin.Refresh
            imgAdmin.Picture = LoadPicture("")
            Exit Sub
        Else
            AdminNew.Hide
                           
            adoAdmin.Refresh
           
            Unload AdminNew
            Exit Sub
        End If
        
    Else
        MsgBox "The Password and Password Confirmation Must Be Same." & vbCrLf + vbCrLf & _
                "Please re-confirm/enter password", vbOKOnly + vbExclamation, "SysMan"
        txtPassConfirm.Text = ""
        txtPassConfirm.SetFocus
        Exit Sub
    End If
    Exit Sub
    
ErrHandler:
        MsgBox "User name already exists. Please specify a different user name.", vbOKOnly, "SysMan"
        txtAdName.SetFocus
        SendKeys highLig
        Exit Sub


End Sub

Private Sub Form_Activate()

    txtAdName.SetFocus
    
End Sub

Private Sub Form_Load()

  Call DataConn(adoAdmin, "Admin")

    On Error GoTo ErrHandler
    
        'txtAdName.SetFocus

        Exit Sub
    
ErrHandler:
    MsgBox "Its Seems You Are Using This System For First Time..." & vbCrLf & _
        "Welcome To FAS System....", vbExclamation + vbInformation, "SysMan"
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'AdoAdmin.Recordset = Nothing
    
End Sub

Private Sub txtAdminId_GotFocus()

    SendKeys highLig

End Sub

Private Sub txtAdminId_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtAdminPass.SetFocus
        SendKeys highLig
    End If
    
End Sub

Private Sub txtAdminPass_GotFocus()

    SendKeys highLig

End Sub

Private Sub txtAdminPass_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtPassConfirm.SetFocus
        SendKeys highLig
    End If
    
End Sub

Private Sub txtAdName_GotFocus()

    SendKeys highLig
    
End Sub

Private Sub txtAdName_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtDesign.SetFocus
        SendKeys highLig
    End If
    
End Sub

Private Sub txtContact_GotFocus()

    SendKeys highLig

End Sub

Private Sub txtContact_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdBrowse.SetFocus
        txtAdminId.SetFocus
        SendKeys highLig
    End If
    
End Sub

Private Sub txtDesign_GotFocus()

    SendKeys highLig

End Sub

Private Sub txtDesign_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtContact.SetFocus
        SendKeys highLig
    End If
    
End Sub

Private Sub txtPassConfirm_GotFocus()

    SendKeys highLig

End Sub

Private Sub txtPassConfirm_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdSave.SetFocus
    End If
End Sub
