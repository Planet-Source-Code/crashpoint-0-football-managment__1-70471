VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form DelPlayer 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  :. Delete Player Information"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Search Criteria"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txtSearchId 
         Height          =   405
         Left            =   2280
         TabIndex        =   16
         Top             =   240
         Width           =   2415
      End
      Begin Project1.desButton cmdSearch 
         Height          =   375
         Left            =   4800
         TabIndex        =   15
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         Caption         =   "Enter Player Id :-"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Caption         =   "Coresponding Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   6015
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00400000&
         ForeColor       =   &H0000FFFF&
         Height          =   1455
         Left            =   240
         ScaleHeight     =   1395
         ScaleWidth      =   5355
         TabIndex        =   5
         Top             =   360
         Width           =   5415
         Begin VB.Line Line1 
            BorderColor     =   &H0000FFFF&
            X1              =   0
            X2              =   5280
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Label lblmes 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "FOUND !!!"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Left            =   1320
            TabIndex        =   13
            Top             =   0
            Width           =   2655
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Player Name :"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Id No : -"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Reg No :-"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblstuName 
            BackStyle       =   0  'Transparent
            Caption         =   "############"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   1680
            TabIndex        =   9
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label lblStuIc 
            BackStyle       =   0  'Transparent
            Caption         =   "############"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   1680
            TabIndex        =   8
            Top             =   720
            Width           =   2655
         End
         Begin VB.Label lblstuCo 
            BackStyle       =   0  'Transparent
            Caption         =   "############"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   1680
            TabIndex        =   7
            Top             =   960
            Width           =   3735
         End
         Begin VB.Label lblErr 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "There Is No Such Record Found !!!"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   735
            Left            =   720
            TabIndex        =   6
            Top             =   480
            Width           =   3975
         End
      End
   End
   Begin VB.TextBox txtAdPass 
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   480
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3240
      Width           =   5295
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000009&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   6015
      Begin Project1.desButton cmdClose 
         Height          =   375
         Left            =   4440
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
      Begin Project1.desButton cmdDelete 
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc adoAdmin 
      Height          =   375
      Left            =   480
      Top             =   4920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Caption         =   "Adodc1"
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
   Begin MSAdodcLib.Adodc adostPer 
      Height          =   375
      Left            =   1800
      Top             =   4920
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      Caption         =   "Adodc1"
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
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   120
      Top             =   2880
      Width           =   6015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Enter Admin Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   18
      Top             =   3000
      Width           =   3255
   End
End
Attribute VB_Name = "DelPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    Unload DelPlayer
    
End Sub

Private Sub cmdDelete_Click()

 On Error Resume Next
    
        APass = adoAdmin.Recordset.Fields("Password")
        
            If txtAdPass.Text = APass Then
                
                If adostPer.Recordset.RecordCount = 0 Then
                    MsgBox "There is no Player's Records Available", vbInformation, "SysMan"
                    Unload Me
                    Exit Sub
                End If
                
                If MsgBox("The Following Record Will Be Deleted !!!" & vbCrLf & _
                        " Do You Proceed ?", vbYesNo + vbInformation, "SysMan") = vbYes Then
                    
                    adostPer.Recordset.delete
                    adostPer.Refresh
                    txtAdPass.Text = ""
                    txtAdPass.BackColor = &HC0C0C0
                    txtAdPass.Locked = True
                    cmdDelete.Enabled = False
                    Picture2.Cls
                    Exit Sub
                Else
                    Exit Sub
                End If
            Else
                MsgBox "The Password Doesn't Match !!!", vbInformation, "SysMan"
                Exit Sub
            End If
End Sub

Private Sub cmdSearch_Click()

  On Error GoTo noData
    
        If Trim(txtSearchId.Text) = "" Then
            Call missing
            txtSearchId.SetFocus
            Exit Sub
        End If
        
        If adostPer.Recordset.RecordCount = 0 Then
            MsgBox "There is no Player's record found in database", vbInformation, "SysMan"
            Exit Sub
        End If
        
            adostPer.Refresh
                
            adostPer.Recordset.Find ("ID_No = '" & Trim(txtSearchId.Text) & "'")
            
            On Error Resume Next
            
                PpName = adostPer.Recordset.Fields("Name")
                PpId = adostPer.Recordset.Fields("ID_No")
                PpRegId = adostPer.Recordset.Fields("Reg_No")
                PpClub = adostPer.Recordset.Fields("Club")
                PpPosi = adostPer.Recordset.Fields("Position")
                Ppdob = adostPer.Recordset.Fields("DOB")
                Ppstate = adostPer.Recordset.Fields("State")
                Ppdoj = adostPer.Recordset.Fields("DOJ")
                PpTFrom = adostPer.Recordset.Fields("TFrom")
                PpStatus = adostPer.Recordset.Fields("Status")
                PpYCrd = adostPer.Recordset.Fields("Yellow_Crd")
                PpRCrd = adostPer.Recordset.Fields("Red_Crd")
                PpPic = adostPer.Recordset.Fields("Picture")
                       
                    lblmes.Visible = True
                    lblmes.Caption = "Found !!!"
                    lblstuName.Visible = True
                    lblstuName.Caption = PpName
                    lblStuIc.Visible = True
                    lblStuIc.Caption = PpId
                    lblstuCo.Visible = True
                    lblstuCo.Caption = PpRegId
                    
                    lblErr.Visible = False
                    Label4.Visible = True
                    Label5.Visible = True
                    Label6.Visible = True
                    
                    txtAdPass.Enabled = True
                    txtAdPass.Locked = False
                    txtAdPass.BackColor = &H80000005
                    txtAdPass.SetFocus
                    
                    cmdDelete.Enabled = True

    Exit Sub
    
noData:

    lblmes.Caption = "Not Found !!!"
    lblstuName.Visible = False
    lblStuIc.Visible = False
    lblstuCo.Visible = False
    lblErr.Visible = True
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    
    txtAdPass.Enabled = False
    txtAdPass.Locked = True
    txtAdPass.BackColor = &HC0C0C0
    txtAdPass.Text = ""

    
End Sub

Private Sub Form_Load()

    Call DataConn(adostPer, "Players")
    Call DataConn(adoAdmin, "Admin")

End Sub
