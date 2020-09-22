VERSION 5.00
Begin VB.Form MainScr 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  :. Welcome To Student Management System...W.I.T"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11205
   Icon            =   "BdrFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "BdrFrm.frx":1042
   ScaleHeight     =   8985
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   4800
      ScaleHeight     =   3015
      ScaleWidth      =   4455
      TabIndex        =   13
      Top             =   3240
      Width           =   4455
      Begin VB.Label lblTime 
         BackColor       =   &H80000009&
         Caption         =   "----"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   49
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblDate 
         BackColor       =   &H80000009&
         Caption         =   "----"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   48
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "Time :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   47
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "Date :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblWelcome 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Image imgCen 
         Height          =   3000
         Left            =   0
         Picture         =   "BdrFrm.frx":CDD9
         Top             =   0
         Width           =   4500
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   2895
      Begin VB.PictureBox picDeletion 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   2655
         TabIndex        =   42
         Top             =   3840
         Width           =   2655
         Begin SysMan.desButton cmdDele 
            Height          =   375
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   2655
            _extentx        =   4683
            _extenty        =   661
            caption         =   "Record Deletion Section"
            capalign        =   2
            backstyle       =   2
            gradient        =   3
            cgradient       =   14737632
            mode            =   0
            value           =   0   'False
            cback           =   -2147483633
            cbhover         =   12632256
            lockhover       =   1
         End
         Begin VB.Label lblFTime 
            BackStyle       =   0  'Transparent
            Caption         =   "Delete Full-Time Records"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   45
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lblPTime 
            BackStyle       =   0  'Transparent
            Caption         =   "Delete Part-Time Records"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   44
            Top             =   1200
            Width           =   1695
         End
      End
      Begin VB.PictureBox picPTime 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   2655
         TabIndex        =   26
         Top             =   2640
         Width           =   2655
         Begin SysMan.desButton cmdPTime 
            Height          =   375
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   2655
            _extentx        =   4683
            _extenty        =   661
            caption         =   "Part Time Courses"
            capalign        =   2
            backstyle       =   2
            gradient        =   3
            cgradient       =   14737632
            mode            =   0
            value           =   0   'False
            cback           =   -2147483633
            cbhover         =   12632256
            lockhover       =   1
         End
         Begin VB.Label lblPTimePaySkm 
            BackColor       =   &H80000009&
            Caption         =   "Payment Options - SKM"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Label lblPTimePayDip 
            BackColor       =   &H80000009&
            Caption         =   "Payment Options - Diploma"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   1680
            Width           =   2415
         End
         Begin VB.Label lblPTimePayHnd 
            BackColor       =   &H80000009&
            Caption         =   "Payment Options - HND"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label lblPTEdit 
            BackColor       =   &H80000009&
            Caption         =   "Edit Part Time Courses"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label lblPTNew 
            BackColor       =   &H80000009&
            Caption         =   "New Part Time Courses"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   600
            Width           =   1935
         End
      End
      Begin VB.PictureBox picOther 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   2655
         TabIndex        =   20
         Top             =   3240
         Width           =   2655
         Begin SysMan.desButton cmdOther 
            Height          =   375
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   2655
            _extentx        =   4683
            _extenty        =   661
            caption         =   "Addtional Info's"
            capalign        =   2
            backstyle       =   2
            gradient        =   3
            cgradient       =   14737632
            mode            =   0
            value           =   0   'False
            cback           =   -2147483633
            cbhover         =   12632256
            lockhover       =   1
         End
         Begin VB.Label lblLogout 
            BackStyle       =   0  'Transparent
            Caption         =   "Log Out"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label lblAbout 
            BackStyle       =   0  'Transparent
            Caption         =   "About"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblHelp 
            BackStyle       =   0  'Transparent
            Caption         =   "System Help"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblLock 
            BackStyle       =   0  'Transparent
            Caption         =   "Lock System"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   600
            Width           =   1935
         End
      End
      Begin VB.PictureBox picAdmin 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   2655
         TabIndex        =   15
         Top             =   240
         Width           =   2655
         Begin SysMan.desButton cmdAdmin 
            Height          =   375
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   2655
            _extentx        =   4683
            _extenty        =   661
            caption         =   "Administrator"
            capalign        =   2
            backstyle       =   2
            gradient        =   3
            cgradient       =   14737632
            mode            =   0
            value           =   0   'False
            cback           =   -2147483633
            cbhover         =   12632256
            lockhover       =   1
         End
         Begin VB.Label lblDelAd 
            BackColor       =   &H80000009&
            Caption         =   "Delete Admin Profile"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label lblEditAd 
            BackColor       =   &H80000009&
            Caption         =   "Edit Admin Profile"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   18
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label lblNewAd 
            BackColor       =   &H80000009&
            Caption         =   "New Admin Profile"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   17
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.PictureBox picSkil 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   2655
         TabIndex        =   7
         Top             =   2040
         Width           =   2655
         Begin SysMan.desButton cmdSkil 
            Height          =   375
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   2655
            _extentx        =   4683
            _extenty        =   661
            caption         =   "Skill Courses"
            capalign        =   2
            backstyle       =   2
            gradient        =   3
            cgradient       =   14737632
            mode            =   0
            value           =   0   'False
            cback           =   -2147483633
            cbhover         =   12632256
            lockhover       =   1
         End
         Begin VB.Label lblSkmPay 
            BackColor       =   &H80000009&
            Caption         =   "Payment Options"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   32
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label lblSkNew 
            BackColor       =   &H80000009&
            Caption         =   "New Skill Profile"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   9
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label lblSkEdit 
            BackColor       =   &H80000009&
            Caption         =   "Edit Skill Profile"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   8
            Top             =   1080
            Width           =   1815
         End
      End
      Begin VB.PictureBox picDip 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   2655
         TabIndex        =   4
         Top             =   1440
         Width           =   2655
         Begin SysMan.desButton cmdDip 
            Height          =   375
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   2655
            _extentx        =   4683
            _extenty        =   661
            caption         =   "Diploma Programmes"
            capalign        =   2
            backstyle       =   2
            gradient        =   3
            cgradient       =   14737632
            mode            =   0
            value           =   0   'False
            cback           =   -2147483633
            cbhover         =   12632256
            lockhover       =   1
         End
         Begin VB.Label lblDipPay 
            BackColor       =   &H80000009&
            Caption         =   "Payment Options"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   31
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label lblDipNew 
            BackColor       =   &H80000009&
            Caption         =   "New Diploma Profile"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   6
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label lblEdiDip 
            BackColor       =   &H80000009&
            Caption         =   "Edit Diploma Profile"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   5
            Top             =   1080
            Width           =   1815
         End
      End
      Begin VB.PictureBox picHnd 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   2655
         TabIndex        =   1
         Top             =   840
         Width           =   2655
         Begin SysMan.desButton cmdHnd 
            Height          =   375
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   2655
            _extentx        =   4683
            _extenty        =   661
            caption         =   "Higher National Diploma ( HND)"
            capalign        =   2
            backstyle       =   2
            gradient        =   3
            cgradient       =   14737632
            mode            =   0
            value           =   0   'False
            cback           =   -2147483633
            cbhover         =   12632256
            lockhover       =   1
         End
         Begin VB.Label lblHndPay 
            BackColor       =   &H80000009&
            Caption         =   "Payment Options"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   30
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label lblNewHnd 
            BackColor       =   &H80000009&
            Caption         =   "New HND Profile"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   3
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label lblEditHnd 
            BackColor       =   &H80000009&
            Caption         =   "Edit HND Profile"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   2
            Top             =   1080
            Width           =   1815
         End
      End
   End
   Begin SysMan.desButton cmdSkmRpt 
      Height          =   375
      Left            =   4920
      TabIndex        =   36
      Top             =   2040
      Width           =   1935
      _extentx        =   4683
      _extenty        =   661
      caption         =   "Skm Report"
      capalign        =   2
      backstyle       =   2
      gradient        =   3
      cgradient       =   14737632
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
      cbhover         =   12632256
      lockhover       =   1
   End
   Begin SysMan.desButton cmdDipRpt 
      Height          =   375
      Left            =   8040
      TabIndex        =   37
      Top             =   1560
      Width           =   1935
      _extentx        =   4683
      _extenty        =   661
      caption         =   "Diploma Report"
      capalign        =   2
      backstyle       =   2
      gradient        =   3
      cgradient       =   14737632
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
      cbhover         =   12632256
      lockhover       =   1
   End
   Begin SysMan.desButton cmdHndRpt 
      Height          =   375
      Left            =   6000
      TabIndex        =   38
      Top             =   1560
      Width           =   1935
      _extentx        =   4683
      _extenty        =   661
      caption         =   "Hnd Report"
      capalign        =   2
      backstyle       =   2
      gradient        =   3
      cgradient       =   14737632
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
      cbhover         =   12632256
      lockhover       =   1
   End
   Begin SysMan.desButton cmdAdRpt 
      Height          =   375
      Left            =   3960
      TabIndex        =   39
      Top             =   1560
      Width           =   1935
      _extentx        =   4683
      _extenty        =   661
      caption         =   "Administrator Report"
      capalign        =   2
      backstyle       =   2
      gradient        =   3
      cgradient       =   14737632
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
      cbhover         =   12632256
      lockhover       =   1
   End
   Begin SysMan.desButton cmdPTRpt 
      Height          =   375
      Left            =   6960
      TabIndex        =   40
      Top             =   2040
      Width           =   1935
      _extentx        =   4683
      _extenty        =   661
      caption         =   "Part Time  Report"
      capalign        =   2
      backstyle       =   2
      gradient        =   3
      cgradient       =   14737632
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
      cbhover         =   12632256
      lockhover       =   1
   End
   Begin VB.Image imgWIT 
      Height          =   735
      Left            =   8520
      Picture         =   "BdrFrm.frx":136D1
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   7800
      X2              =   10440
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   7800
      X2              =   10440
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "shag@crashpoint'0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7920
      TabIndex        =   25
      Top             =   8400
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1470
      Left            =   960
      Picture         =   "BdrFrm.frx":192A3
      Top             =   0
      Width           =   10500
   End
End
Attribute VB_Name = "MainScr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum pStatus
    Max = 1
    Min = 0
End Enum

Public sAdmin As pStatus
Public sHND As pStatus
Public sDIP As pStatus
Public sSK As pStatus
Public sPT As pStatus
Public sOt As pStatus
Public sDel As pStatus

Public skPic As String
Public dipPic As String

Private Sub cmdAdmin_Click()

  Dim temp As Long
  
     MainScr.Refresh
    cmdAdmin.Refresh
    cmdHnd.Refresh
    cmdDip.Refresh
    cmdSkil.Refresh
    cmdPTime.Refresh
    cmdOther.Refresh

    If sAdmin = Max Then
        sAdmin = Min
        For temp = 1 To 1695 - 255
            picAdmin.Height = picAdmin.Height - 1
            picHND.Top = picAdmin.Top + picAdmin.Height + 20
            picDip.Top = picHND.Top + picHND.Height + 20
            picSkil.Top = picDip.Top + picDip.Height + 20
            picPTime.Top = picSkil.Top + picSkil.Height + 20
            picOther.Top = picPTime.Top + picPTime.Height + 20
            picDeletion.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next '  TEMP '  TEMP
        'cmdAdmin.BackColor = &H808080
     ElseIf sAdmin = Min Then
        sAdmin = Max
        If sHND = Max Then
            cmdHnd_Click
        End If
        If sDIP = Max Then
            cmdDip_Click
        End If
        If sSK = Max Then
            cmdSkil_Click
        End If
        If sPT = Max Then
            cmdPTime_Click
        End If
        If sOt = Max Then
            cmdOther_Click
        End If
        If sDel = Max Then
            cmdDele_Click
        End If
        For temp = 1 To 1695 - 255
            picAdmin.Height = picAdmin.Height + 1
            picHND.Top = picAdmin.Top + picAdmin.Height + 20
            picDip.Top = picHND.Top + picHND.Height + 20
            picSkil.Top = picDip.Top + picDip.Height + 20
            picPTime.Top = picSkil.Top + picSkil.Height + 20
            picOther.Top = picPTime.Top + picPTime.Height + 20
            picDeletion.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next '  TEMP '  TEMP
        'cmdAdmin.BackColor = &HE0E0E0
    End If


End Sub

Private Sub cmdAdRpt_Click()

    AdminRpt.Show vbModal
    
End Sub

Private Sub cmdDele_Click()

Dim temp As Long

    MainScr.Refresh
    cmdAdmin.Refresh
    cmdHnd.Refresh
    cmdDip.Refresh
    cmdSkil.Refresh
    cmdPTime.Refresh
    cmdOther.Refresh
    
    If sDel = Max Then
        sDel = Min
        For temp = 1 To 1695 - 255
            picDeletion.Height = picDeletion.Height - 1
            DoEvents
        Next '  TEMP '  TEMP
        'cmdSkil.BackColor = &H808080
     ElseIf sDel = Min Then
        sDel = Max
        If sAdmin = Max Then
            cmdAdmin_Click
        End If
        If sHND = Max Then
            cmdHnd_Click
        End If
        If sDIP = Max Then
            cmdDip_Click
        End If
        If sSK = Max Then
            cmdSkil_Click
        End If
        If sPT = Max Then
            cmdPTime_Click
        End If
        If sOt = Max Then
            cmdOther_Click
        End If
        For temp = 1 To 1695 - 255
            picDeletion.Height = picDeletion.Height + 1
            DoEvents
        Next '  TEMP '  TEMP
        'cmdSkil.BackColor = &HE0E0E0
    End If

End Sub

Private Sub cmdDip_Click()

  Dim temp As Long

    MainScr.Refresh
    cmdAdmin.Refresh
    cmdHnd.Refresh
    cmdDip.Refresh
    cmdSkil.Refresh
    cmdPTime.Refresh
    cmdOther.Refresh
    
    If sDIP = Max Then
        sDIP = Min
        For temp = 1 To 1695 - 255
            picDip.Height = picDip.Height - 1
            picSkil.Top = picDip.Top + picDip.Height + 20
            picPTime.Top = picSkil.Top + picSkil.Height + 20
            picOther.Top = picPTime.Top + picPTime.Height + 20
            picDeletion.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next '  TEMP '  TEMP
        'cmdDip.BackColor = &H808080
     ElseIf sDIP = Min Then
        sDIP = Max
        If sAdmin = Max Then
            cmdAdmin_Click
        End If
        If sHND = Max Then
            cmdHnd_Click
        End If
        If sSK = Max Then
            cmdSkil_Click
        End If
        If sPT = Max Then
            cmdPTime_Click
        End If
        If sOt = Max Then
            cmdOther_Click
        End If
        If sDel = Max Then
            cmdDele_Click
        End If
        For temp = 1 To 1695 - 255
            picDip.Height = picDip.Height + 1
            picSkil.Top = picDip.Top + picDip.Height + 20
            picPTime.Top = picSkil.Top + picSkil.Height + 20
            picOther.Top = picPTime.Top + picPTime.Height + 20
            picDeletion.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next '  TEMP '  TEMP
        'cmdDip.BackColor = &HE0E0E0
    End If

End Sub

Private Sub cmdDipRpt_Click()

    DipRpt.Show vbModal
    
End Sub

Private Sub cmdHnd_Click()

  Dim temp As Long
  
     MainScr.Refresh
    cmdAdmin.Refresh
    cmdHnd.Refresh
    cmdDip.Refresh
    cmdSkil.Refresh
    cmdPTime.Refresh
    cmdOther.Refresh

    If sHND = Max Then
        sHND = Min
        For temp = 1 To 1695 - 255
            picHND.Height = picHND.Height - 1
            picDip.Top = picHND.Top + picHND.Height + 20
            picSkil.Top = picDip.Top + picDip.Height + 20
            picPTime.Top = picSkil.Top + picSkil.Height + 20
            picOther.Top = picPTime.Top + picPTime.Height + 20
            picDeletion.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next '  TEMP '  TEMP
        'cmdHnd.BackColor = &H808080
     ElseIf sHND = Min Then
        sHND = Max
        If sDIP = Max Then
            cmdDip_Click
        End If
        If sSK = Max Then
            cmdSkil_Click
        End If
        If sAdmin = Max Then
            cmdAdmin_Click
        End If
        If sPT = Max Then
            cmdPTime_Click
        End If
        If sOt = Max Then
            cmdOther_Click
        End If
        If sDel = Max Then
            cmdDele_Click
        End If
        For temp = 1 To 1695 - 255
            picHND.Height = picHND.Height + 1
            picDip.Top = picHND.Top + picHND.Height + 20
            picSkil.Top = picDip.Top + picDip.Height + 20
            picPTime.Top = picSkil.Top + picSkil.Height + 20
            picOther.Top = picPTime.Top + picPTime.Height + 20
            picDeletion.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next '  TEMP '  TEMP
        'cmdHnd.BackColor = &HE0E0E0
    End If

End Sub

Private Sub cmdHndRpt_Click()

    HndRpt.Show vbModal
    
End Sub

Private Sub cmdOther_Click()

Dim temp As Long

    MainScr.Refresh
    cmdAdmin.Refresh
    cmdHnd.Refresh
    cmdDip.Refresh
    cmdSkil.Refresh
    cmdPTime.Refresh
    cmdOther.Refresh
    
    If sOt = Max Then
        sOt = Min
        For temp = 1 To 2655 - 255
            picOther.Height = picOther.Height - 1
            picDeletion.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next '  TEMP '  TEMP
        'cmdSkil.BackColor = &H808080
     ElseIf sOt = Min Then
        sOt = Max
        If sAdmin = Max Then
            cmdAdmin_Click
        End If
        If sHND = Max Then
            cmdHnd_Click
        End If
        If sDIP = Max Then
            cmdDip_Click
        End If
        If sSK = Max Then
            cmdSkil_Click
        End If
        If sPT = Max Then
            cmdPTime_Click
        End If
        If sDel = Max Then
            cmdDele_Click
        End If
        For temp = 1 To 2655 - 255
            picOther.Height = picOther.Height + 1
            picDeletion.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next '  TEMP '  TEMP
        'cmdSkil.BackColor = &HE0E0E0
    End If


End Sub

Private Sub cmdPTime_Click()

    Dim temp As Long

    MainScr.Refresh
    cmdAdmin.Refresh
    cmdHnd.Refresh
    cmdDip.Refresh
    cmdSkil.Refresh
    cmdPTime.Refresh
    cmdOther.Refresh
    
        If sPT = Max Then
            sPT = Min
            For temp = 1 To 2655 - 255
                picPTime.Height = picPTime.Height - 1
                picOther.Top = picPTime.Top + picPTime.Height + 20
                picDeletion.Top = picOther.Top + picOther.Height + 20
                DoEvents
            Next '  TEMP '  TEMP
            'cmdSkil.BackColor = &H808080
         ElseIf sPT = Min Then
            sPT = Max
            If sAdmin = Max Then
                cmdAdmin_Click
            End If
            If sHND = Max Then
                cmdHnd_Click
            End If
            If sDIP = Max Then
                cmdDip_Click
            End If
            If sSK = Max Then
                cmdSkil_Click
            End If
            If sOt = Max Then
                cmdOther_Click
            End If
            If sDel = Max Then
                cmdDele_Click
            End If
            For temp = 1 To 2655 - 255
                picPTime.Height = picPTime.Height + 1
                picOther.Top = picPTime.Top + picPTime.Height + 20
                picDeletion.Top = picOther.Top + picOther.Height + 20
                DoEvents
            Next '  TEMP '  TEMP
            'cmdSkil.BackColor = &HE0E0E0
    End If

End Sub

Private Sub cmdPTRpt_Click()

    PTimeRpt.Show vbModal
    
End Sub

Private Sub cmdSkil_Click()

Dim temp As Long
   
    MainScr.Refresh
    cmdAdmin.Refresh
    cmdHnd.Refresh
    cmdDip.Refresh
    cmdSkil.Refresh
    cmdPTime.Refresh
    cmdOther.Refresh
    

    If sSK = Max Then
        sSK = Min
        For temp = 1 To 1695 - 255
            picSkil.Height = picSkil.Height - 1
            picPTime.Top = picSkil.Top + picSkil.Height + 20
            picOther.Top = picPTime.Top + picPTime.Height + 20
            picDeletion.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next '  TEMP '  TEMP
        'cmdSkil.BackColor = &H808080
     ElseIf sSK = Min Then
        sSK = Max
        If sAdmin = Max Then
            cmdAdmin_Click
        End If
        If sHND = Max Then
            cmdHnd_Click
        End If
        If sDIP = Max Then
            cmdDip_Click
        End If
        If sPT = Max Then
            cmdPTime_Click
        End If
        If sOt = Max Then
            cmdOther_Click
        End If
        If sDel = Max Then
            cmdDele_Click
        End If
        For temp = 1 To 1695 - 255
            picSkil.Height = picSkil.Height + 1
            picPTime.Top = picSkil.Top + picSkil.Height + 20
            picOther.Top = picPTime.Top + picPTime.Height + 20
            picDeletion.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next '  TEMP '  TEMP
        'cmdSkil.BackColor = &HE0E0E0
    End If
    
End Sub


Private Sub cmdSkmRpt_Click()

    SkmRpt.Show vbModal
    
End Sub

Private Sub Form_Load()

    MainScr.Refresh
    cmdAdmin.Refresh
    cmdHnd.Refresh
    cmdDip.Refresh
    cmdSkil.Refresh
    cmdPTime.Refresh
    cmdOther.Refresh
    
    lblDate.Caption = Date
    lblTime.Caption = Time
    
End Sub

Private Sub Form_Terminate()

    If MsgBox("Are You Sure Exiting The System", vbYesNo, "SysMan") = vbYes Then
        End
    Else
        Exit Sub
    End If
    
End Sub

Private Sub lblDelAd_Click()

    AdminDel.Show vbModal
    
End Sub

Private Sub lblDelDip_Click()

    'dipPic = "C:\CrashPoint'0\Student Management System\Graphics\dip.jpg"
                
    With HNDStudDel
    
        .Show
        .Caption = "Delete Diploma Students Profile"
        '.imgPic.Picture = LoadPicture(dipPic)
        .lblDip.Visible = True
        .lblSkill.Visible = False
        .lblHNd.Visible = False
        
    End With
    
End Sub

Private Sub lblDelHnd_Click()

    HNDStudDel.Show vbModal
    HNDStudDel.lblHNd.Visible = True
    HNDStudDel.lblDip.Visible = False
    HNDStudDel.lblSkill.Visible = False
    
End Sub

Private Sub lblDipNew_Click()

    'dipPic = AppDir + "Images\DIP.jpg"
                
    With HNDStudNew
                
        .Show
        .Caption = "New Diploma Programmes Student Registration"
        '.imgPic.Picture = LoadPicture(App.Path & "\Images\DIP.jpg")
                    
        .optAE.Visible = False
        .optEE.Visible = False
        .optME.Visible = False
        .optSE.Visible = False
        .optJIP.Visible = False
        .optJJ.Visible = False
        .optJMI.Visible = False
        .optJMK.Visible = False
        .optJSK.Visible = False
                    
        .optAuto.Visible = True
        .optDEE.Visible = True
        .optDME.Visible = True
        .optDSC.Visible = True
                    
    End With
    
End Sub

Private Sub lblDipPay_Click()

    With PaymentFrmDip
        .Show
    End With

End Sub

Private Sub lblEdiDip_Click()

    'dipPic = "C:\CrashPoint'0\Student Management System\Graphics\dip.jpg"
                
    With HNDStudEdit
    
        .Show
        .Caption = "Edit Diploma Programmes Student Profile"
        '.imgPic.Picture = LoadPicture(dipPic)
        .lblDip.Visible = True
        .lblHigh.Visible = False
        .lblskm.Visible = False
                    
    End With
    
End Sub

Private Sub lblEditAd_Click()

    AdminEdit.Show vbModal
    
End Sub

Private Sub lblEditHnd_Click()

    HNDStudEdit.Show vbModal
    HNDStudEdit.lblHigh.Visible = True
    HNDStudEdit.lblDip.Visible = False
    HNDStudEdit.lblskm.Visible = False
    
End Sub

Private Sub lblFTime_Click()

    HNDStudDel.Show vbModal

End Sub

Private Sub lblHndPay_Click()

        PaymentForm.Show

End Sub

Private Sub lblLock_Click()

    ScreenLock.Show vbModal
    
End Sub

Private Sub lblLogout_Click()

    Unload MainScr
    AdminLogin.Show
    
End Sub

Private Sub lblNewAd_Click()

    AdminNew.Show vbModal
    
End Sub

Private Sub lblNewHnd_Click()

    HNDStudNew.Show vbModal
    
End Sub

Private Sub lblPTDel_Click()

    PTimeDelSelect.Show vbModal
    
End Sub

Private Sub lblPTEdit_Click()

    PTimeEditSelect.Show vbModal
    
End Sub

Private Sub lblPTimePay_Click()

End Sub

Private Sub lblPTime_Click()

    PartTimeDelDip.Show vbModal
    
End Sub

Private Sub lblPTimePayDip_Click()

    PTPayFrmDip.Show vbModal

End Sub

Private Sub lblPTimePayHnd_Click()

    PTPayHndFrm.Show vbModal

End Sub

Private Sub lblPTimePaySkm_Click()

    PTPayFrmSkm.Show vbModal

End Sub

Private Sub lblPTNew_Click()

    PTimeSelect.Show vbModal
    
End Sub

Private Sub lblSkDel_Click()

    'skPic = "C:\CrashPoint'0\Student Management System\Graphics\skillc.jpg"
                
    With HNDStudDel
                    
        .Show
        .Caption = "Delete Skilled Courses Student Profile"
        '.imgPic.Picture = LoadPicture(skPic)
        .lblSkill.Visible = True
        .lblDip.Visible = False
        .lblHNd.Visible = False
        
    End With
    
End Sub

Private Sub lblSkEdit_Click()

    'skPic = "C:\CrashPoint'0\Student Management System\Graphics\skillc.jpg"
                
    With HNDStudEdit
                
        .Show
        .Caption = "Edit Skilled Courses Student Profile"
        '.imgPic.Picture = LoadPicture(skPic)
        .cmbCCode1.Visible = True
        .cmbCCode1.Locked = True
                    
        .lblskm.Visible = True
        .lblHigh.Visible = False
        .lblDip.Visible = False
                                                                                              
    End With
    
End Sub

Private Sub lblSkmPay_Click()
    
    With PaymentFrmSkm
        .Show
    End With

End Sub

Private Sub lblSkNew_Click()
                       
    With HNDStudNew
        .optAE.Visible = False
        .optEE.Visible = False
        .optME.Visible = False
        .optSE.Visible = False
                
        .Show
        '.imgPic.Picture = LoadPicture("")
            
        ' skPic = "C:\CrashPoint'0\Student Management System\Graphics\SkillC.jpg"

        '.imgPic.Picture = LoadPicture(skPic)
        .Caption = "New Skilled Courses Student Registration"
                
        .optJMK.Visible = True
        .Label13.Visible = True
        .cmbCode1.Visible = True
                
        .optJSK.Visible = True
        .Label44.Visible = True
        .cmbCode3.Visible = True
            
        .optJMI.Visible = True
        .Label26.Visible = True
        .cmbCode2.Visible = True
                
        .optJIP.Visible = True
        .Label48.Visible = True
        .cmbCode4.Visible = True
                
        .optJJ.Visible = True
        .Label47.Visible = True
        .cmbCode5.Visible = True
    
    End With
              
End Sub

