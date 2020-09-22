Attribute VB_Name = "Declaration"
Option Explicit


Public AName As String 'stores librarian first name
Public AId As String 'stores librarian middle name
Public AUsername As String 'stores librarian last name
Public APass As String 'stores librarian password
Public AAdd1 As String 'stores librarian first name
Public AAdd2 As String 'stores librarian middle name
Public APoscode As String 'stores librarian last name
Public ACity As String 'stores librarian password
Public AState As String 'stores librarian first name
Public ACountry As String 'stores librarian middle name
Public ATel As String 'stores librarian last name
Public AEMail As String 'stores librarian password
Public APic As String

Public CName As String 'stores librarian first name
Public CId As String 'stores librarian middle name
Public CUsername As String 'stores librarian last name
Public CPass As String 'stores librarian password
Public CAdd1 As String 'stores librarian first name
Public CAdd2 As String 'stores librarian middle name
Public CPoscode As String 'stores librarian last name
Public CCity As String 'stores librarian password
Public CState As String 'stores librarian first name
Public CCountry As String 'stores librarian middle name
Public CTel As String 'stores librarian last name
Public CEMail As String 'stores librarian password
Public CiPic As String

Public PName As String 'stores librarian first name
Public PId As String 'stores librarian middle name
Public PRegId As String 'stores librarian last name
Public PClub As String 'stores librarian password
Public PPosi As String 'stores librarian first name
Public Pdob As String 'stores librarian middle name
Public Pstate As String 'stores librarian last name
Public Pdoj As String 'stores librarian password
Public PTFrom As String 'stores librarian first name
Public PStatus As String 'stores librarian middle name
Public PYCrd As String 'stores librarian last name
Public PRCrd As String 'stores librarian password
Public PPic As String

Public PpName As String 'stores librarian first name
Public PpId As String 'stores librarian middle name
Public PpRegId As String 'stores librarian last name
Public PpClub As String 'stores librarian password
Public PpPosi As String 'stores librarian first name
Public Ppdob As String 'stores librarian middle name
Public Ppstate As String 'stores librarian last name
Public Ppdoj As String 'stores librarian password
Public PpTFrom As String 'stores librarian first name
Public PpStatus As String 'stores librarian middle name
Public PpYCrd As String 'stores librarian last name
Public PpRCrd As String 'stores librarian password
Public PpPic As String

Public CMName As String
Public CCName As String
Public CFound As String
Public CLea As String
Public CPic As String

Public tempUser As String
Public tempPass As String

Public tempUser2 As String
Public tempPass2 As String


Public Const highLig = "{HOME}+{END}"
Public Const HiLyt = "{HOME}+{END}"

Public Main_On As Boolean

Public Sub CenterFrm(ByVal Parentfrm As MDIForm, ByVal Childfrm As Form) 'used for the frmInsignia

    Childfrm.Left = (Parentfrm.Width \ 2) - (Childfrm.Width \ 2)
    Childfrm.Top = (Parentfrm.ScaleHeight \ 2) - (Childfrm.Height \ 2)

End Sub

