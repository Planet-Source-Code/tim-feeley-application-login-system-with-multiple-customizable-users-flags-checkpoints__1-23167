VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Database Login"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   405
      Left            =   3270
      TabIndex        =   7
      Top             =   1230
      Width           =   915
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Log In"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   90
      TabIndex        =   6
      Top             =   1230
      Width           =   915
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2730
      PasswordChar    =   " "
      TabIndex        =   5
      Top             =   810
      Width           =   1455
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   1020
      TabIndex        =   4
      Top             =   810
      Width           =   765
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   4155
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Please enter your user name and password information to begin the program."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   30
         TabIndex        =   1
         Top             =   210
         Width           =   4065
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   210
      Left            =   1890
      TabIndex        =   3
      Top             =   855
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "User Name:"
      Height          =   210
      Left            =   90
      TabIndex        =   2
      Top             =   862
      Width           =   840
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogin_Click()
If UserList(ulUserExists, txtUserID) = False Then MsgBox "No user exists.": Exit Sub
If VerifyPassword(txtUserID, txtPassword) Then
    Form1.Tag = txtUserID
    Unload Me
    Form1.Show
Else
    MsgBox "Login error."
    txtUserID = ""
    txtPassword = ""
    txtUserID.SetFocus
End If
End Sub
Private Sub MODBuildCollection(vvDelim As Variant, vcolColl As Collection, vsInStr As String)
    Dim pnCnt As Integer
    Dim psHold As String
    Dim pnPos As Integer
    Dim pnPrev As Integer
    pnPrev = 0
    pnPos = 0


    For pnCnt = vcolColl.Count To 1 Step -1
        vcolColl.Remove pnCnt
    Next


    Do
        pnPrev = pnPos
        pnPos = InStr(pnPos + 1, Trim$(vsInStr), vvDelim)


        If pnPos > 0 Then
            psHold = Trim$(Mid$(Trim$(vsInStr), pnPrev + 1, (pnPos - 1) - pnPrev))
            vcolColl.Add psHold
        Else
            psHold = Trim$(Mid$(Trim$(vsInStr), pnPrev + 1, Len(Trim$(vsInStr))))


            If Trim$(psHold) <> "" Then
                vcolColl.Add psHold
            End If
            Exit Do
        End If
    Loop
End Sub

