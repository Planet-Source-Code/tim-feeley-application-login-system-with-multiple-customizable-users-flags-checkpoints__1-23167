VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MySoft Invoices"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Close"
      Height          =   435
      Left            =   2160
      TabIndex        =   7
      Top             =   2370
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Return Value"
      Height          =   585
      Left            =   210
      TabIndex        =   6
      Top             =   1680
      Width           =   2985
      Begin VB.Label lblReturn 
         Alignment       =   2  'Center
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
         Left            =   90
         TabIndex        =   8
         Top             =   210
         Width           =   2835
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Flag 4 Prohibited"
      Height          =   465
      Left            =   1710
      TabIndex        =   5
      Top             =   1080
      Width           =   1605
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Flags 1,2 Prohibited"
      Height          =   465
      Left            =   1710
      TabIndex        =   4
      Top             =   570
      Width           =   1605
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Flag 1 Prohibited"
      Height          =   465
      Left            =   1710
      TabIndex        =   3
      Top             =   60
      Width           =   1605
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Flags 1,2,3,4, Req'd"
      Height          =   465
      Left            =   60
      TabIndex        =   2
      Top             =   1080
      Width           =   1605
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Flags 1,2 Required"
      Height          =   465
      Left            =   60
      TabIndex        =   1
      Top             =   570
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Flag 1 Required"
      Height          =   465
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1605
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If CPVerify(Form1.Tag, "1") Then lblReturn = "True" Else lblReturn = "False"

End Sub

Private Sub Command2_Click()
If CPVerify(Form1.Tag, "2") Then lblReturn = "True" Else lblReturn = "False"
End Sub

Private Sub Command3_Click()
If CPVerify(Form1.Tag, "3") Then lblReturn = "True" Else lblReturn = "False"
End Sub

Private Sub Command4_Click()
If CPVerify(Form1.Tag, "-1") Then lblReturn = "True" Else lblReturn = "False"
End Sub

Private Sub Command5_Click()
If CPVerify(Form1.Tag, "-2") Then lblReturn = "True" Else lblReturn = "False"
End Sub

Private Sub Command6_Click()
If CPVerify(Form1.Tag, "-3") Then lblReturn = "True" Else lblReturn = "False"
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Form_Load()
' First, it is a good idea to set some Tag or Global Variable
' to equal the user name. For this, Form1.Tag will be used.
' To see the actual code, browse the frmLogin cmdLogin click event.

End Sub
