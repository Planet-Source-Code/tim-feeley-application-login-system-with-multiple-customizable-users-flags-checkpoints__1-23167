VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Database Application"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Load Sample Application"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1860
      TabIndex        =   3
      Top             =   510
      Width           =   2775
   End
   Begin VB.CommandButton cmdUE 
      Caption         =   "User Editor"
      Height          =   555
      Left            =   90
      TabIndex        =   2
      Top             =   1260
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Checkpoint Editor"
      Height          =   585
      Left            =   90
      TabIndex        =   1
      Top             =   600
      Width           =   1515
   End
   Begin VB.CommandButton A 
      Caption         =   "Flag Editor"
      Height          =   465
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   1515
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub A_Click()
frmAccessEdit.Show

End Sub

Private Sub cmdUE_Click()
frmUserEdit.Show
End Sub

Private Sub Command1_Click()
frmCheckpoints.Show
End Sub

Private Sub Command2_Click()
frmLogin.Show
End Sub
