VERSION 5.00
Begin VB.Form frmAccessEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flag Editor"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
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
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picAdd 
      BorderStyle     =   0  'None
      Height          =   2505
      Left            =   30
      ScaleHeight     =   2505
      ScaleWidth      =   3315
      TabIndex        =   8
      Top             =   30
      Visible         =   0   'False
      Width           =   3315
      Begin VB.CommandButton cmdAddCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2250
         TabIndex        =   14
         Top             =   1740
         Width           =   855
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "Ok"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1740
         Width           =   855
      End
      Begin VB.TextBox txtFlagDesc 
         Height          =   825
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   780
         Width           =   2955
      End
      Begin VB.TextBox txtFlagChar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         MaxLength       =   1
         TabIndex        =   10
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   210
         Left            =   150
         TabIndex        =   11
         Top             =   570
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Flag Character:"
         Height          =   210
         Left            =   150
         TabIndex        =   9
         Top             =   0
         Width           =   1110
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   405
      Left            =   2430
      TabIndex        =   7
      Top             =   2100
      Width           =   885
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   315
      Left            =   2430
      TabIndex        =   6
      Top             =   900
      Width           =   855
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   315
      Left            =   2430
      TabIndex        =   5
      Top             =   510
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   2430
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Flags"
      Height          =   2505
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.TextBox txtDescription 
         BackColor       =   &H00C0C0C0&
         Height          =   525
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1830
         Width           =   2115
      End
      Begin VB.ListBox lstFlags 
         Height          =   1320
         Left            =   90
         TabIndex        =   1
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   210
         Left            =   90
         TabIndex        =   2
         Top             =   1590
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAccessEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
txtFlagChar.Tag = "ADD"
picAdd.Visible = True
txtFlagChar = ""
txtFlagDesc = ""
txtFlagChar.SetFocus
End Sub

Private Sub cmdAddCancel_Click()
picAdd.Visible = False
End Sub

Private Sub cmdAddNew_Click()
If txtFlagChar.Tag = "ADD" Then If AccessFlag(afAddFlag, txtFlagChar, txtFlagDesc) = True Then Call reenum: picAdd.Visible = False
If txtFlagChar.Tag = "EDIT" Then If AccessFlag(afEditFlag, lstFlags.Text, , txtFlagChar, txtFlagDesc) = True Then Call reenum: picAdd.Visible = False
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
picAdd.Visible = True
txtFlagChar.Tag = "EDIT"
txtFlagChar = lstFlags.Text
txtFlagDesc = AccessFlag(afGetDescription, lstFlags.Text)
End Sub

Private Sub cmdRemove_Click()
If AccessFlag(afRemFlag, lstFlags.Text) = True Then Call reenum: picAdd.Visible = False
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Call reenum
End Sub

Private Sub lstFlags_Click()
txtDescription = AccessFlag(afGetDescription, lstFlags.Text)
End Sub
Sub reenum()
lstFlags.Clear
Dim lbEnum As New Collection
Call AccessFlag(afEnumerate, , , , , lbEnum)
For i = 1 To lbEnum.Count
    lstFlags.AddItem lbEnum(i)
Next i

End Sub
