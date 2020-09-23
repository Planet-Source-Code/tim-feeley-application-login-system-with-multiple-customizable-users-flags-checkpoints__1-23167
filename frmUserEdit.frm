VERSION 5.00
Begin VB.Form frmUserEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Editor"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
   ScaleHeight     =   3600
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3570
      TabIndex        =   7
      Top             =   3150
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Caption         =   "Editor"
      Height          =   3015
      Left            =   1980
      TabIndex        =   9
      Top             =   30
      Width           =   2625
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   960
         TabIndex        =   13
         Top             =   1020
         Width           =   1605
      End
      Begin VB.Frame Frame3 
         Caption         =   "Flags"
         Height          =   1575
         Left            =   90
         TabIndex        =   6
         Top             =   1350
         Width           =   2445
         Begin VB.ListBox lstFlags 
            Height          =   1260
            ItemData        =   "frmUserEdit.frx":0000
            Left            =   90
            List            =   "frmUserEdit.frx":0002
            Style           =   1  'Checkbox
            TabIndex        =   12
            Top             =   240
            Width           =   2265
         End
      End
      Begin VB.TextBox txtUserName 
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Top             =   630
         Width           =   1605
      End
      Begin VB.TextBox txtUserID 
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   210
         Left            =   135
         TabIndex        =   14
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User Name:"
         Height          =   210
         Left            =   90
         TabIndex        =   11
         Top             =   690
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User ID:"
         Height          =   210
         Left            =   330
         TabIndex        =   10
         Top             =   300
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Users"
      Height          =   3015
      Left            =   60
      TabIndex        =   8
      Top             =   30
      Width           =   1845
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   405
         Left            =   120
         Picture         =   "frmUserEdit.frx":0004
         TabIndex        =   1
         Top             =   2490
         Width           =   525
      End
      Begin VB.CommandButton cmdRem 
         Caption         =   "Rem"
         Height          =   405
         Left            =   690
         Picture         =   "frmUserEdit.frx":0106
         TabIndex        =   2
         Top             =   2490
         Width           =   525
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   405
         Left            =   1260
         Picture         =   "frmUserEdit.frx":0208
         TabIndex        =   3
         Top             =   2490
         Width           =   525
      End
      Begin VB.ListBox lstUsers 
         Height          =   2160
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmUserEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub reenum()
lstFlags.Clear
lstUsers.Clear
Dim lbEnum As New Collection
Dim lbEnum2 As New Collection
Call AccessFlag(afEnumerate, , , , , lbEnum)
For i = 1 To lbEnum.Count
    lstFlags.AddItem lbEnum(i) & " - " & AccessFlag(afGetDescription, lbEnum(i))
Next i
Call UserList(ulEnumerate, , , , , , , , , lbEnum2)
For i = 1 To lbEnum2.Count
    lstUsers.AddItem lbEnum2(i)
Next i
End Sub

Private Sub cmdAdd_Click()
lstUsers.Enabled = False
cmdAdd.Enabled = False
cmdRem.Enabled = False
cmdSave.Tag = "NEW"
cmdClose.Caption = "Cancel"
Call ClearList(lstFlags)
txtUserID = ""
txtUserName = ""
txtPassword = ""
End Sub

Private Sub cmdClose_Click()
If cmdClose.Caption = "Cancel" Then
    txtUserID = ""
    txtUserName = ""
    txtPassword = ""
    Call ClearList(lstFlags)
   MsgBox "User has NOT been added."
   cmdSave.Tag = ""
   cmdSave.Enabled = False
   cmdAdd.Enabled = True
   cmdRem.Enabled = True
   cmdClose.Caption = "Close"
   lstUsers.Enabled = True
Else
    Unload Me
End If
End Sub

Private Sub cmdRem_Click()
Call UserList(ulRemUser, lstUsers.Text)
Call reenum
MsgBox "User deleted."
End Sub

Private Sub cmdSave_Click()
If cmdSave.Tag = "NEW" Then
    If UserList(ulAddUser, txtUserID, txtUserName, SetList(lstFlags), txtPassword) = True Then
        MsgBox "User has been added."
        cmdSave.Tag = ""
        cmdSave.Enabled = False
        cmdAdd.Enabled = True
        cmdRem.Enabled = True
        cmdClose.Caption = "Close"
        txtUserID = ""
        txtUserName = ""
        txtPassword = ""
        Call ClearList(lstFlags)
        Call reenum
        lstUsers.Enabled = True
        lstUsers.SetFocus
        Exit Sub
    Else
        MsgBox "Could not add user, please verify entries and try again or click cancel to abort."
        Exit Sub
    End If
ElseIf cmdSave.Tag = "" Then
    If UserList(ulEditUser, lstUsers.Text, , , , txtUserID, txtUserName, SetList(lstFlags), txtPassword) = True Then
        MsgBox "User has been updated."
        Call reenum
        cmdSave.Enabled = False
        lstUsers.SetFocus
        Exit Sub
    Else
        MsgBox "Could not update user, please verify entries and try again or click cancel to abort."
        Exit Sub
    End If
End If
End Sub

Private Sub Form_Load()
Call reenum
End Sub

Private Sub lstFlags_Click()
cmdSave.Enabled = True
End Sub

Private Sub lstUsers_Click()
txtUserID = lstUsers.Text
txtUserName = UserList(ulGetRealName, lstUsers.Text)
txtPassword = UserList(ulGetPass, lstUsers.Text)
Call ClearList(lstFlags)
uFlags = UserList(ulGetFlags, lstUsers.Text)
For i = 1 To Len(uFlags)
    Call CheckList(lstFlags, Mid(uFlags, i, 1))
Next
cmdSave.Enabled = False
End Sub

Sub CheckList(lb As ListBox, flag)
For i = 0 To lb.ListCount - 1
    If Left(lb.List(i), 1) = flag Then lb.Selected(i) = True
Next
End Sub
Function SetList(lb As ListBox)
rv = ""
For i = 0 To lb.ListCount - 1
    If lb.Selected(i) = True Then rv = rv & Left(lb.List(i), 1)
Next
SetList = rv
End Function
Sub ClearList(lb As ListBox)
For i = 0 To lb.ListCount - 1
    lb.Selected(i) = False
Next
End Sub

Private Sub txtPassword_Change()
cmdSave.Enabled = True
End Sub

Private Sub txtUserID_Change()
cmdSave.Enabled = True
End Sub

Private Sub txtUserName_Change()
cmdSave.Enabled = True
End Sub
