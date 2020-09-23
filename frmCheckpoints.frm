VERSION 5.00
Begin VB.Form frmCheckpoints 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Checkpoint Editor"
   ClientHeight    =   2775
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6090
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
   ScaleHeight     =   2775
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   2460
      TabIndex        =   20
      Top             =   2100
      Width           =   885
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2460
      TabIndex        =   19
      ToolTipText     =   "Edits the highlighted flag."
      Top             =   1290
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Editor"
      Height          =   2505
      Left            =   3390
      TabIndex        =   7
      Top             =   0
      Width           =   2655
      Begin VB.PictureBox picHide 
         Height          =   2235
         Left            =   60
         ScaleHeight     =   2175
         ScaleWidth      =   2445
         TabIndex        =   17
         Top             =   210
         Width           =   2505
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Not available"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   30
            TabIndex        =   18
            Top             =   960
            Width           =   2385
         End
      End
      Begin VB.ListBox lstProhibit 
         Height          =   1320
         Left            =   1770
         Sorted          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "A user will fail the checkpoint if ANY of the prohibited flags are in his/her access."
         Top             =   1050
         Width           =   645
      End
      Begin VB.ListBox lstRequired 
         Height          =   1320
         Left            =   1020
         Sorted          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "To pass the checkpoint, the user must have all of the flags in the 'Required Flags' list box."
         Top             =   1050
         Width           =   645
      End
      Begin VB.ListBox lstFlags 
         Height          =   1320
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Flags that may be used. Right click an item for more options."
         Top             =   1050
         Width           =   735
      End
      Begin VB.TextBox txtCheckpoint 
         Height          =   315
         Left            =   1170
         TabIndex        =   9
         ToolTipText     =   "Displays the Checkpoint ID name."
         Top             =   210
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Prohibited Flags:"
         Height          =   420
         Left            =   1770
         TabIndex        =   15
         ToolTipText     =   "A user will fail the checkpoint if ANY of the prohibited flags are in his/her access."
         Top             =   600
         Width           =   720
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Required Flags:"
         Height          =   420
         Left            =   1020
         TabIndex        =   14
         ToolTipText     =   "To pass the checkpoint, the user must have all of the flags in the 'Required Flags' list box."
         Top             =   600
         Width           =   705
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Available Flags:"
         Height          =   420
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   705
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Checkpoint ID:"
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Checkpoints"
      Height          =   2505
      Left            =   90
      TabIndex        =   3
      Top             =   0
      Width           =   2295
      Begin VB.ListBox lstCheckpoints 
         Height          =   1320
         ItemData        =   "frmCheckpoints.frx":0000
         Left            =   90
         List            =   "frmCheckpoints.frx":0002
         TabIndex        =   5
         ToolTipText     =   "Listing of current checkpoints. Highlight and click Remove or Edit to do the corresponding function"
         Top             =   240
         Width           =   2115
      End
      Begin VB.TextBox txtContents 
         BackColor       =   &H00C0C0C0&
         Height          =   525
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         ToolTipText     =   "Previews the contents of the checkpoint."
         Top             =   1830
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contents:"
         Height          =   210
         Left            =   90
         TabIndex        =   6
         Top             =   1620
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   2460
      TabIndex        =   2
      ToolTipText     =   "Enables the editor to add a new checkpoint entry."
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   315
      Left            =   2460
      TabIndex        =   1
      ToolTipText     =   "Removes the highligted flag."
      Top             =   510
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   315
      Left            =   2460
      TabIndex        =   0
      ToolTipText     =   "Edits the highlighted flag."
      Top             =   900
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "For help on how to use an item, hover the mouse pointer over it."
      Enabled         =   0   'False
      Height          =   210
      Left            =   30
      TabIndex        =   16
      Top             =   2550
      Width           =   4620
   End
   Begin VB.Menu mnuAvail 
      Caption         =   "Available"
      Visible         =   0   'False
      Begin VB.Menu miAddReq 
         Caption         =   "&Add to Required"
      End
      Begin VB.Menu miAddProhibit 
         Caption         =   "&Add to Prohibited"
      End
      Begin VB.Menu miSEP1 
         Caption         =   "-"
      End
      Begin VB.Menu miCANC1 
         Caption         =   "Cancel Menu"
      End
   End
   Begin VB.Menu mnuAllowed 
      Caption         =   "Allowed"
      Visible         =   0   'False
      Begin VB.Menu miRemAllowed 
         Caption         =   "&Remove from Allowed"
      End
      Begin VB.Menu miMoveProb 
         Caption         =   "&Move to Prohibited"
      End
      Begin VB.Menu miSEP2 
         Caption         =   "-"
      End
      Begin VB.Menu miCANC2 
         Caption         =   "Cancel Menu"
      End
   End
   Begin VB.Menu mnuProhibited 
      Caption         =   "Prohibited"
      Visible         =   0   'False
      Begin VB.Menu miRemProhib 
         Caption         =   "&Remove from Prohibited"
      End
      Begin VB.Menu miMoveAllow 
         Caption         =   "&Move to Allowed"
      End
      Begin VB.Menu miSEP3 
         Caption         =   "-"
      End
      Begin VB.Menu miCANC3 
         Caption         =   "Cancel Menu"
      End
   End
End
Attribute VB_Name = "frmCheckpoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function ParseFlags()
flagstr = ""
For Approve = 0 To lstRequired.ListCount - 1
    flagstr = flagstr & lstRequired.List(Approve)
Next
flagstr = flagstr & "-"
For Prohibit = 0 To lstProhibit.ListCount - 1
    flagstr = flagstr & lstProhibit.List(Prohibit)
Next
ParseFlags = flagstr
End Function

Sub reenum()
lstCheckpoints.Clear
Dim lbEnum As New Collection
Call CheckPoint(cpEnumerate, , , , , lbEnum)
For i = 1 To lbEnum.Count
    lstCheckpoints.AddItem lbEnum(i)
Next i

End Sub

Private Sub cmdAdd_Click()
lstProhibit.Clear
lstRequired.Clear
lstFlags.Clear
Dim lbEnum As New Collection
Call AccessFlag(afEnumerate, , , , , lbEnum)
For i = 1 To lbEnum.Count
    lstFlags.AddItem lbEnum(i)
Next i
txtCheckpoint.SetFocus
txtCheckpoint.Text = ""
cmdCancel.Visible = True
picHide.Visible = False
cmdAdd.Enabled = False
cmdRemove.Enabled = False
cmdEdit.Caption = "Save"
cmdEdit.Tag = "NEW"
End Sub

Private Sub cmdCancel_Click()
picHide.Visible = True
lstCheckpoints.Enabled = True
cmdAdd.Enabled = True
cmdRemove.Enabled = True
cmdEdit.Tag = ""
cmdEdit.Caption = "Edit"
cmdCancel.Visible = False
End Sub

Private Sub cmdEdit_Click()
If cmdEdit.Caption = "Save" Then
    If cmdEdit.Tag = "NEW" Then
        If CheckPoint(cpAddCheckpoint, txtCheckpoint, ParseFlags) = True Then
            MsgBox "Entry added."
            Call reenum
            picHide.Visible = True
            cmdAdd.Enabled = True
            cmdRemove.Enabled = True
            cmdCancel.Visible = False
            cmdEdit.Caption = "Edit"
            cmdEdit.Tag = ""
            Exit Sub
        Else
            MsgBox "Due to an error, the entry was not added. Please retry, or click cancel."
            Exit Sub
        End If
    ElseIf cmdEdit.Tag = "EDIT" Then
        If CheckPoint(cpEditCheckpoint, lstCheckpoints.Text, , txtCheckpoint, ParseFlags) = True Then
            Call reenum
            MsgBox "Entry updated."
            picHide.Visible = True
            cmdAdd.Enabled = True
            lstCheckpoints.Enabled = True
            cmdCancel.Visible = False
            cmdRemove.Enabled = True
            cmdEdit.Caption = "Edit"
            cmdEdit.Tag = ""
            Exit Sub
        End If
    End If
ElseIf cmdEdit.Caption = "Edit" Then
    cmdCancel.Visible = True
    cmdEdit.Caption = "Save"
    cmdEdit.Tag = "EDIT"
    cmdCancel.Visible = False
    picHide.Visible = False
    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
    txtCheckpoint = lstCheckpoints.Text
    tmpV$ = CheckPoint(cpGetFlags, lstCheckpoints.Text)
    tallowed = ParseCPFlags(ptRequired, tmpV)
    tpro = ParseCPFlags(ptProhibit, tmpV)
    lstRequired.Clear
    lstProhibit.Clear
    lstFlags.Clear
    stAK = ""
    For i = 1 To Len(tallowed)
        lstRequired.AddItem Mid(tallowed, i, 1)
        stAK = stAK & Mid(tallowed, i, 1)
    Next
    For i = 1 To Len(tpro)
        lstProhibit.AddItem Mid(tpro, i, 1)
        stAK = stAK & Mid(tpro, i, 1)
    Next
    End If
    Dim strNE As New Collection
    Call AccessFlag(afEnumerate, , , , , strNE)
    For i = 1 To strNE.Count
        If InStr(stAK, strNE.Item(i)) = 0 Then lstFlags.AddItem strNE.Item(i)
    Next i
    txtCheckpoint.SetFocus
    txtCheckpoint.Text = lstCheckpoints.Text
    lstCheckpoints.Enabled = False
End Sub

Private Sub cmdRemove_Click()
Call CheckPoint(cpRemCheckpoint, lstCheckpoints.Text)
Call reenum
MsgBox "Checkpoint deleted."

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call reenum
End Sub

Private Sub lstCheckpoints_Click()
tmpV$ = CheckPoint(cpGetFlags, lstCheckpoints.Text)
tallowed = ParseCPFlags(ptRequired, tmpV)
tpro = ParseCPFlags(ptProhibit, tmpV)
txtContents = "Required: " & tallowed & vbCrLf & "Prohibited: " & tpro
End Sub

Private Sub lstFlags_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstFlags.ListIndex <> -1 Then If Button = 2 Then Call PopupMenu(mnuAvail)
End Sub

Private Sub lstProhibit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstProhibit.ListIndex <> -1 Then If Button = 2 Then Call PopupMenu(mnuProhibited)
End Sub

Private Sub lstRequired_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstRequired.ListIndex <> -1 Then If Button = 2 Then Call PopupMenu(mnuAllowed)
End Sub

Private Sub miAddProhibit_Click()
lstProhibit.AddItem lstFlags.Text
lstFlags.RemoveItem lstFlags.ListIndex
End Sub

Private Sub miAddReq_Click()
lstRequired.AddItem lstFlags.Text
lstFlags.RemoveItem lstFlags.ListIndex
End Sub

Private Sub miMoveAllow_Click()
lstAllow.AddItem lstProhibit.Text
lstProhibit.RemoveItem lstProhibit.ListIndex
End Sub

Private Sub miMoveProb_Click()
lstProhibit.AddItem lstRequired.Text
lstRequired.RemoveItem lstRequired.ListIndex
End Sub

Private Sub miRemAllowed_Click()
lstFlags.AddItem lstRequired.Text
lstRequired.RemoveItem lstRequired.ListIndex
End Sub

Private Sub miRemProhib_Click()
lstFlags.AddItem lstProhibit.Text
lstProhibit.RemoveItem lstProhibit.ListIndex

End Sub
