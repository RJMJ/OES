VERSION 5.00
Begin VB.Form frm_subchoice 
   BackColor       =   &H00CFB8A8&
   Caption         =   "Subject Choice"
   ClientHeight    =   5490
   ClientLeft      =   4455
   ClientTop       =   2295
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   4515
   Begin VB.ListBox List1 
      BackColor       =   &H00CFB8A8&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      ItemData        =   "Subchoice.frx":0000
      Left            =   2160
      List            =   "Subchoice.frx":0002
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00CFB8A8&
      Caption         =   "Start"
      Height          =   495
      Left            =   1680
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00CFB8A8&
      Caption         =   "Choose the Subject"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   1740
      Left            =   120
      Picture         =   "Subchoice.frx":0004
      Top             =   1680
      Width           =   1740
   End
End
Attribute VB_Name = "frm_subchoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
On Error GoTo error_sub

If timstatus = 0 Then
timstatus = 1                   ' For Starting timer
Else: timstatus = 2
End If
qsubject = List1.Text
List1.RemoveItem (List1.ListIndex)
If List1.ListCount = 0 Then
subchoice = 1
End If
'List1.ListIndex = 0
Me.Hide
frm_qpaper.Show

' Error Handler
If Err.Number <> 0 Then
error_sub: MsgBox ("Please Select any of the subject")
End If
End Sub




Private Sub Form_Load()
Connection1
'stream = "science"
'sclass = 12
Set rec = New ADODB.Recordset
rec.CursorType = adOpenDynamic
rec.LockType = adLockOptimistic
'sql = "select subject from subject where stream =" & "'" & stream & "'"
sql = "select distinct subject from question where class =" & sclass & " and subject in(" & "select subject from subject where stream =" & "'" & stream & "'" & ")"
rec.Open sql, cn, , , adCmdText

rec.MoveFirst
Do While Not rec.EOF
List1.AddItem (rec!subject)
rec.MoveNext
Loop

List1.ListIndex = 0
End Sub


Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmd1_Click
End If
End Sub
