VERSION 5.00
Begin VB.Form formrpt 
   Caption         =   "Result"
   ClientHeight    =   3120
   ClientLeft      =   5280
   ClientTop       =   3750
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "formrpt.frx":0000
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   Begin VB.CommandButton Command2 
      BackColor       =   &H00CFB8A8&
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00CFB8A8&
      Caption         =   "Preview"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00CFB8A8&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00CFB8A8&
      Caption         =   "Select Student Roll No. Below:-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00CFB8A8&
      Caption         =   "Roll No :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "formrpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()
'rec.MoveFirst
'Do While Not rec.EOF
'If rec!sname = Combo1.Text Then
'Exit Do
'Else: rec.MoveNext
'End If
'Loop
'rptroll = rec!roll
rptroll = Combo1.Text
End Sub

Private Sub Command1_Click()
If de1.rsrscmd3.State = 1 Then
de1.rsrscmd3.Close
End If
If rptflag = 2 Then
de1.rsrscmd3.Open "select * from totmarks where roll =" & rptroll, cn
rpttotresult.Show
ElseIf rptflag = 1 Then
de1.rsrscmd4.Open "select * from marks where roll =" & rptroll, cn
rptsubresult.Show
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Connection1

Set rec = New ADODB.Recordset
rec.CursorType = adOpenDynamic
rec.LockType = adLockOptimistic
rec.Open "select distinct roll from marks order by roll", cn, , , adCmdText

Do While Not rec.EOF
Combo1.AddItem (rec!roll)
rec.MoveNext
Loop
'Combo1.ListIndex = 0
End Sub
