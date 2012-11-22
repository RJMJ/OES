VERSION 5.00
Begin VB.Form frm_qpaper 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Question Paper"
   ClientHeight    =   8325
   ClientLeft      =   1920
   ClientTop       =   1335
   ClientWidth     =   10515
   FillColor       =   &H00C0FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Qpaper.frx":0000
   ScaleHeight     =   8325
   ScaleWidth      =   10515
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Tag             =   "a"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox txtc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Index           =   3
      Left            =   2400
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   5880
      Width           =   6615
   End
   Begin VB.TextBox txtc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Index           =   2
      Left            =   2400
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   5040
      Width           =   6615
   End
   Begin VB.TextBox txtc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   4200
      Width           =   6615
   End
   Begin VB.TextBox txtc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Index           =   0
      Left            =   2400
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3360
      Width           =   6615
   End
   Begin VB.TextBox lblquestion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1560
      Width           =   6615
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   9000
      Top             =   1080
   End
   Begin VB.CommandButton cmdfinishexam 
      BackColor       =   &H00808080&
      Cancel          =   -1  'True
      Caption         =   "Finished Module"
      Height          =   435
      Left            =   4920
      MaskColor       =   &H00800080&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H00808080&
      Caption         =   "&Next"
      Default         =   -1  'True
      Height          =   435
      Left            =   3360
      MaskColor       =   &H00800080&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   3
      Tag             =   "c"
      Top             =   5160
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   4
      Tag             =   "d"
      Top             =   6000
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Tag             =   "b"
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   720
      TabIndex        =   24
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label lbltotmark 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   465
      Left            =   9960
      TabIndex        =   23
      Top             =   1080
      Width           =   465
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Question Paper"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   360
      TabIndex        =   22
      Top             =   0
      Width           =   2070
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   390
      Left            =   5760
      TabIndex        =   21
      Top             =   120
      Width           =   90
   End
   Begin VB.Label lbltotqremain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   7320
      TabIndex        =   20
      Top             =   720
      Width           =   45
   End
   Begin VB.Label lbltotq 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   5040
      TabIndex        =   19
      Top             =   720
      Width           =   45
   End
   Begin VB.Label lblremaintime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   3000
      TabIndex        =   18
      Top             =   720
      Width           =   45
   End
   Begin VB.Label lbltottime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   1080
      TabIndex        =   17
      Top             =   720
      Width           =   45
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining Questions :"
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Left            =   5640
      TabIndex        =   16
      Top             =   720
      Width           =   1590
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Questions :"
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Left            =   3720
      TabIndex        =   15
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time remaining :"
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Left            =   1800
      TabIndex        =   14
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Time :"
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   840
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9480
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   480
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Width           =   465
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marks "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   9480
      TabIndex        =   11
      Top             =   600
      Width           =   930
   End
   Begin VB.Image Image1 
      Height          =   7080
      Left            =   0
      Picture         =   "Qpaper.frx":0342
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   10920
   End
End
Attribute VB_Name = "frm_qpaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim flag As Integer
Dim totqright As Integer
Dim totqattempt As Integer
Dim roll1 As Integer
Dim adocmd As New ADODB.Command


Private Sub cmdfinishexam_Click()


If flag = 1 Then        'giving marks
totqright = totqright + 1
flag = 0
End If

For i = 0 To 3
If Option1(i).Value = True Then
totqattempt = totqattempt + 1
Option1(i).Value = False    ' For Clearing option boxes
End If
Next i
totq = lbltotq.Caption
totmarks = totqright * 2
adocmd.ActiveConnection = cn
'adocmd.CommandText = "insert into marks values(roll," & "'" & qsubject & "'" & ",totq,totqattempt,totqright,totmarks)"

sql = "Insert Into marks values"
sql = sql & "" & "(" & "" & "" & roll & "" & "," & "" & "'" & qsubject & "'" & ","
sql = sql & "" & "" & totq & "" & "," & "" & "" & totqattempt & "" & ","
sql = sql & "" & "" & totqright & "" & "," & "" & "" & totmarks & ""
sql = sql & ")"
'"insert into marks values(1," & "'" & qsubject & "'" & ", & " '" & totq & "'" & ",1,1,1)"
adocmd.CommandText = sql
'adocmd.CommandType = adCmdText
'adocmd.Execute ("insert into marks values(1,'dd',1,1,1,1)"), cn, adCmdText
adocmd.Execute
Timer1.Enabled = False

If subchoice = 1 Then
MsgBox ("You Have Finished Exam!!!!")
Unload Me: frm_qresult.Show
Else
frm_subchoice.List1.ListIndex = 0
Unload Me: frm_subchoice.Show

End If
End Sub

Private Sub cmdnext_Click()
If flag = 1 Then        'giving marks
totqright = totqright + 1
flag = 0
End If
lbltotmark.Caption = totqright
For i = 0 To 3
If Option1(i).Value = True Then
totqattempt = totqattempt + 1
Option1(i).Value = False    ' For Clearing option boxes
End If
Next i


If Not rec.AbsolutePosition = rec.RecordCount - 1 Then
rec.MoveNext
Label8.Caption = Label8.Caption + 1
lbltotmark.Caption = totqright
Else:
rec.MoveNext
cmdnext.Enabled = False
Label8.Caption = Label8.Caption + 1
lbltotmark.Caption = totqright
End If

lbltotqremain.Caption = lbltotqremain.Caption - 1
End Sub



Private Sub Form_Load()
Connection1
roll1 = roll
'qsubject = "vb"
'stream = "science"
totqright = 0
totmarks = 0
totqattempt = 0
totq = 0



Set rec = New ADODB.Recordset
rec.CursorType = adOpenDynamic
rec.LockType = adLockOptimistic
str = "select * from question where subject like " & "'" & qsubject & "' and class = " & sclass
rec.Open str, cn, , , adCmdText

Set lblquestion.DataSource = rec
lblquestion.DataField = "question"

Set txtc(0).DataSource = rec
txtc(0).DataField = "choice1"

Set txtc(1).DataSource = rec
txtc(1).DataField = "choice2"

Set txtc(2).DataSource = rec
txtc(2).DataField = "choice3"

Set txtc(3).DataSource = rec
txtc(3).DataField = "choice4"

lbltotq.Caption = rec.RecordCount
lbltotqremain.Caption = lbltotq.Caption - 1

For i = 0 To 3
Option1(i).Value = False        ' For Clearing option boxes
Next i
lbltotmark.Caption = totqright
lbldate.Caption = Format(Now(), "dd-mmmm-yyyy")
lbltottime.Caption = "30 min"
If timstatus = 1 Then
tim = 30
End If
Timer1.Enabled = True
lblremaintime.Caption = (tim) & " min"

If rec.AbsolutePosition = rec.RecordCount Then
cmdnext.Enabled = False
End If
End Sub





Private Sub Option1_Click(Index As Integer)

If rec!answer = Option1(Index).Tag Then       'checking answer
flag = 1
Else: flag = 0
End If
End Sub

Private Sub Timer1_Timer()
If tim = 0 Then
Timer1.Enabled = False
MsgBox ("Time Over !!!!!")
subchoice = 1
cmdfinishexam_Click
Else: tim = tim - 1
lblremaintime.Caption = tim & " min"
End If

End Sub



Private Sub txtc_Click(Index As Integer)
Option1(Index).SetFocus
End Sub
