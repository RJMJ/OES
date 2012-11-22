VERSION 5.00
Begin VB.Form frm_qresult 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Result"
   ClientHeight    =   5265
   ClientLeft      =   4035
   ClientTop       =   3435
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   5130
   Begin VB.PictureBox cd1 
      Height          =   480
      Left            =   4080
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   24
      Top             =   480
      Width           =   1200
   End
   Begin VB.ComboBox cmbsub 
      Height          =   315
      Left            =   3240
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdexit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
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
      Left            =   3000
      TabIndex        =   2
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label lblclass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3240
      TabIndex        =   23
      Top             =   1680
      Width           =   75
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   22
      Top             =   1680
      Width           =   465
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   21
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Result Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   1800
      TabIndex        =   20
      Top             =   120
      Width           =   1470
   End
   Begin VB.Label Lblper 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2760
      TabIndex        =   19
      Top             =   3840
      Width           =   555
   End
   Begin VB.Label lblname 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3240
      TabIndex        =   18
      Top             =   960
      Width           =   45
   End
   Begin VB.Label lblqattempt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3240
      TabIndex        =   17
      Top             =   2760
      Width           =   45
   End
   Begin VB.Label lbltotqright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3240
      TabIndex        =   16
      Top             =   3120
      Width           =   45
   End
   Begin VB.Label lbltotq 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3240
      TabIndex        =   15
      Top             =   2400
      Width           =   45
   End
   Begin VB.Label lblroll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3240
      TabIndex        =   14
      Top             =   1320
      Width           =   45
   End
   Begin VB.Label lbltotmarks 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3240
      TabIndex        =   13
      Top             =   3480
      Width           =   45
   End
   Begin VB.Label lblgrd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   12
      Top             =   4200
      Width           =   75
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   11
      Top             =   3840
      Width           =   990
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   10
      Top             =   4200
      Width           =   525
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Question Attempted"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Correct Answers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Questions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Marks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Roll Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1845
      Left            =   1440
      Picture         =   "Qresult.frx":0000
      Top             =   360
      Width           =   2085
   End
End
Attribute VB_Name = "frm_qresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim grade As String
Dim totq As Integer
Dim qattempt As Integer
Dim marks As Integer
Dim per As Integer
Dim grade1 As String
Dim sql As String
Dim adocmd As New ADODB.Command

Private Sub cmbsub_Change()
rec.Move (cmbsub.ListIndex + 1)
End Sub

Private Sub cmbsub_Click()
rec.MoveFirst
rec.Move (cmbsub.ListIndex)
calc
End Sub

Private Sub cmdexit_Click()
For i = 0 To cmbsub.ListCount - 1
cmbsub.ListIndex = i

totq = totq + lbltotq.Caption
qattempt = qattempt + Val(lblqattempt.Caption)
totqright = totqright + Val(lbltotqright.Caption)
marks = marks + Val(lbltotmarks.Caption)
per = Lblper.Caption
grade1 = lblgrd.Caption

Next i
adocmd.ActiveConnection = cn
sql = "Insert Into totmarks values"
sql = sql & "" & "(" & "" & "" & roll & "" & ","
sql = sql & "" & "" & totq & "" & "," & "" & "" & qattempt & "" & ","
sql = sql & "" & "" & totqright & "" & "," & "" & "" & marks & "" '& ","
'sql = sql & "" & "" & per & "" & "," & "" & "'" & grade1 & "'" & ""
sql = sql & ")"
adocmd.CommandText = sql
adocmd.Execute

cmbsub.ListIndex = 0
MsgBox ("Thank You!!!!!")
Unload Me
End
End Sub

Private Sub cmdprint_Click()
'On Error GoTo error_print
'frm_qresult.PrintForm
'cd1.ShowPrinter
'Exit Sub
'error_print: If MsgBox("Sorry, Printer Error", vbExclamation, "Error:") = vbOK Then
 '               End If

End Sub

Private Sub Form_Load()
'For i = 1 To 21999999
'Next i
Connection1
'roll = 1
'sname = ("dheeraj")
'roll1 = roll
Set rec = New ADODB.Recordset
rec.CursorType = adOpenDynamic
rec.LockType = adLockOptimistic
str = "select * from marks where roll =" & roll
rec.Open str, cn, , , adCmdText

rec.MoveFirst
Do While Not rec.EOF
cmbsub.AddItem (rec!subject)        ' Entering Subject in Combo
rec.MoveNext
Loop

rec.MoveFirst
lblname.Caption = UCase(sname)
lblroll.Caption = roll
If sclass = 12 Then
lblclass.Caption = "XII"
ElseIf sclass = 11 Then
lblclass.Caption = "XI"
End If
cmbsub.ListIndex = 0
calc
End Sub
Private Sub calc()
lbltotq.Caption = rec!totq
lblqattempt.Caption = rec!totqattempt
lbltotqright.Caption = rec!totqright
lbltotmarks.Caption = rec!marks
Lblper.Caption = Round((lbltotmarks.Caption * 100) / (lbltotq.Caption * 2), 2)
Select Case Int(Lblper.Caption)
Case 80 To 100
    grade = "A"
Case 60 To 80
    grade = "B"
Case 40 To 60
    grade = "C"
Case Is < 40
    grade = "Fail"
End Select
lblgrd.Caption = grade
'Set rec = Nothing
End Sub

