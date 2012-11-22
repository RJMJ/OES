VERSION 5.00
Begin VB.Form frm_mainresult 
   Caption         =   "Result"
   ClientHeight    =   5700
   ClientLeft      =   4050
   ClientTop       =   2850
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "MainResult.frx":0000
   ScaleHeight     =   6504.993
   ScaleMode       =   0  'User
   ScaleWidth      =   6705
   Begin VB.PictureBox cd1 
      Height          =   480
      Left            =   5040
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   25
      Top             =   480
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CFB8A8&
      Caption         =   "Details"
      Height          =   2775
      Left            =   360
      TabIndex        =   15
      Top             =   2520
      Width           =   3855
      Begin VB.ComboBox cmbsub 
         BackColor       =   &H00CFB8A8&
         Height          =   315
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Lblper 
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
         Left            =   2880
         TabIndex        =   7
         Top             =   2040
         Width           =   75
      End
      Begin VB.Label lblqattempt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2880
         TabIndex        =   4
         Top             =   960
         Width           =   45
      End
      Begin VB.Label lbltotqright 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2880
         TabIndex        =   5
         Top             =   1320
         Width           =   45
      End
      Begin VB.Label lbltotq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2880
         TabIndex        =   3
         Top             =   600
         Width           =   45
      End
      Begin VB.Label lbltotmarks 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2880
         TabIndex        =   6
         Top             =   1680
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
         Left            =   2880
         TabIndex        =   8
         Top             =   2400
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
         Left            =   240
         TabIndex        =   22
         Top             =   2040
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
         Left            =   240
         TabIndex        =   21
         Top             =   2400
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
         Left            =   240
         TabIndex        =   20
         Top             =   960
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
         Left            =   240
         TabIndex        =   19
         Top             =   1320
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
         Left            =   240
         TabIndex        =   18
         Top             =   600
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
         Left            =   240
         TabIndex        =   17
         Top             =   1680
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
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CFB8A8&
      Caption         =   "Students"
      Height          =   1695
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   4095
      Begin VB.ComboBox cmbroll 
         BackColor       =   &H00CFB8A8&
         Height          =   315
         Left            =   2280
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblclass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2280
         TabIndex        =   24
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Label Label3 
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
         TabIndex        =   23
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label lblname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2280
         TabIndex        =   1
         Top             =   840
         Width           =   45
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
         TabIndex        =   14
         Top             =   360
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
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00CFB8A8&
      Cancel          =   -1  'True
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00CFB8A8&
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   2010
      Left            =   4920
      Picture         =   "MainResult.frx":BE37A
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   465
      Left            =   840
      TabIndex        =   11
      Top             =   120
      Width           =   1125
   End
End
Attribute VB_Name = "frm_mainresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim rec2 As ADODB.Recordset
Dim rec3 As ADODB.Recordset

Private Sub cmbroll_Click()
'Set rec = Nothing
On Error GoTo error_result
cmbsub.Clear
lbltotq.Caption = ""
lblqattempt.Caption = ""
lbltotqright.Caption = ""
lbltotmarks.Caption = ""
Lblper.Caption = ""
lblgrd.Caption = ""

rec.MoveFirst
Do While Not rec.EOF
If rec!roll = cmbroll.Text Then
Exit Do
Else: rec.MoveNext
End If
Loop
cmbroll.Text = rec!roll
lblname.Caption = rec!sname
lblclass.Caption = rec!Class

If resulttype = 1 Then
    
    Set rec1 = New ADODB.Recordset
    rec1.CursorType = adOpenDynamic
    rec1.LockType = adLockOptimistic
    str = "select subject from marks where roll = " & Val(cmbroll.Text)
    rec1.Open str, cn, , , adCmdText
    rec1.MoveFirst
    
    
    Do While Not rec1.EOF
    cmbsub.AddItem (rec1!subject)
    rec1.MoveNext
    Loop

Else
    Set rec3 = New ADODB.Recordset
    rec3.CursorType = adOpenDynamic
    rec3.LockType = adLockOptimistic
    str = "select * from totmarks where roll = " & cmbroll
    rec3.Open str, cn, , , adCmdText
    
    calc
    
End If
Exit Sub
error_result: MsgBox ("No Records")

End Sub



Private Sub cmbsub_Click()

Set rec2 = New ADODB.Recordset
rec2.CursorType = adOpenDynamic
rec2.LockType = adLockOptimistic
'str = "select * from marks where roll= (select roll from student where sname like '" & "a" & "'" & ")"
str = "select * from marks where roll = " & Val(cmbroll.Text) & "and subject like '" & cmbsub.Text & "'"
rec2.Open str, cn, , , adCmdText
'rec1.MoveFirst
calc

End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
'On Error GoTo error_print
'frm_mainresult.PrintForm
'cd1.ShowPrinter
'Exit Sub
'error_print: If MsgBox("Sorry, Printer Error", vbExclamation, "Error:") = vbOK Then
 '               End If

End Sub

Private Sub Form_Load()
'resulttype = 2
If resulttype = 2 Then
Label1.Visible = False
cmbsub.Visible = False
Label2.Caption = "Overall Result"
Else: Label2.Caption = "Subject-Wise Result"
End If

Connection1
Set rec = New ADODB.Recordset
rec.CursorType = adOpenDynamic
rec.LockType = adLockOptimistic
str = "select roll,sname,class from student order by roll"
rec.Open str, cn, , , adCmdText



Do While Not rec.EOF
cmbroll.AddItem (rec!roll)
rec.MoveNext
Loop


End Sub

Private Sub calc()
If resulttype = 1 Then
lbltotq.Caption = rec2!totq
lblqattempt.Caption = rec2!totqattempt
lbltotqright.Caption = rec2!totqright
lbltotmarks.Caption = rec2!marks
ElseIf resulttype = 2 Then
    lbltotq.Caption = rec3!totq
    lblqattempt.Caption = rec3!totqattempt
    lbltotqright.Caption = rec3!totqright
    lbltotmarks.Caption = rec3!marks

End If
Lblper.Caption = Round(((Val(lbltotqright.Caption) * 100) / Val(lbltotq.Caption)), 2)
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
End Sub


