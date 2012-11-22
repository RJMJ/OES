VERSION 5.00
Begin VB.Form frm_QueEntry 
   Caption         =   "Question Entry"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Que_Entry.frx":0000
   ScaleHeight     =   8295
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Que_Entry.frx":13E696
      Left            =   6720
      List            =   "Que_Entry.frx":13E6A0
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   680
      Width           =   735
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00CFB8A8&
      Caption         =   "Option4"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   6600
      Width           =   855
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00CFB8A8&
      Caption         =   "Option3"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   6600
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00CFB8A8&
      Caption         =   "Option2"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   6600
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00CFB8A8&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   6600
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CFB8A8&
      Height          =   7095
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Width           =   10095
      Begin VB.CommandButton cmddel 
         Caption         =   "Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         Picture         =   "Que_Entry.frx":13E6AC
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   6480
         Width           =   1455
      End
      Begin VB.CommandButton cmdchange 
         Caption         =   "Change"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Picture         =   "Que_Entry.frx":13E74E
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   6480
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   1920
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   3720
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4800
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   3720
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   3720
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CommandButton CmdChoice 
         Caption         =   "&New"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   0
         Left            =   3120
         Picture         =   "Que_Entry.frx":13E9F1
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "ADD"
         ToolTipText     =   "Add New Record"
         Top             =   6495
         Width           =   800
      End
      Begin VB.CommandButton CmdChoice 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   4
         Left            =   5640
         Picture         =   "Que_Entry.frx":13EECF
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "EXIT"
         ToolTipText     =   "Exit Form"
         Top             =   6495
         Width           =   800
      End
      Begin VB.CommandButton CmdChoice 
         Caption         =   "&Cancel"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   2
         Left            =   4800
         Picture         =   "Que_Entry.frx":13F395
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "CANCEL"
         ToolTipText     =   "Search Record"
         Top             =   6495
         Width           =   800
      End
      Begin VB.CommandButton CmdChoice 
         Caption         =   "&Save"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   1
         Left            =   3960
         Picture         =   "Que_Entry.frx":13F44F
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "SAVE"
         ToolTipText     =   "Save Record"
         Top             =   6495
         Width           =   800
      End
      Begin VB.CommandButton CmdChoice 
         Caption         =   "Na&vigate"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "NAVIGATE"
         ToolTipText     =   "Move To First"
         Top             =   6135
         Width           =   1110
      End
      Begin VB.CommandButton CmdChoice 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   6195
         Picture         =   "Que_Entry.frx":13F541
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "LAST"
         ToolTipText     =   "Move To Last"
         Top             =   6135
         Width           =   800
      End
      Begin VB.CommandButton CmdChoice 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   3675
         Picture         =   "Que_Entry.frx":13F62B
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "FIRST"
         ToolTipText     =   "Move To First"
         Top             =   6135
         Width           =   800
      End
      Begin VB.CommandButton CmdChoice 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   5355
         Picture         =   "Que_Entry.frx":13F715
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "NEXT"
         ToolTipText     =   "Move To Next"
         Top             =   6135
         Width           =   800
      End
      Begin VB.CommandButton CmdChoice 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   4515
         Picture         =   "Que_Entry.frx":13F7CB
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "PREVIOUS"
         ToolTipText     =   "Move To Previous"
         Top             =   6135
         Width           =   800
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00CFB8A8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   22
         Top             =   5160
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Image Image1 
         Height          =   1545
         Left            =   8280
         Picture         =   "Que_Entry.frx":13F881
         Top             =   135
         Width           =   1785
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Option 4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7080
         TabIndex        =   30
         Top             =   3360
         Width           =   1065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Option 3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4800
         TabIndex        =   29
         Top             =   3360
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Option 2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         TabIndex        =   28
         Top             =   3360
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Option 1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   27
         Top             =   3360
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Question No :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Question  :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   1320
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   3120
         X2              =   6480
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2520
         X2              =   6960
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Label Label3 
         BackColor       =   &H00CFB8A8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   4680
         Width           =   2775
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Correct Answer -->>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   23
         Top             =   4680
         Width           =   2550
      End
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   680
      Width           =   2295
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class:"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5640
      TabIndex        =   32
      Top             =   600
      Width           =   840
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Questions Entry"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   3720
      TabIndex        =   31
      Top             =   0
      Width           =   2325
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   600
      TabIndex        =   20
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frm_QueEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim rec1 As ADODB.Recordset
Dim flagadd As Integer
Dim flagdel As Integer

Private Sub cmdchange_Click()
Call enable(True)
Call locked(False)
Combo1.SetFocus
For i = 5 To 9
CmdChoice(i).Enabled = False
Next i

CmdChoice(1).Enabled = True
CmdChoice(2).Enabled = True
CmdChoice(4).Enabled = True
End Sub

Private Sub CmdChoice_Click(Index As Integer)
Select Case Index
Case 4: Unload Me
Case 0: rec.AddNew
        flagadd = 1
        Call enable(True)
        Call locked(False)
        For i = 5 To 8
        CmdChoice(i).Enabled = False
        Next i
        CmdChoice(2).Enabled = True
        CmdChoice(9).Enabled = False
        CmdChoice(1).Enabled = True
         CmdChoice(0).Enabled = False
         cmdchange.Enabled = False
         cmddel.Enabled = False
         databind
Case 1:
        If (Combo2.Text = "" Or Combo1.Text = "") Then
        MsgBox ("Please Select both class and subject")
        Exit Sub
        End If
        rec.Update
        
        Call enable(False)
        Call locked(True)
        CmdChoice(9).Enabled = True
         CmdChoice(1).Enabled = False
          CmdChoice(0).Enabled = True
         If flagadd = 1 Then
         flagadd = 0
         End If
         CmdChoice(2).Enabled = False
Case 2: 'Call locked(False)
        'CmdChoice(1).Enabled = True
        ' CmdChoice(0).Enabled = False
        ' CmdChoice(9).Enabled = False
        ' CmdChoice(2).Enabled = False
        '  For i = 5 To 8
       ' CmdChoice(i).Enabled = False
       ' Next i
        rec.CancelUpdate
        If flagadd = 1 Then
         flagadd = 0
         End If
         Call enable(False)
        Call locked(True)
        For i = 5 To 8
        CmdChoice(i).Enabled = False
        Next i
        CmdChoice(2).Enabled = False
        CmdChoice(9).Enabled = True
        CmdChoice(1).Enabled = False
         CmdChoice(0).Enabled = True
       
Case 5: rec.MoveFirst
        GoTo lab
Case 6:
        If Not rec.AbsolutePosition = 1 Then
        rec.MovePrevious
        Else: rec.MoveLast
        End If
        GoTo lab
Case 7:
        If Not rec.AbsolutePosition = rec.RecordCount Then
        rec.MoveNext
        Else: rec.MoveFirst
        End If
        GoTo lab
Case 8: rec.MoveLast
        GoTo lab
Case 9: For i = 5 To 8
        CmdChoice(i).Enabled = True
        Next i
        CmdChoice(2).Enabled = True
        Call enable(True)
        Combo1.locked = False
        Combo2.locked = False
        cmdchange.Enabled = True
        cmddel.Enabled = True
End Select
lab: If Label3.Caption = "a" Then
Option1.Value = True
ElseIf Label3.Caption = "b" Then
Option2.Value = True
ElseIf Label3.Caption = "c" Then
Option3.Value = True
Else: Option4.Value = True
End If

End Sub

Private Sub cmddel_Click()
On Error Resume Next
flagdel = 1
rec.Delete
Call enable(False)
Call locked(True)
MsgBox ("Successfully Deleted")
Combo1.Text = ""
Combo2.Text = ""
Text1.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""

End Sub

Private Sub Combo1_Click()
On Error GoTo error_query
If flagadd = 0 Then
Set rec = New ADODB.Recordset
rec.CursorType = adOpenDynamic
rec.LockType = adLockOptimistic
sql = "select * from question where subject like '" & Combo1.Text & "' and class = " & Combo2.Text & " order by queid"
rec.Open sql, cn, , , adCmdText
If rec.RecordCount = 0 Then
MsgBox ("Sorry,No Questions Found")
Combo1.Text = ""
 For i = 5 To 8
        CmdChoice(i).Enabled = False
        Next i
       
Exit Sub
End If
databind
End If
Exit Sub
error_query: MsgBox ("Please Select both class and subject")

End Sub



Private Sub Combo2_Click()
If flagadd = 0 Then
Set rec = New ADODB.Recordset
rec.CursorType = adOpenDynamic
rec.LockType = adLockOptimistic
sql = "select * from question where subject like '" & Combo1.Text & "' and class = " & Combo2.Text & " order by queid"
rec.Open sql, cn, , , adCmdText
If rec.RecordCount = 0 Then
MsgBox ("Sorry,No Questions Found")
Combo1.Text = ""
 For i = 5 To 8
        CmdChoice(i).Enabled = False
        Next i
       
Exit Sub
End If
databind
End If
End Sub

Private Sub Form_Load()
Connection1
Set rec = New ADODB.Recordset
rec.CursorType = adOpenDynamic
rec.LockType = adLockOptimistic
rec.Open "select * from question order by queid", cn, , , adCmdText

Set rec1 = New ADODB.Recordset
rec1.CursorType = adOpenDynamic
rec1.LockType = adLockOptimistic
sql = "select subject from subject"
rec1.Open sql, cn, , , adCmdText

rec1.MoveFirst
Do While Not rec1.EOF
Combo1.AddItem (rec1!subject)
rec1.MoveNext
Loop





End Sub





Private Sub Option1_Click()
rec!answer = "a"
End Sub

Private Sub Option2_Click()
rec!answer = "b"
End Sub

Private Sub Option3_Click()
rec!answer = "c"
End Sub

Private Sub Option4_Click()
rec!answer = "d"
End Sub



Public Sub databind()
On Error Resume Next
Set Combo1.DataSource = rec
Combo1.DataField = "subject"
Set Combo2.DataSource = rec
Combo2.DataField = "Class"

Set Text1.DataSource = rec
Text1.DataField = "question"

Set Text3.DataSource = rec
Text3.DataField = "choice1"

Set Text4.DataSource = rec
Text4.DataField = "choice2"

Set Text5.DataSource = rec
Text5.DataField = "choice3"

Set Text6.DataSource = rec
Text6.DataField = "choice4"

Set Text7.DataSource = rec
Text7.DataField = "queid"

Set Label3.DataSource = rec
Label3.DataField = "answer"


If Label3.Caption = "a" Then
Option1.Value = True
ElseIf Label3.Caption = "b" Then
Option2.Value = True
ElseIf Label3.Caption = "c" Then
Option3.Value = True
Else: Option4.Value = True
End If

End Sub

Public Sub enable(a As String)

Combo1.Enabled = a
Combo2.Enabled = a
Text1.Enabled = a
Text3.Enabled = a
Text4.Enabled = a
Text5.Enabled = a
Text6.Enabled = a
Text7.Enabled = a
Option1.Enabled = a
Option2.Enabled = a
Option3.Enabled = a
Option4.Enabled = a
If flagadd = 0 Then
If flagdel = 1 Then
flagdel = 0
Exit Sub
End If
databind
End If
End Sub

Public Sub locked(a As String)
Combo1.locked = a
Combo2.locked = a
Text1.locked = a
Text3.locked = a
Text4.locked = a
Text5.locked = a
Text6.locked = a
Text7.locked = a

End Sub



Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii < 47 Or KeyAscii >= 55 Then
KeyAscii = 0
End If
End Sub
