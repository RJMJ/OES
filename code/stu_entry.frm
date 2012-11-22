VERSION 5.00
Begin VB.Form frm_stuentry 
   Caption         =   "Student Entry"
   ClientHeight    =   4785
   ClientLeft      =   4665
   ClientTop       =   3150
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "stu_entry.frx":0000
   ScaleHeight     =   4785
   ScaleWidth      =   6195
   Begin VB.CommandButton cmdchange 
      Caption         =   "Change"
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2520
      Width           =   1095
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
      Left            =   2475
      Picture         =   "stu_entry.frx":BA696
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "PREVIOUS"
      ToolTipText     =   "Move To Previous"
      Top             =   3495
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
      Left            =   3315
      Picture         =   "stu_entry.frx":BA74C
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "NEXT"
      ToolTipText     =   "Move To Next"
      Top             =   3495
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
      Left            =   1635
      Picture         =   "stu_entry.frx":BA802
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "FIRST"
      ToolTipText     =   "Move To First"
      Top             =   3495
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
      Index           =   8
      Left            =   4155
      Picture         =   "stu_entry.frx":BA8EC
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "LAST"
      ToolTipText     =   "Move To Last"
      Top             =   3495
      Width           =   800
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
      Left            =   1080
      Picture         =   "stu_entry.frx":BA9D6
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "ADD"
      ToolTipText     =   "Add New Record"
      Top             =   3855
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
      Left            =   3600
      Picture         =   "stu_entry.frx":BAEB4
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "EXIT"
      ToolTipText     =   "Exit Form"
      Top             =   3855
      Width           =   795
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
      Left            =   2760
      Picture         =   "stu_entry.frx":BB37A
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "CANCEL"
      ToolTipText     =   "Search Record"
      Top             =   3855
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
      Left            =   1920
      Picture         =   "stu_entry.frx":BB434
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "SAVE"
      ToolTipText     =   "Save Record"
      Top             =   3855
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "NAVIGATE"
      ToolTipText     =   "Move To First"
      Top             =   3495
      Width           =   1110
   End
   Begin VB.ComboBox cmbstream 
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
      Height          =   315
      ItemData        =   "stu_entry.frx":BB526
      Left            =   1680
      List            =   "stu_entry.frx":BB528
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox txtname 
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
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox txtroll 
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
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
   Begin VB.ComboBox cmbclass 
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
      Height          =   315
      ItemData        =   "stu_entry.frx":BB52A
      Left            =   1680
      List            =   "stu_entry.frx":BB534
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   1485
      Left            =   4560
      Picture         =   "stu_entry.frx":BB540
      Top             =   120
      Width           =   1560
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   480
      X2              =   4920
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student name :"
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
      TabIndex        =   19
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Roll No.:"
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
      TabIndex        =   18
      Top             =   840
      Width           =   765
   End
   Begin VB.Label Label5 
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
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student Entry..."
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
      Left            =   240
      TabIndex        =   16
      Top             =   240
      Width           =   2325
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stream"
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
      TabIndex        =   15
      Top             =   2280
      Width           =   600
   End
End
Attribute VB_Name = "frm_stuentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset
Dim rec1 As ADODB.Recordset
Private Sub databind()
Set txtroll.DataSource = rec
txtroll.DataField = "roll"

Set txtname.DataSource = rec
txtname.DataField = "sname"

Set cmbclass.DataSource = rec
cmbclass.DataField = "class"

Set cmbstream.DataSource = rec
cmbstream.DataField = "stream"

End Sub



Private Sub cmdchange_Click()
txtname.Enabled = True
txtroll.Enabled = True
cmbclass.Enabled = True
cmbstream.Enabled = True
CmdChoice(1).Enabled = True
CmdChoice(2).Enabled = True
txtname.locked = False
        txtroll.locked = False
        cmbclass.locked = False
        cmbstream.locked = False
txtroll.SetFocus
End Sub

Private Sub CmdChoice_Click(Index As Integer)
Select Case Index
Case 4: Unload Me
Case 0: rec.AddNew
        databind
        txtname.Enabled = True
        txtroll.Enabled = True
        cmbclass.Enabled = True
        cmbstream.Enabled = True
        txtname.locked = False
        txtroll.locked = False
        cmbclass.locked = False
        cmbstream.locked = False
        CmdChoice(0).Enabled = False
        CmdChoice(1).Enabled = True
        CmdChoice(2).Enabled = True
        For i = 5 To 9
        CmdChoice(i).Enabled = False
        Next i
        cmddelete.Enabled = False
        cmdchange.Enabled = False
Case 1: rec.Update
        txtname.locked = True
        txtroll.locked = True
        cmbclass.locked = True
        cmbstream.locked = True
        CmdChoice(0).Enabled = True
        CmdChoice(2).Enabled = False
        CmdChoice(9).Enabled = True
        CmdChoice(1).Enabled = False
        rec.MoveFirst
Case 2: rec.CancelUpdate
        CmdChoice(0).Enabled = True
        CmdChoice(9).Enabled = True
        CmdChoice(1).Enabled = False
        CmdChoice(2).Enabled = False
Case 4: Unload Me
Case 5: rec.MoveFirst
        
Case 6:
        If Not rec.AbsolutePosition = 1 Then
        rec.MovePrevious
        Else: rec.MoveLast
        End If
       
Case 7:
        If Not rec.AbsolutePosition = rec.RecordCount Then
        rec.MoveNext
        Else: rec.MoveFirst
        End If
        
Case 8: rec.MoveLast
        
Case 9: databind
        CmdChoice(1).Enabled = False
        For i = 5 To 8
        CmdChoice(i).Enabled = True
        Next i
        CmdChoice(9).Enabled = False
        cmddelete.Enabled = True
        cmdchange.Enabled = True
        txtname.Enabled = True
        txtroll.Enabled = True
        cmbclass.Enabled = True
        cmbstream.Enabled = True
End Select
End Sub

Private Sub cmddelete_Click()
rec.Delete
rec.MoveFirst
End Sub

Private Sub Form_Load()
Connection1
Set rec = New ADODB.Recordset
rec.CursorType = adOpenDynamic
rec.LockType = adLockOptimistic
rec.Open "select * from student order by roll", cn, , , adCmdText


Set rec1 = New ADODB.Recordset
'rec1.CursorType = adOpenDynamic
'rec.LockType = adLockOptimistic
rec1.Open "select distinct stream from subject", cn, , , adCmdText

rec1.MoveFirst
Do While Not rec1.EOF
cmbstream.AddItem (rec1!stream)
rec1.MoveNext
Loop

End Sub
