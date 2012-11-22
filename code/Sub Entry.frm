VERSION 5.00
Begin VB.Form frm_subentry 
   Caption         =   "Subjects Viewer"
   ClientHeight    =   3210
   ClientLeft      =   3135
   ClientTop       =   3450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Sub Entry.frx":0000
   ScaleHeight     =   3210
   ScaleWidth      =   9450
   Begin VB.ComboBox txtstream 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   3855
   End
   Begin VB.TextBox txtstream1 
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
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox txtsub 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1440
      Width           =   3855
   End
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
      Left            =   7080
      TabIndex        =   11
      Top             =   840
      Width           =   1095
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
      Left            =   3000
      Picture         =   "Sub Entry.frx":8BF7A
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "ADD"
      ToolTipText     =   "Add New Record"
      Top             =   2535
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
      Left            =   5520
      Picture         =   "Sub Entry.frx":8C458
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "EXIT"
      ToolTipText     =   "Exit Form"
      Top             =   2535
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
      Left            =   4680
      Picture         =   "Sub Entry.frx":8C91E
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "CANCEL"
      ToolTipText     =   "Search Record"
      Top             =   2535
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
      Left            =   3840
      Picture         =   "Sub Entry.frx":8C9D8
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "SAVE"
      ToolTipText     =   "Save Record"
      Top             =   2535
      Width           =   800
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
      Left            =   7080
      TabIndex        =   12
      Top             =   1440
      Width           =   1095
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "NAVIGATE"
      ToolTipText     =   "Move To First"
      Top             =   2175
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
      Left            =   6075
      Picture         =   "Sub Entry.frx":8CACA
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "LAST"
      ToolTipText     =   "Move To Last"
      Top             =   2175
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
      Left            =   3555
      Picture         =   "Sub Entry.frx":8CBB4
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "FIRST"
      ToolTipText     =   "Move To First"
      Top             =   2175
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
      Left            =   5235
      Picture         =   "Sub Entry.frx":8CC9E
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "NEXT"
      ToolTipText     =   "Move To Next"
      Top             =   2175
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
      Left            =   4395
      Picture         =   "Sub Entry.frx":8CD54
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "PREVIOUS"
      ToolTipText     =   "Move To Previous"
      Top             =   2175
      Width           =   800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "STREAM:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1560
      TabIndex        =   16
      Top             =   840
      Width           =   1515
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Entry :"
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
      TabIndex        =   15
      Top             =   120
      Width           =   2145
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SUBJECT NAME :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      TabIndex        =   14
      Top             =   1440
      Width           =   2475
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2400
      X2              =   6840
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2400
      X2              =   6840
      Y1              =   2520
      Y2              =   2520
   End
End
Attribute VB_Name = "frm_subentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As ADODB.Recordset
Private Sub cmdchange_Click()
txtsub.locked = False
txtstream.locked = False
CmdChoice(1).Enabled = True
CmdChoice(2).Enabled = True
End Sub

Private Sub CmdChoice_Click(Index As Integer)
Select Case Index
Case 0: rec.AddNew
        databind
        CmdChoice(0).Enabled = False
        CmdChoice(1).Enabled = True
        CmdChoice(2).Enabled = True
        For i = 5 To 9
        CmdChoice(i).Enabled = False
        Next i
        cmddelete.Enabled = False
        cmdchange.Enabled = False
        txtsub.Enabled = True
        txtstream.Enabled = True
        txtsub.locked = False
        txtstream.locked = False
Case 1: rec.Update
        CmdChoice(0).Enabled = True
        CmdChoice(2).Enabled = False
        CmdChoice(9).Enabled = True
        txtsub.locked = True
        txtstream.locked = True
        CmdChoice(1).Enabled = False
Case 2: rec.CancelUpdate
        txtstream.Enabled = False
        txtsub.Enabled = False
        CmdChoice(0).Enabled = True
        CmdChoice(9).Enabled = True
        CmdChoice(1).Enabled = False
        CmdChoice(2).Enabled = False
        
Case 4: Unload Me

Case 9: databind
        CmdChoice(1).Enabled = False
        For i = 5 To 8
        CmdChoice(i).Enabled = True
        Next i
        CmdChoice(9).Enabled = False
        cmddelete.Enabled = True
        cmdchange.Enabled = True
        txtsub.Enabled = True
        txtstream.Enabled = True
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
rec.Open "select * from subject", cn, , , adCmdText

Set rec1 = New ADODB.Recordset
rec1.CursorType = adOpenDynamic
rec1.LockType = adLockOptimistic
rec1.Open "select distinct stream from subject", cn, , , adCmdText

Do While Not rec1.EOF
txtstream.AddItem (rec1!stream)
rec1.MoveNext
Loop

End Sub

Public Sub databind()
Set txtstream.DataSource = rec
txtstream.DataField = "stream"

Set txtsub.DataSource = rec
txtsub.DataField = "subject"
End Sub

