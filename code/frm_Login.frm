VERSION 5.00
Begin VB.Form frm_Login 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4680
   ClientLeft      =   4005
   ClientTop       =   3000
   ClientWidth     =   6810
   FillColor       =   &H00C0C0FF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2765.1
   ScaleMode       =   0  'User
   ScaleWidth      =   6553.084
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   6000
      Top             =   840
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtpass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtuser 
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UserName"
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
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   1080
      Shape           =   2  'Oval
      Top             =   1080
      Visible         =   0   'False
      Width           =   4575
   End
End
Attribute VB_Name = "frm_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()


If LCase(txtuser) = "admin" And LCase(txtpass) = "admin" Then
Unload Me: MDIForm1.Show
Else
    rec.MoveFirst
    Do While Not rec.EOF
        If rec!roll = Val(txtpass.Text) - 100 And LCase(rec!sname) = LCase(txtuser.Text) Then
        z = 1
        Exit Do
        End If
        rec.MoveNext
    Loop
    If z = 1 Then
    roll = rec!roll
    stream = rec!stream
    sname = rec!sname
    'If rec!Class = "XI" Then sclass = 11
    'Else: sclass = "XII"
    'End If
    sclass = rec!Class
    Unload Me: frm_intro.Show
    Else: MsgBox ("Invalid UserName And Password")
    txtuser.SetFocus
    End If
End If
End Sub
Private Sub Form_Load()
Connection1
Set rec = cn.Execute("student")
z = 0
Me.Picture = LoadPicture(App.Path + "\splash.jpg")
End Sub


Private Sub Timer1_Timer()
Shape1.Visible = True
lbl1.Visible = True
lbl2.Visible = True
txtuser.Visible = True
txtpass.Visible = True
cmdOK.Visible = True
cmdcancel.Visible = True
Timer1.Enabled = False
End Sub

Private Sub txtpass_GotFocus()
'SendKeys "{HOME}+{END}"
End Sub



Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdOK_Click
End If
End Sub



'Private Sub txtuser_Change()
'txtuser.Text = UCase(Mid(txtuser.Text, 1, 1)) & Mid(txtuser.Text, 2)
'txtuser.SelStart = Len(txtuser.Text)
'End Sub

Private Sub txtuser_GotFocus()
'SendKeys "{HOME}+{END}"
End Sub

'End Sub
Private Sub txtuser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtpass.SetFocus
End If
End Sub

Private Sub txtuser_LostFocus()
txtuser.Text = UCase(Mid(txtuser.Text, 1, 1)) & Mid(txtuser.Text, 2)
End Sub
