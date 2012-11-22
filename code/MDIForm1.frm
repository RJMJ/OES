VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Online Examination"
   ClientHeight    =   8220
   ClientLeft      =   690
   ClientTop       =   -2115
   ClientWidth     =   11880
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   11820
      TabIndex        =   2
      Top             =   7845
      Width           =   11880
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      Height          =   9960
      Left            =   0
      ScaleHeight     =   9900
      ScaleWidth      =   11820
      TabIndex        =   1
      Top             =   495
      Width           =   11880
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      Picture         =   "MDIForm1.frx":0442
      ScaleHeight     =   435
      ScaleWidth      =   11820
      TabIndex        =   0
      Top             =   0
      Width           =   11880
   End
   Begin VB.Menu stu_master 
      Caption         =   "&Student_Master"
      Begin VB.Menu stuentry 
         Caption         =   "Student &Entry"
      End
      Begin VB.Menu stushow 
         Caption         =   "&Show Students"
      End
   End
   Begin VB.Menu que_entry 
      Caption         =   "&Question_Master"
      Begin VB.Menu queentry 
         Caption         =   "Question &Entry"
      End
      Begin VB.Menu queshow 
         Caption         =   "&Show Quesions"
      End
   End
   Begin VB.Menu subentry 
      Caption         =   "S&ubject Entry"
   End
   Begin VB.Menu result 
      Caption         =   "&Results"
      Begin VB.Menu resultsub 
         Caption         =   "&Subject Wise"
      End
      Begin VB.Menu resulttot 
         Caption         =   "&Overall Result"
      End
   End
   Begin VB.Menu rep 
      Caption         =   "Repor&ts"
      Visible         =   0   'False
      Begin VB.Menu rptresult 
         Caption         =   "&Results"
         Begin VB.Menu subrep 
            Caption         =   "&Subject Wise"
         End
         Begin VB.Menu totrep 
            Caption         =   "&Overall Result"
         End
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu rtpstu 
         Caption         =   "&Students List"
      End
      Begin VB.Menu rtpque 
         Caption         =   "&Questions List"
      End
   End
   Begin VB.Menu about 
      Caption         =   "&About Us"
   End
   Begin VB.Menu exit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
frm_about.Show
End Sub

Private Sub exit_Click()
If MsgBox("Would you like to close this session....", vbCritical + vbYesNo + vbDefaultButton2, "Exit") = vbYes Then
End
End If
End Sub


Private Sub MDIForm_Load()
MDIForm1.BackColor = &H80000005
Picture2.Picture = LoadPicture(App.Path & "\main.jpg")
End Sub



Private Sub queentry_Click()
frm_QueEntry.Show
End Sub

Private Sub queshow_Click()
frm_queshow.Show
End Sub

Private Sub resultsub_Click()
resulttype = 1
frm_mainresult.Show
End Sub

Private Sub resulttot_Click()
resulttype = 2
frm_mainresult.Show
End Sub

Private Sub rtpque_Click()
rptque.Show
End Sub

Private Sub rtpstu_Click()
rptstu.Show
End Sub

Private Sub stuentry_Click()
frm_stuentry.Show
End Sub

Private Sub stushow_Click()
frm_stushow.Show
End Sub

Private Sub subentry_Click()
frm_subentry.Show
End Sub

Private Sub subrep_Click()
rptflag = 1
formrpt.Show
End Sub

Private Sub totrep_Click()
rptflag = 2
formrpt.Show
End Sub
