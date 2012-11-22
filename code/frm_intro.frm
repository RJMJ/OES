VERSION 5.00
Begin VB.Form frm_intro 
   Caption         =   "Instructions"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   Picture         =   "frm_intro.frx":0000
   ScaleHeight     =   7290
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00CFB8A8&
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00CFB8A8&
      Caption         =   "Start Exam"
      Default         =   -1  'True
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1665
      Left            =   7320
      Picture         =   "frm_intro.frx":13E696
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "General Instructions"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   390
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   3180
   End
   Begin VB.Label lblgreetings 
      BackColor       =   &H00CFB8A8&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label lblinstructions 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFB8A8&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   7515
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frm_intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                          'General declaration
Dim mFileSysObj As New FileSystemObject  'General declaration
Dim mFile As File                        'General declaration
Dim mTxtStream As TextStream             'General declaration
Dim instructions As String

Private Sub Command1_Click()
Unload Me
frm_subchoice.Show
End Sub

Private Sub Command2_Click()
If MsgBox("Do You really want to Quit the Exam", vbYesNo + vbCritical) = vbYes Then
    Unload Me
    End If

End Sub

Private Sub Form_Load()
Set mFile = mFileSysObj.GetFile(App.Path + "\Instructions.txt")

 'Open a text stream for reading to the file
 Set mTxtStream = mFile.OpenAsTextStream(ForReading)
   
 'Read the data
 instructions = mTxtStream.ReadAll
 

 

 lblgreetings.Caption = "Welcome " & sname & "!"
                                                        
 'Place only the String portion representing the
 'name in the TextBox.
 lblinstructions.Caption = instructions
 
End Sub

