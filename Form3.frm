VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "On the Game"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1366
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3360
      Top             =   2520
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   8880
      Width           =   1815
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00808080&
      Caption         =   "Option4"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6960
      Width           =   5775
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00808080&
      Caption         =   "Option3"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6960
      Width           =   5775
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00808080&
      Caption         =   "Option2"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   5775
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Next"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14160
      TabIndex        =   6
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1335
      Left            =   15840
      TabIndex        =   9
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1335
      Left            =   3360
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quiz Master Pro 1.0"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   975
      Left            =   4800
      TabIndex        =   0
      Top             =   840
      Width           =   10215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Retrieve the question here"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1635
      Left            =   3480
      TabIndex        =   1
      Top             =   3240
      Width           =   13005
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sec As Integer
Dim cnt As Integer
Dim qptr As Integer
Dim aans As Integer
Dim uans As Integer
Dim ftch As String
Dim Ins As String
Dim para As String

Private Sub Timer1_Timer()
        'Decrementing the Counter
        sec = sec - 1
        Label4.Caption = sec
        'Pass on when the time limit is reached
        If sec = 0 Then
            Call Command1_Click
        End If
End Sub

Private Sub hash()
    qptr = ((qptr + afa) Mod totq) + 1
       
End Sub

Private Sub getpara()
    'Getting the Configuration
    para = "select * from QuizConfig"
    rs.Open para, cn, adOpenDynamic, adLockOptimistic
    totq = rs.Fields("Tot_Q")
    qperp = rs.Fields("Q_per_P")
    tint = rs.Fields("T_Int")
    afa = rs.Fields("A_Fact")
    rs.Close
End Sub
Private Sub fetchQ()
    If cnt = qperp Then Command1.Caption = "&Finish"
        
        Call hash
        ftch = "Select * from Quest where Sno=" & qptr & ""
        rs.Open ftch, cn, adOpenDynamic, adLockBatchOptimistic
             
            'Updating the Question
            Label1.Caption = rs.Fields("Que")
            Option1.Caption = rs.Fields("Opt1")
            Option2.Caption = rs.Fields("Opt2")
            Option3.Caption = rs.Fields("Opt3")
            Option4.Caption = rs.Fields("Opt4")
            Label3.Caption = cnt
            aans = rs.Fields("Ans")
        
        sec = tint
        Label4.Caption = sec
        cnt = cnt + 1
        uans = 0
        rs.Close
        
End Sub

Private Sub Command1_Click()
    
    'Getting the user answer
    If Option1.Value = True Then
        uans = 1
        Option1.Value = False
    End If
    
    If Option2.Value = True Then
        uans = 2
        Option2.Value = False
    End If
    
    If Option3.Value = True Then
        uans = 3
        Option3.Value = False
    End If
    
    If Option4.Value = True Then
        uans = 4
        Option4.Value = False
    End If
    
    'Comparing the user answer with actual answer
    If uans = aans Then Score = Score + 1
    
    
    If Not cnt = qperp + 1 Then
        Call fetchQ
    
    Else
        
        'Inserting the user information
        Ins = "insert into Userlog(ID,S_Name,S_Course,S_Dept,S_Year,College,Score) values('" & id & "','" & P_Name & "','" & P_Course & "','" & P_Dept & "','" & P_Year & "','" & P_College & "'," & Score & ")"
        rs.Open Ins, cn, adOpenDynamic, adLockOptimistic
        Form4.Show
        
    End If
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()
    Call getpara
    Randomize
    'Generating the Initial random pointer
    qptr = Fix(totq * Rnd)
    cnt = 1
    Option1.Value = False
    Option2.Value = False
    Option3.Value = False
    Option4.Value = False
    Call fetchQ
End Sub
