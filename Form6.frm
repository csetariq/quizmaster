VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00404040&
   Caption         =   "Admin Control Panel"
   ClientHeight    =   8145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8910
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Change Password"
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
      Left            =   1920
      TabIndex        =   8
      Top             =   6960
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Toppers List"
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
      Left            =   4800
      TabIndex        =   7
      Top             =   6960
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Log out"
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
      Left            =   120
      TabIndex        =   6
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
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
      Left            =   7080
      TabIndex        =   5
      Top             =   6960
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time limit for each Question"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   390
      Left            =   1905
      TabIndex        =   2
      Top             =   4080
      Width           =   3870
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No of Questions per player"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   390
      Left            =   2145
      TabIndex        =   1
      Top             =   3000
      Width           =   3630
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Control Panel"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   660
      Left            =   3045
      TabIndex        =   0
      Top             =   960
      Width           =   3045
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ftch As String
Dim wrt As String
Private Sub Command1_Click()
        
    'Getting Questions per Player,Time Interval and Addition Factor
    qperp = Text1.Text
    tint = Text2.Text
    Call getafa
    
    'Updating Questions per Player,Time Interval and Addition Factor
    wrt = "update QuizConfig set Q_per_P = " & qperp & ",T_Int=" & tint & ",A_Fact=" & afa & " where Tot_Q = " & totq & ""
    rs.Open wrt, cn, adOpenDynamic, adLockOptimistic
    MsgBox "Updated Successfully", vbInformation + vbOKOnly, "Quiz Master says.."
           
End Sub

Private Sub Command2_Click()
    Unload Me
    Form1.Show
End Sub

Private Sub Command3_Click()
    Unload Me
    Form5.Show
End Sub

Private Sub Command4_Click()
    Unload Me
    Form8.Show
End Sub

Private Sub Form_Load()
            
    'Showing the previous Configuration
    ftch = "select * from QuizConfig"
    rs.Open ftch, cn, adOpenDynamic, adLockOptimistic
    Text1.Text = rs.Fields("Q_per_P")
    Text2.Text = rs.Fields("T_Int")
    rs.Close
        
End Sub

