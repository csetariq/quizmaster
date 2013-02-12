VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login to Proceed"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "#"
      TabIndex        =   3
      ToolTipText     =   "Use - '12345'"
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Use your Roll No. not Reg No."
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " &OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Must be Unique)"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Password"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Username"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ftchID As String
Dim ftchUSR As String
Private Sub sherr()
    MsgBox "Invalid Username or Password", vbExclamation + vbOKOnly, "Quiz Master Says.."
End Sub

Private Sub Command1_Click()
    If rs.State = 1 Then rs.Close
        rs.Open ftchID, cn, adOpenDynamic, adLockOptimistic
    If Not Text1.Text = rs.Fields("Admin_ID") Then
        If rs.State = 1 Then rs.Close
        rs.Open "select ID,S_Name from Userlog where ID = '" & Text1.Text & "'", cn, adOpenDynamic, adLockOptimistic
            If rs.EOF = True Then
                If Text2.Text = "12345" Then
                
                    rs.Close
                    id = Text1.Text
                    Unload Me
                    Form2.Show
                Else
                    Call sherr
                End If
            Else
                MsgBox "You've already taken the test", vbInformation + vbOKOnly, "Quiz Master Says.."
            End If
            
    
    Else
        
        If Text2.Text = rs.Fields("Admin_PWD") Then
            rs.Close
            Unload Me
            Form6.Show
        Else
            Call sherr
        End If
        
        
    End If
           
End Sub

Private Sub Command2_Click()
    If rs.State = True Then rs.Close
    End
End Sub

Private Sub Form_Load()
    Call updateatstart
    ftchID = "select * from Admin"
    'ftchUSR = "select ID,S_Name from Userlog where ID = '" & Text1.Text & "'"
    'Opening the Table 'Admin' to verify username and password
    
End Sub
