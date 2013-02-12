Attribute VB_Name = "Module1"
Global id As String
Global P_Name As String
Global P_Course As String
Global P_Dept As String
Global P_Year As String
Global P_College As String
Global Score As Integer
Global totq As Integer
Global qperp As Integer
Global tint As Integer
Global afa As Integer
Global ptotq As Integer
Dim ftchALL As String
Dim ftchNo As String
Global conn As String

Global cn  As New ADODB.Connection
Global rs  As New ADODB.Recordset
Public Sub getafa()
    afa = (totq / qperp) - 1
End Sub
Public Sub updateatstart()
    ftchALL = "select * from Quest"
    ftchNo = "select * from QuizConfig"
    
    'getting previous total questions and question per player
    rs.Open ftchNo, cn, adOpenDynamic, adLockOptimistic
    ptotq = rs.Fields("Tot_Q")
    qperp = rs.Fields("Q_per_P")
    tint = rs.Fields("T_Int")
    rs.Close
    
    'Updatating actual no of total questions and addition factor
    rs.Open ftchALL, cn, adOpenDynamic, adLockOptimistic
    rs.MoveLast
    totq = rs.Fields("Sno")
    Call getafa
    rs.Close
    
    'Updating the QuizConfig table
    rs.Open "update QuizConfig set Tot_Q = " & totq & ",A_Fact=" & afa & " where Tot_Q = " & ptotq & "", cn, adOpenDynamic, adLockOptimistic

End Sub
Sub main()
    conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Quiz Engine.mdb;Persist Security Info=False;Jet OLEDB:Database Password=amoM5Pc6"
    cn.Open conn
    Form1.Show
End Sub
