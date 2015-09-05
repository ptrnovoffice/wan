Attribute VB_Name = "Koneksi"
Public Conn As ADODB.Connection
Public Rs_Scdl As New ADODB.Recordset
Public StrCon As String

Public Sub ConMain()
On Error Resume Next
    Set Conn = New ADODB.Connection
    
    '==== MySQLSERVER Eka Laptop ===================
    'StrCon = "DRIVER={MySQL ODBC 5.1 Driver};" _
        & "Persist Security Info=False;" _
        & "SERVER=localhost;UID=root;port=3306;" _
        & "PWD=Sp1d3rm4n4;DATABASE=attpayroll;" _
        & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 163841
      
    '==== WANINDO CLIENT ===================
    StrCon = "DRIVER={MySQL ODBC 5.1 Driver};" _
          & "Persist Security Info=False;" _
          & "SERVER=server_absen;UID=root;port=3306;" _
          & "PWD=Sp1d3rm4n4;DATABASE=attpayroll;" _
          & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 163841
          
    Conn.ConnectionString = StrCon
    Conn.Open
    If Conn.State = 1 Then
        'MsgBox "Connection Database Berhasil"
    Else
        'MsgBox "Maaf Aplikasi tidak dapat terhubung dengan Server...", vbExclamation
    End If
End Sub

'---------------------------------------------
'---------- BUKA RECORD KONDISI --------------
'--------------- BY PTR.NOV-------------------
Sub OpRecScdl(qry, isProcedure As Boolean)
Set Rs_Scdl = New ADODB.Recordset
Rs_Scdl.CursorLocation = adUseClient
    If isProcedure Then
        Rs_Scdl.Open qry, Conn, adOpenStatic, adLockOptimistic, adCmdStoredProc
      Else
        Rs_Scdl.Open qry, Conn, adOpenStatic, adLockOptimistic, adCmdText
    End If
End Sub
