Attribute VB_Name = "modutil"
Public Sub getKoneksi()
    conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "/bridgingvclaim.mdb;Persist Security Info=False"
    Set dbconn = New ADODB.Connection
    If dbconn.State = adStateOpen Then Exit Sub
    dbconn.CursorLocation = adUseClient
    dbconn.ConnectionString = conString
    dbconn.Open
    If dbconn.State = adStateClosed Then
        Call MsgBox("Koneksi ke database Bermasalah")
        Exit Sub
    End If
End Sub
Public Sub msubrec(rsQ As ADODB.Recordset, query As String)
    Set rsQ = New ADODB.Recordset
    Call getKoneksi
    rsQ.Open query, dbconn, adOpenStatic, adLockReadOnly
End Sub
Public Sub msubdcSource(dcName As DataCombo, rsQ As ADODB.Recordset, query As String)
    Set rsQ = New ADODB.Recordset
    Set rsQ = dbconn.Execute(query)
    Set dcName.RowSource = rsQ
    dcName.BoundColumn = rsQ(0).Name
    dcName.ListField = rsQ(1).Name
    Exit Sub
End Sub
Public Sub centerForm(ByRef oForm1 As Form, ByVal oForm2 As Form)
    oForm1.Left = (oForm2.Width - oForm1.Width) / 2
    oForm1.Top = (oForm2.Height - 1500 - oForm1.Height) / 2
End Sub
