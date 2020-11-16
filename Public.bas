Attribute VB_Name = "Public"
Public RunExport As Boolean
Public Db As Database
Public Rs As Recordset
Public Table As TableDef
Public Field As Field

Public Sub DisplayTables()
    frmViewer.listTables.Clear
    For Each Table In Db.TableDefs
        If Not Table.Name Like "MSys*" Then
            frmViewer.listTables.AddItem (Table.Name)
        End If
    Next Table
'Count tables
    frmViewer.lbl(0).Caption = "Tables"
    frmViewer.lbl(0).Caption = frmViewer.lbl(0).Caption & " (" & frmViewer.listTables.ListCount & ")"
End Sub

Public Sub DisplayFields()
Dim tIndex As Integer
    frmViewer.listFields.Clear
    For Each Table In Db.TableDefs
        If Table.Name Like frmViewer.listTables.Text Then
            For Each Field In Db.TableDefs(tIndex).Fields
            frmViewer.listFields.AddItem Field.Name
    Next Field
    End If
tIndex = tIndex + 1
Next Table
'Count tables
    frmViewer.lbl(1).Caption = "Fields"
    frmViewer.lbl(1).Caption = frmViewer.lbl(1).Caption & " (" & frmViewer.listFields.ListCount & ")"
End Sub

Public Sub LoadDatabase(NameDb As String)
'This is just for verify that file is a database
On Error GoTo myErr:
    Set Db = OpenDatabase(NameDb)
    frmMain.StatusBar.Panels(1).Text = NameDb
Exit Sub
myErr:
    If Err.Number >= 1 Then
        MsgBox Err.Description, vbCritical, Err.Number
    End If
End Sub


Public Sub ExportExcel()
If frmViewer.Grid.TextMatrix(0, 1) = vbNullString Then
    MsgBox "Zero records to export", vbCritical
Else
    On Error GoTo myErr
    RunExport = True
    TotalRegs = frmViewer.Grid.Rows
    Load frmExport
    frmExport.prBar.Min = 0
    frmExport.prBar.Max = TotalRegs
    frmExport.Show , frmViewer
    Dim AppExcel As Variant, myStr As String
    Set AppExcel = CreateObject("Excel.application")
    AppExcel.Visible = False
    AppExcel.Workbooks.Add
    For r = 0 To frmViewer.Grid.Rows - 1
        DoEvents
        If RunExport Then
        myCount = myCount + 1
        frmExport.lbl.Caption = "Exporting " & myCount & " of " & TotalRegs
        frmExport.prBar.Value = myCount
        frmExport.Refresh
        For c = 0 To frmViewer.Grid.Cols - 1
            myStr = frmViewer.Grid.TextMatrix(r, c)
            AppExcel.cells(r + 2, c + 1).Formula = Trim(myStr)
        Next c
    End If
    Next r
    Unload frmExport
    AppExcel.Visible = True
    Set AppExcel = Nothing
End If
Exit Sub
myErr:
If Err.Number >= 1 Then
    MsgBox Err.Description, vbCritical, Err.Number
End If
End Sub

