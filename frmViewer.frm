VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmViewer 
   Caption         =   "Viewer: "
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   9315
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data getData 
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Bindings        =   "frmViewer.frx":0000
      Height          =   3495
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6165
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.ListBox listFields 
      Height          =   2010
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.ListBox listTables 
      Height          =   2010
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lbl 
      Caption         =   "Records"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   4815
   End
   Begin VB.Label lbl 
      Caption         =   "Fields"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Tables"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Menu mnuClose 
      Caption         =   "Close"
   End
   Begin VB.Menu nmuTools 
      Caption         =   "Tools"
      Begin VB.Menu smnuExport 
         Caption         =   "Export to Excel ..."
      End
      Begin VB.Menu smnuSQL 
         Caption         =   "SQL"
      End
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.Caption = "Viewer: "
    Me.Caption = Me.Caption & frmMain.cd.FileTitle
    Me.getData.DatabaseName = frmMain.cd.FileName
    DisplayTables
End Sub

Private Sub listFields_DblClick()
Dim strSQL As String
    strSQL = "select all(" & Trim(listFields.Text) & ") from " & Trim(listTables.Text)
    getData.RecordSource = strSQL
    getData.Refresh
'Count records
    lbl(2).Caption = "Records"
    lbl(2).Caption = lbl(2).Caption & " (" & listTables.Text & " - " & listFields.Text & " - " & Grid.Rows - 1 & ")"
End Sub

Private Sub listTables_Click()
    DisplayFields
End Sub

Private Sub listTables_DblClick()
    getData.RecordSource = listTables.Text
    getData.Refresh
'Count records
    lbl(2).Caption = "Records"
    lbl(2).Caption = lbl(2).Caption & " (" & listTables.Text & " - " & Grid.Rows - 1 & ")"
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub smnuExport_Click()
    ExportExcel
End Sub

Private Sub smnuSQL_Click()
    Load frmSQL
    frmSQL.Show vbModal
End Sub
