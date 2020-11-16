VERSION 5.00
Begin VB.Form frmSQL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   6360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Execute"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtSQL 
      Height          =   525
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "frmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdExecute_Click()
If txtSQL.Text <> vbNullString Then
    On Error GoTo myErr
    frmViewer.getData.RecordSource = txtSQL
    frmViewer.getData.Refresh
    frmViewer.lbl(2).Caption = "Records"
    frmViewer.lbl(2).Caption = frmViewer.lbl(2).Caption & " " & "(" & "SQL" & " - " & frmViewer.Grid.Rows - 1 & ")"
End If
Exit Sub
myErr:
If Err.Number >= 1 Then
    MsgBox Err.Description, vbCritical, Err.Number
End If
End Sub
