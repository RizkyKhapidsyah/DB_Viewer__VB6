VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export to Excel"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar prBar 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   6360
      Y1              =   700
      Y2              =   700
   End
   Begin VB.Line Line1 
      DrawMode        =   16  'Merge Pen
      X1              =   240
      X2              =   6360
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lbl 
      Caption         =   "Exporting 0 of 0"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub


Private Sub cmdCancel_Click()
If (MsgBox("Cancel operation", vbCritical + vbYesNo)) = vbYes Then
    RunExport = False
End If
End Sub


Private Sub Form_Load()

End Sub
