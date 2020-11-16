VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Viewer & Tools"
   ClientHeight    =   4155
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   6840
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   3840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11360
            MinWidth        =   11360
            Text            =   "Database not loaded"
            TextSave        =   "Database not loaded"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "11:25 a.m."
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "Database"
      Begin VB.Menu smnuLoad 
         Caption         =   "Load database"
      End
      Begin VB.Menu smnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuViewer 
      Caption         =   "Viewer ..."
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnuViewer_Click()
    If StatusBar.Panels(1).Text = "Database not loaded" Then
        MsgBox StatusBar.Panels(1).Text, vbCritical
    Else
        Load frmViewer
        frmViewer.Show , frmMain
    End If
End Sub

Private Sub smnuClose_Click()
    End
End Sub

Private Sub smnuLoad_Click()
    cd.Filter = "Microsoft database | *.mdb| All files |*.*"
    cd.Flags = cdlOFNFileMustExist
    cd.ShowOpen
    LoadDatabase (cd.FileName)
End Sub
