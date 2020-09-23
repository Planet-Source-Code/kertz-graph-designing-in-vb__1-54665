VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmMain 
   Caption         =   "Project Graph 1.0"
   ClientHeight    =   4905
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6855
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Show Stacking"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   0
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show Legend"
      Height          =   255
      Left            =   5520
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   5040
      ScaleHeight     =   330
      ScaleWidth      =   1575
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmMain.frx":0442
      Left            =   0
      List            =   "frmMain.frx":046A
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4575
      Left            =   0
      OleObjectBlob   =   "frmMain.frx":0508
      TabIndex        =   1
      Top             =   360
      Width           =   6855
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1080
      Top             =   1305
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1305
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29D8
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AEA
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BFC
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D0E
            Key             =   "Copy"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Import Database..."
         Shortcut        =   ^I
      End
      Begin VB.Menu export 
         Caption         =   "&Export As Bitmap..."
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu copy 
         Caption         =   "Copy to Clipboard"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
'setting options
If Check1.Value = 1 Then
frmMain.MSChart1.ShowLegend = True
End If
If Check1.Value = 0 Then
frmMain.MSChart1.ShowLegend = False
End If
End Sub

Private Sub Check2_Click()
'setting options
If Check2.Value = 1 Then
frmMain.MSChart1.Stacking = True
End If
If Check2.Value = 0 Then
frmMain.MSChart1.Stacking = False
End If
End Sub

Private Sub Combo1_Click()
'setting chart type
If Combo1.Text = Combo1.List(0) Then
    MSChart1.ChartType = VtChChartType2dArea
End If
If Combo1.Text = Combo1.List(1) Then
    MSChart1.ChartType = VtChChartType2dBar
End If
If Combo1.Text = Combo1.List(2) Then
    MSChart1.ChartType = VtChChartType2dCombination
End If
If Combo1.Text = Combo1.List(3) Then
    MSChart1.ChartType = VtChChartType2dLine
End If
If Combo1.Text = Combo1.List(4) Then
    MSChart1.ChartType = VtChChartType2dPie
End If
If Combo1.Text = Combo1.List(5) Then
    MSChart1.ChartType = VtChChartType2dStep
End If
If Combo1.Text = Combo1.List(6) Then
    MSChart1.ChartType = VtChChartType2dXY
End If
If Combo1.Text = Combo1.List(7) Then
    MSChart1.ChartType = VtChChartType3dArea
End If
If Combo1.Text = Combo1.List(8) Then
    MSChart1.ChartType = VtChChartType3dBar
End If
If Combo1.Text = Combo1.List(9) Then
    MSChart1.ChartType = VtChChartType3dCombination
End If
If Combo1.Text = Combo1.List(10) Then
    MSChart1.ChartType = VtChChartType3dLine
End If
If Combo1.Text = Combo1.List(11) Then
    MSChart1.ChartType = VtChChartType3dStep
End If
End Sub

Private Sub copy_Click()
MSChart1.EditCopy
MsgBox "Chart Image has been copied to clipboard", vbInformation
End Sub
Private Sub export_Click()
With dlgCommonDialog
        .DialogTitle = "Export Chat to..."
        .CancelError = False
        .Filter = "Bitmap Image (*.bmp)|*.bmp"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    '
    Dim strSaveFile As String
    '
    strSaveFile = sFile
    'copying graph to clipboard and then saving...
    MSChart1.EditCopy
    SavePicture Clipboard.GetData, strSaveFile
End Sub

Private Sub Form_Load()
'setting initial graph style
Combo1.Text = Combo1.List(0)
End Sub

Private Sub Form_Resize()
  'resizing chart with form
    With MSChart1
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
    
End Sub

Private Sub mnuFileExit_Click()
    'end the program
    End

End Sub

Private Sub mnuFilePrint_Click()
 On Error GoTo errh:
    Me.PrintForm
errh:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical
        Exit Sub
    End If
End Sub

Private Sub mnuFileOpen_Click()
    On Error GoTo errh:
    Dim sFile, sFilex, tablex As String
    sFilex = sFile
       
    With dlgCommonDialog
        .DialogTitle = "Import Database..."
        .CancelError = False
        .Filter = "Access Database (*.mdb)|*.mdb"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    'prompt to input table name
    tablex = InputBox("Enter table name")
    'reading table
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    '
    Dim strProvider As String
    Dim strDataSource As String
    Dim strSQL As String
    '
    strProvider = "Microsoft.Jet.OLEDB.4.0"
    strDataSource = sFile
    strSQL = "SELECT * FROM " & tablex
    '
    cnn.Open "provider=" & strProvider & ";Data Source=" & strDataSource
    rst.Open strSQL, cnn, adOpenStatic
    'setting database to chart
    With MSChart1
        Set .DataSource = rst
    End With
errh:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical
        Exit Sub
    End If
End Sub


