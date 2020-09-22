VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmPlot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XPlot"
   ClientHeight    =   5565
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6420
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   428
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   5085
      TabIndex        =   13
      ToolTipText     =   "Clear Drawing Area"
      Top             =   3375
      Width           =   1020
   End
   Begin MSScriptControlCtl.ScriptControl SC 
      Left            =   5085
      Top             =   2250
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.CheckBox chkConstrain 
      Height          =   195
      Left            =   6150
      TabIndex        =   9
      ToolTipText     =   "Constrain X and Y Scale"
      Top             =   405
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.TextBox txtScaleY 
      Height          =   285
      Left            =   5400
      MaxLength       =   5
      TabIndex        =   7
      Text            =   "10"
      Top             =   510
      Width           =   510
   End
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   5730
      Top             =   2265
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtScaleX 
      Height          =   285
      Left            =   5400
      MaxLength       =   5
      TabIndex        =   4
      Text            =   "10"
      Top             =   195
      Width           =   510
   End
   Begin VB.TextBox txtEquation 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   735
      TabIndex        =   2
      Text            =   "Sin(x)"
      ToolTipText     =   "Write Your Function Here."
      Top             =   4920
      Width           =   3990
   End
   Begin VB.CommandButton cmdPlot 
      Caption         =   "Plot"
      Default         =   -1  'True
      Height          =   360
      Left            =   5085
      TabIndex        =   1
      Top             =   4920
      Width           =   990
   End
   Begin VB.PictureBox DrawArea 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   210
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   0
      Top             =   210
      Width           =   4515
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   18
         X1              =   0
         X2              =   301
         Y1              =   285
         Y2              =   285
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   17
         X1              =   0
         X2              =   301
         Y1              =   270
         Y2              =   270
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   16
         X1              =   0
         X2              =   301
         Y1              =   255
         Y2              =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   15
         X1              =   0
         X2              =   301
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   14
         X1              =   0
         X2              =   301
         Y1              =   225
         Y2              =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   13
         X1              =   0
         X2              =   301
         Y1              =   210
         Y2              =   210
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   12
         X1              =   0
         X2              =   301
         Y1              =   195
         Y2              =   195
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   11
         X1              =   0
         X2              =   301
         Y1              =   180
         Y2              =   180
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   10
         X1              =   0
         X2              =   301
         Y1              =   165
         Y2              =   165
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   9
         X1              =   0
         X2              =   301
         Y1              =   135
         Y2              =   135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   8
         X1              =   0
         X2              =   301
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   7
         X1              =   0
         X2              =   301
         Y1              =   105
         Y2              =   105
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   6
         X1              =   0
         X2              =   301
         Y1              =   90
         Y2              =   90
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   5
         X1              =   0
         X2              =   301
         Y1              =   75
         Y2              =   75
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   4
         X1              =   0
         X2              =   301
         Y1              =   60
         Y2              =   60
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   3
         X1              =   0
         X2              =   301
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   2
         X1              =   0
         X2              =   301
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   1
         X1              =   0
         X2              =   301
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   18
         X1              =   90
         X2              =   90
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   17
         X1              =   75
         X2              =   75
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   16
         X1              =   60
         X2              =   60
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   15
         X1              =   45
         X2              =   45
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   14
         X1              =   30
         X2              =   30
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   13
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   12
         X1              =   285
         X2              =   285
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   11
         X1              =   270
         X2              =   270
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   10
         X1              =   255
         X2              =   255
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   9
         X1              =   240
         X2              =   240
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   8
         X1              =   225
         X2              =   225
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   7
         X1              =   210
         X2              =   210
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   6
         X1              =   195
         X2              =   195
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   5
         X1              =   180
         X2              =   180
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   4
         X1              =   165
         X2              =   165
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   3
         X1              =   135
         X2              =   135
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   2
         X1              =   120
         X2              =   120
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Index           =   1
         X1              =   105
         X2              =   105
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         DrawMode        =   9  'Not Mask Pen
         Index           =   0
         X1              =   150
         X2              =   150
         Y1              =   0
         Y2              =   301
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         DrawMode        =   9  'Not Mask Pen
         Index           =   0
         X1              =   0
         X2              =   301
         Y1              =   150
         Y2              =   150
      End
   End
   Begin VB.Label lblBusy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BUSY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5295
      TabIndex        =   12
      Top             =   1035
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label PosY 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4860
      TabIndex        =   11
      Top             =   1800
      Width           =   1635
   End
   Begin VB.Label PosX 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4860
      TabIndex        =   10
      Top             =   1500
      Width           =   1635
   End
   Begin VB.Line Line4 
      Index           =   1
      X1              =   403
      X2              =   403
      Y1              =   39
      Y2              =   47
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   403
      X2              =   403
      Y1              =   19
      Y2              =   27
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   395
      X2              =   404
      Y1              =   47
      Y2              =   47
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   395
      X2              =   404
      Y1              =   19
      Y2              =   19
   End
   Begin VB.Label Label3 
      Caption         =   "§"
      Height          =   240
      Left            =   6000
      TabIndex        =   8
      Top             =   390
      Width           =   210
   End
   Begin VB.Label Label2 
      Caption         =   "ScaleY"
      Height          =   240
      Index           =   1
      Left            =   4800
      TabIndex        =   6
      Top             =   555
      Width           =   540
   End
   Begin VB.Label Label2 
      Caption         =   "ScaleX"
      Height          =   240
      Index           =   0
      Left            =   4800
      TabIndex        =   5
      Top             =   225
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "f(x)="
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   3
      Top             =   4980
      Width           =   510
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveas 
         Caption         =   "Save as"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Option"
      Begin VB.Menu mnuQuality 
         Caption         =   "Draw Quality"
         Begin VB.Menu mnuLow 
            Caption         =   "Low"
         End
         Begin VB.Menu mnuMedium 
            Caption         =   "Medium"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuhigh 
            Caption         =   "High"
         End
         Begin VB.Menu mnuVhigh 
            Caption         =   "Very High (Extremely Slow)"
         End
      End
      Begin VB.Menu mnuSolid 
         Caption         =   "Solid"
      End
      Begin VB.Menu mnuClrReplot 
         Caption         =   "Clear Before Replot"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuSyntax 
         Caption         =   "Syntax Help"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const SRCCOPY = &HCC0020
Dim Buffer As Long, Buffhdc As Long
Dim Quality As Double
Dim Cancel As Boolean
Dim Tempfile As String

Private Sub Form_Load()
'Backbuffer in case the form is redrawn
    Buffer = CreateCompatibleBitmap(DrawArea.hdc, 301, 301)
    Buffhdc = CreateCompatibleDC(DrawArea.hdc)
    SelectObject Buffhdc, Buffer
    Me.Show
    DrawArea.BackColor = RGB(250, 250, 248)
    cmdClear_Click
    Quality = 0.1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteObject Buffer
    DeleteDC Buffhdc
    End
End Sub

Private Sub cmdPlot_Click()
    On Error Resume Next
    Dim color As Long, x2 As Double, x As Double, y As Double, x1 As Double, y1 As Double
    Randomize
    
    If cmdPlot.Caption = "Plot" Then
        
        If mnuClrReplot.Checked = True Then cmdClear_Click
        cmdPlot.Caption = "Cancel"
        lblBusy.Visible = True
        cmdClear.Enabled = False
        Cancel = False
        color = RGB(250 * Rnd, 250 * Rnd, 255 * Rnd)
        x1 = 0
        y1 = 0
        
        'Used the script control to parse and execute the string in txtequation
        SC.Reset
        SC.AddCode "Function RunThis(X)" & vbCrLf & "RunThis=" & txtEquation & vbCrLf & "End Function"
        
        If mnuSolid.Checked = True Then
            
            For x2 = -150 To 150 Step Quality
                If Cancel = True Then Exit For
                x = (x2 / 150) * CDbl(txtScaleX)
                y = SC.Run("runthis", x)
                y = -y
                y = y * 150 / CDbl(txtScaleY) + 150
                If Not (x1 = 0 And y1 = 0) Then DrawArea.Line (x1, y1)-(x2 + 150, y), color
                x1 = x2 + 150
                y1 = y
                DoEvents
            Next x2
            
        Else
            
            For x2 = -150 To 150 Step Quality
                If Cancel = True Then Exit For
                x = (x2 / 150) * CDbl(txtScaleX)
                y = SC.Run("runthis", x)
                y = -y
                y = y * 150 / CDbl(txtScaleY) + 150
                SetPixel DrawArea.hdc, x2 + 150, y, color
                SetPixel Buffhdc, x2 + 150, y, color
                x1 = x2 + 150
                y1 = y
                DoEvents
            Next x2
            
        End If
        
        cmdPlot.Caption = "Plot"
        lblBusy.Visible = False
        cmdClear.Enabled = True
        
    Else
        Cancel = True
    End If
End Sub

Private Sub DrawArea_paint()
    BitBlt DrawArea.hdc, 0, 0, 301, 301, Buffhdc, 0, 0, SRCCOPY
End Sub

Private Sub cmdClear_Click()
    DrawArea.Cls
    BitBlt Buffhdc, 0, 0, 301, 301, DrawArea.hdc, 0, 0, SRCCOPY
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuHigh_Click()
    Quality = 0.01
    mnuhigh.Checked = True
    mnuLow.Checked = False
    mnuVhigh.Checked = False
    mnuMedium.Checked = False
End Sub

Private Sub mnuLow_Click()
    Quality = 1
    mnuhigh.Checked = False
    mnuLow.Checked = True
    mnuVhigh.Checked = False
    mnuMedium.Checked = False
End Sub

Private Sub mnuMedium_Click()
    Quality = 0.1
    mnuhigh.Checked = False
    mnuLow.Checked = False
    mnuVhigh.Checked = False
    mnuMedium.Checked = True
End Sub

Private Sub mnuVhigh_Click()
    Quality = 0.001
    mnuhigh.Checked = False
    mnuLow.Checked = False
    mnuVhigh.Checked = True
    mnuMedium.Checked = False
End Sub

Private Sub mnuClrReplot_Click()
    mnuClrReplot.Checked = Not mnuClrReplot.Checked
End Sub

Private Sub mnuSave_Click()
If Cdlg.FileName = "" Then
    mnuSaveas_Click
Else
    save
End If
End Sub

Private Sub mnuSaveas_Click()
    On Error Resume Next
    
    'Saves the Graph Equation and related settings
    Cdlg.CancelError = True
    Tempfile = Cdlg.FileName
    If Cdlg.FileName = "" Then Cdlg.FileName = "Graph.xpl"
    Cdlg.DialogTitle = "Save as XPlot File"
    Cdlg.Flags = &H2
    Cdlg.DefaultExt = ".xpl"
    Cdlg.Filter = "XPlot |*.xpl"
    Cdlg.InitDir = App.Path & "\Samples\"
    Cdlg.ShowSave
    
    'Error 32755 is generated when cancel button is pressed on save dialog
    'without this check picture is saved even if cancel button is pressed
    If Err.Number = 32755 Then
        Cdlg.FileName = Tempfile
        Exit Sub
    End If
    frmPlot.Caption = "XPlot - " & Cdlg.FileTitle
    save
End Sub

Private Sub save()
    On Error GoTo errorh
    ff = FreeFile
    Open Cdlg.FileName For Output As #ff
    Print #ff, txtEquation.Text
    Print #ff, chkConstrain.Value
    Print #ff, txtScaleX.Text
    Print #ff, txtScaleY.Text
    Print #ff, mnuSolid.Checked
    Print #ff, mnuQuality.Enabled
    Print #ff, Quality
    Print #ff, mnuLow.Checked
    Print #ff, mnuMedium.Checked
    Print #ff, mnuhigh.Checked
    Print #ff, mnuVhigh.Checked
    Print #ff, mnuClrReplot.Checked
    Close #ff
    Exit Sub
    
errorh:
    Cdlg.FileName = Tempfile
    MsgBox "Can't Save The File", vbCritical + vbOKOnly, "Error"
    Close #ff
End Sub

Private Sub mnuNew_Click()
    Cdlg.FileName = ""
    frmPlot.Caption = "XPlot"
    txtEquation = "Sin(x)"
    chkConstrain.Value = Checked
    txtScaleX = 10
    mnuSolid.Checked = False
    mnuQuality.Enabled = True
    mnuMedium_Click
    DrawArea.Cls
End Sub

Private Sub mnuOpen_Click()
    On Error Resume Next
    
    'opens the Graph Equation and related settings
    Tempfile = Cdlg.FileName
    Cdlg.CancelError = True
    Cdlg.DialogTitle = "Open XPlot File"
    Cdlg.Flags = &H400
    Cdlg.DefaultExt = ".xpl"
    Cdlg.Filter = "XPlot |*.xpl"
    Cdlg.InitDir = App.Path & "\Samples\"
    Cdlg.ShowOpen
    
    'Error 32755 is generated when cancel button is pressed on save dialog
    'without this check picture is saved even if cancel button is pressed
    If Err.Number = 32755 Then
        Cdlg.FileName = Tempfile
        Exit Sub
    End If
    
    On Error GoTo errorh
    ff = FreeFile
    Open Cdlg.FileName For Input As #ff
    Input #ff, a
    txtEquation = a
    Input #ff, a
    chkConstrain.Value = a
    Input #ff, a
    txtScaleX = a
    Input #ff, a
    txtScaleY = a
    Input #ff, a
    mnuSolid.Checked = a
    Input #ff, a
    mnuQuality.Enabled = a
    Input #ff, a
    Quality = a
    Input #ff, a
    mnuLow.Checked = a
    Input #ff, a
    mnuMedium.Checked = a
    Input #ff, a
    mnuhigh.Checked = a
    Input #ff, a
    mnuVhigh.Checked = a
    Input #ff, a
    mnuClrReplot.Checked = a
    Close #ff
    frmPlot.Caption = "XPlot - " & Cdlg.FileTitle
    cmdClear_Click
    cmdPlot_Click
    Exit Sub
errorh:
    Cdlg.FileName = Tempfile
    MsgBox "Can't Load File", vbCritical + vbOKOnly, "Error"
    Close #ff
End Sub

Private Sub mnuSolid_Click()
    mnuSolid.Checked = Not mnuSolid.Checked
    mnuQuality.Enabled = Not mnuQuality.Enabled
    mnuLow_Click
End Sub

Private Sub mnuSyntax_Click()
    frmHelp.Show vbModal
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub txtScaleX_Change()
'To check whether to change both values in scale text boxes simultaneously
    If chkConstrain.Value = 1 Then txtScaleY = txtScaleX
End Sub

Private Sub txtScaleY_Change()
    If chkConstrain.Value = 1 Then txtScaleX = txtScaleY
End Sub

Private Sub chkConstrain_Click()
    If chkConstrain.Value = 1 Then txtScaleY = txtScaleX
End Sub

Private Sub txtScaleX_KeyPress(KeyAscii As Integer)
'to enable only numeric values and backspace and del etc
    If Not (KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or IsNumeric(Chr(KeyAscii))) Then KeyAscii = 0
End Sub

Private Sub txtScaleY_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or IsNumeric(Chr(KeyAscii))) Then KeyAscii = 0
End Sub

Private Sub DrawArea_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'calculate and show the X and Y coordiantes when mouse moves over drawing area
    PosX = "X= " & Round(x * CDbl(txtScaleX) / 150 - CDbl(txtScaleX), 4)
    PosY = "Y= " & Round(-(y * CDbl(txtScaleY) / 150 - CDbl(txtScaleY)), 4)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Clear the coordinate info area when mouse moves out of drawing area
    PosX = ""
    PosY = ""
End Sub

