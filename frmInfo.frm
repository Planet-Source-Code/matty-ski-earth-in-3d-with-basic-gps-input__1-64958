VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Information"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbProgComm2 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   5550
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnable 
      Caption         =   "&Disable"
      Height          =   375
      Left            =   1680
      TabIndex        =   31
      ToolTipText     =   "Disables 3D Globe"
      Top             =   3690
      Width           =   735
   End
   Begin VB.CommandButton cmdSaveRaw 
      Caption         =   "..."
      Height          =   255
      Left            =   1800
      TabIndex        =   30
      Top             =   6480
      Width           =   375
   End
   Begin VB.TextBox txtRawData 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   29
      ToolTipText     =   "Raw Comm Data"
      Top             =   6360
      Width           =   2415
   End
   Begin VB.ComboBox cmbProgComm1 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   27
      ToolTipText     =   "Throws data to this Comm port for another program, requires a Virtual Serial program"
      Top             =   5190
      Width           =   1215
   End
   Begin VB.ComboBox cmbGPSComm 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Start Stats"
      Height          =   375
      Left            =   1440
      TabIndex        =   23
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdSettings 
      Caption         =   "&Cold Start"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtDegComp 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "N"
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtDeg 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "0"
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtLat 
      Height          =   285
      Left            =   1080
      TabIndex        =   14
      Text            =   "56,N"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtLong 
      Height          =   285
      Left            =   1080
      TabIndex        =   13
      Text            =   "3,E"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtAlt 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtSats 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtZo 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtZ 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtFPS 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0"
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "Prog. Comm via spoof:"
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   32
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Comm Port Error"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   28
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lbl 
      Caption         =   "Prog. Comm via spoof:"
      Height          =   495
      Index           =   13
      Left            =   120
      TabIndex        =   25
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "GPS Comm:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   24
      Top             =   4800
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderStyle     =   2  'Dash
      X1              =   0
      X2              =   2640
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label lbl 
      Caption         =   "Direction:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   20
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Latitude:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Longitude:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Altitude:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Satellites:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderStyle     =   2  'Dash
      X1              =   0
      X2              =   2640
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lbl 
      Caption         =   "Zoom Z:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Z Rads:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Y Rads:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "X Rads:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "FPS:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Move with cusor keys, Pg Up/Dn for zoom, Del/End for Rotate"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   2535
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbGPSComm_Click()
    On Error Resume Next
    ' Set GPS Port
    frmMain.MSComm1.PortOpen = False
    Err = 0
    frmMain.MSComm1.CommPort = cmbGPSComm
    frmMain.MSComm1.PortOpen = True
    
    If Err Then lbl(14).Visible = True Else lbl(14).Visible = False
End Sub


Private Sub cmbProgComm1_Click()
    On Error Resume Next
    ' Set Other Program Port
    frmMain.MSComm2.PortOpen = False
    Err = 0
    frmMain.MSComm2.CommPort = cmbProgComm1
    frmMain.MSComm2.PortOpen = True
    
    If Err Then lbl(14).Visible = True Else lbl(14).Visible = False
End Sub

Private Sub cmbProgComm2_Click()
    On Error Resume Next
    ' Set Other Program Port
    frmMain.MSComm3.PortOpen = False
    Err = 0
    frmMain.MSComm3.CommPort = cmbProgComm2
    frmMain.MSComm3.PortOpen = True
    
    If Err Then lbl(14).Visible = True Else lbl(14).Visible = False
End Sub

Private Sub cmdEnable_Click()
    If cmdEnable.Caption = "&Disable" Then
        frmMain.Tmr.Enabled = False
        cmdEnable.Caption = "&Enable"
        
    Else
        frmMain.Tmr.Enabled = True
        cmdEnable.Caption = "&Disable"
        
    End If
End Sub

Private Sub cmdSave_Click()
    On Error Resume Next
    Dim I As Long, O As Long
    Dim strTmp As String
    
    
    If cmdSave.Caption = "&Start Stats" Then
        ' Start stats file
        strTmp = InputBox("Enter a file name," & vbCrLf & "of the CSV file to log", "Save as", "GPS Logger " & Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH-MM-SS") & ".csv")
        strTmp = Trim(strTmp)
        If strTmp = "" Then Exit Sub
        
        Open strTmp For Output As #2
        Print #2, "Date,Time,Longitude,Latitude,Direction,Speed"
        cmdSave.Caption = "&End Stats"
        StatsRunning = 1
        
    Else
        ' Close stats file
        For I = 1 To UBound(LoggerBuffer, 2)
            For O = 0 To 4
                Print #2, LoggerBuffer(O, I) & ",";
            Next O
            Print #2, LoggerBuffer(5, I)
        Next I
        Print #2, Date & "," & Time & "," & Left(txtLong, Len(txtLong) - 2) & "," & Left(txtLat, Len(txtLat) - 2) & "," & txtDeg & ",0"
        Close #2
        
        ReDim LoggerBuffer(5, 0) As String
        cmdSave.Caption = "&Start Stats"
        StatsRunning = 0
        
    End If
End Sub

Private Sub cmdSaveRaw_Click()
    Dim strTmp As String
    
    
    If SaveRawData = 0 Then
        strTmp = InputBox("Enter a file name," & vbCrLf & "of the Raw Data file to log", "Save as", "GPS Raw Data " & Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH-MM-SS") & ".txt")
        strTmp = Trim(strTmp)
        If strTmp = "" Then Exit Sub
        Open strTmp For Output As #3
        SaveRawData = 1
        
    Else
        SaveRawData = 0
        Close #3
        MsgBox "Stopped Saving Data", , "Stopped"
        
    End If
End Sub

Private Sub cmdSettings_Click()
    frmMain.MSComm1.Output = Chr(&H10) & Chr(2) & Chr(&H12) & Chr(&H85) & Chr(1) & Chr(1) & Chr(1) & Chr(1) & Chr(1) & Chr(2) & Chr(1) & Chr(&HD4) & Chr(7) & Chr(&H28)
    ' Then wait 5 minutes   ;(
End Sub

Private Sub Form_Activate()
    ' Screen update prob, so this has been added
    If frmMain.MSComm1.PortOpen = False Or _
       frmMain.MSComm2.PortOpen = False Or _
       frmMain.MSComm3.PortOpen = False Then lbl(14).Visible = True
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim I As Long
    
    ' Populate Comm Ports
    For I = 1 To 16
        cmbGPSComm.AddItem I
        cmbProgComm1.AddItem I
        cmbProgComm2.AddItem I
    Next I
    
    cmbGPSComm.ListIndex = GetSetting(App.Title, "Settings", "GPS Comm", 3) - 1
    cmbProgComm1.ListIndex = GetSetting(App.Title, "Settings", "Prog Comm 1", 5) - 1
    cmbProgComm2.ListIndex = GetSetting(App.Title, "Settings", "Prog Comm 2", 7) - 1
    cmbProgComm2.ToolTipText = cmbProgComm1.ToolTipText
    
    
    If Err Then lbl(14).Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Settings", "GPS Comm", cmbGPSComm.ListIndex + 1
    SaveSetting App.Title, "Settings", "Prog Comm 1", cmbProgComm1.ListIndex + 1
    SaveSetting App.Title, "Settings", "Prog Comm 2", cmbProgComm2.ListIndex + 1
End Sub

Private Sub txtDeg_Change()
    If Not IsNumeric(txtDeg) Then Exit Sub
    
    If txtDeg > 315 Or txtDeg <= 45 Then txtDegComp = "N"
    If txtDeg > 45 And txtDeg <= 135 Then txtDegComp = "E"
    If txtDeg > 135 And txtDeg <= 225 Then txtDegComp = "S"
    If txtDeg > 225 And txtDeg <= 315 Then txtDegComp = "W"
End Sub
