VERSION 5.00
Begin VB.Form frmPOI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Points Of Interest"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   Icon            =   "frmPOI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEnable 
      Caption         =   "&Disable"
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtLonVar 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "0.0016"
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox txtLatVar 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "0.001"
      Top             =   4560
      Width           =   2415
   End
   Begin VB.ListBox lstPOIs 
      Height          =   2400
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "&Info"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Re&load"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "pocketgps_uk_sc.csv"
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(Activate GPS)"
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   12
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "Longitude Variance:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label lbl 
      Caption         =   "Latitude Variance:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label lbl 
      Caption         =   "Near POIs:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderStyle     =   2  'Dash
      X1              =   0
      X2              =   2640
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lbl 
      Caption         =   "File:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "frmPOI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim POIstr() As String ' X
Dim POIsng() As Single ' X,X ' Should really be Double

Private Sub cmdEnable_Click()
    If cmdEnable.Caption = "&Disable" Then cmdEnable.Caption = "&Enable" Else cmdEnable.Caption = "&Disable"
End Sub

Private Sub cmdInfo_Click()
    MsgBox "Uses CSV files (3 cells per line), ignores first Header line" & vbCrLf & vbCrLf & "E.G. Download database from http://www.pocketgpsworld.com/uksafetycameras.php" & vbCrLf & "pocketgps_uk_sc_xxx##.zip\single file version\csv file\other\pocketgps_uk_sc.csv" & vbCrLf & "(Now requires registration, sorry was too slow to release this)", , "POI files"
End Sub

Private Sub cmdLoad_Click()
    On Error Resume Next
    Dim I As Long
    ReDim POIstr(0) As String
    'Dim MinMaxLat(1) As Single, MinMaxLong(1) As Single
    
    
    Open txtFile For Input As #1
    Line Input #1, POIstr(0)
    Do
        ReDim Preserve POIstr(I)
        ReDim Preserve POIsng(1, I)
        Input #1, POIsng(0, I), POIsng(1, I), POIstr(I)
        I = I + 1
    Loop Until EOF(1)
    Close #1
    
    If Err Then lbl(9).Caption = "Error loading" Else lbl(9).Caption = Format(I, "#,###,###") & " POIs Loaded"
End Sub

Private Sub Form_Load()
    lstPOIs.ToolTipText = "Decimal Lat/Long: 0.1=36432 ft 22176 ft || 0.01=3643.2 ft 2217.6 ft || 0.001=364.32 ft 221.76 ft || 0.0001=36.43 ft 22.18 ft || 0.00001=3.64 ft 2.22 ft"
    
    cmdLoad_Click
End Sub

Sub CheckPOI(Lat As Single, Lon As Single) ' Double really
    'On Error Resume Next
    Dim I As Long
    Dim tmpLat As Single, tmpLong As Single ' Double really
    
    
    If cmdEnable.Caption = "&Enable" Then Exit Sub
    
    lstPOIs.Clear
    tmpLat = Val(txtLatVar)
    tmpLong = Val(txtLonVar)
    For I = 0 To UBound(POIsng(), 2)
        If (Lat < (POIsng(1, I) + tmpLat)) And (Lat > (POIsng(1, I) - tmpLat)) And _
           (Lon < (POIsng(0, I) + tmpLong)) And (Lon > (POIsng(0, I) - tmpLong)) Then
                lstPOIs.AddItem "Lat    " & POIsng(0, I)
                lstPOIs.AddItem "Long " & POIsng(1, I)
                lstPOIs.AddItem "Desc " & POIstr(I)
                lstPOIs.AddItem ""
        End If
        'DoEvents
    Next I
    
    If lstPOIs.ListCount > 0 Then Beep
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub
