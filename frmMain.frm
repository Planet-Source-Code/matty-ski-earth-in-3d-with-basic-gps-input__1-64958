VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GPS Globe"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   569
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrComm3 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   840
      Top             =   3600
   End
   Begin VB.Timer tmrComm2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   840
      Top             =   2880
   End
   Begin VB.Timer tmrComm1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   840
      Top             =   1920
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      BaudRate        =   4800
      SThreshold      =   1
   End
   Begin VB.Timer Tmr 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   1320
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   120
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      BaudRate        =   4800
      SThreshold      =   1
   End
   Begin MSCommLib.MSComm MSComm3 
      Left            =   120
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      BaudRate        =   4800
      SThreshold      =   1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Here is a good Virutal Serial Comm Port program to try:
'   http://www.eltima.com/download/vspdxp/
' Or if you have 2 physical comm ports, you can null modem them together.  Cross over pins 2 & 3 - they should be enough





Option Explicit
'Private Declare Function PlaySoundA Lib "winmm.dll" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
'Private Declare Function GetCursor Lib "user32" () As Long
'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function SetCursor Lib "user32" (LYVal hCursor As Long) As Long
'Private Declare Function SetCursorPos Lib "user32" (LYVal X As Long, LYVal Y As Long) As Long
'Private Type POINTAPI
'    X As Long
'    Y As Long
'End Type

' Mostly copied from the SDK
Dim g_DX As New DirectX8
Dim g_D3DX As New D3DX8
Dim g_D3D As Direct3D8                  ' Used to create the D3DDevice
Dim g_D3DDevice As Direct3DDevice8      ' Our rendering device
Dim g_Mesh As D3DXMesh                  ' Our Mesh
Dim g_MeshMaterials() As D3DMATERIAL8   ' Mesh Material data
Dim g_MeshTextures() As Direct3DTexture8 ' Mesh Textures
Dim g_NumMaterials As Long


'Const g_pi = 3.14159265358979
Const rad As Single = 0.0174532925 ' pi/180
Dim Earth As D3DMATRIX


Dim g_VB As Direct3DVertexBuffer8
Dim Vertices(4) As D3DVECTOR      ' For highlighting GPS square
Dim HighlightMaterial As D3DMATERIAL8

Dim ButX As Long, ButY As Long, ButZ As Long, ButZo As Long
Dim RotX As Single, RotY As Single, RotZ As Single, ZoomZ As Single

Dim MouseDown As Byte, MouseX As Single, MouseY As Single

'Dim mainfont As D3DXFont
'Dim MainFontDesc As IFont
'Dim TextRect As RECT
'Dim fnt As New StdFont

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 27: Unload Me
    Case 38: ButX = -1: If Shift = 1 Then ButX = -10 ' Up
    Case 40: ButX = 1: If Shift = 1 Then ButX = 10   ' Down
    Case 37: ButY = -1: If Shift = 1 Then ButY = -10 ' Left
    Case 39: ButY = 1: If Shift = 1 Then ButY = 10   ' Right
    Case 46: ButZ = -1: If Shift = 1 Then ButZ = -10 ' Del
    Case 35: ButZ = 1: If Shift = 1 Then ButZ = 10   ' End
    Case 65, 33: ButZo = 1
    Case 90, 34: ButZo = -1
    End Select
    
    
    
    RotX = RotX + (ButX * rad)
    RotY = RotY + (ButY * rad)
    RotZ = RotZ + (ButZ * rad)
    Dim matWorld(2) As D3DMATRIX
    D3DXMatrixRotationAxis matWorld(0), vec3(1, 0, 0), ButX * rad
    D3DXMatrixRotationAxis matWorld(1), vec3(0, 1, 0), ButY * rad
    D3DXMatrixRotationAxis matWorld(2), vec3(0, 0, 1), ButZ * rad
    D3DXMatrixMultiply Earth, Earth, matWorld(0)
    D3DXMatrixMultiply Earth, Earth, matWorld(1)
    D3DXMatrixMultiply Earth, Earth, matWorld(2)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 38, 40: ButX = 0
    Case 37, 39: ButY = 0
    Case 46, 35: ButZ = 0
    Case 65, 33, 90, 34: ButZo = 0
    End Select
End Sub

Private Sub Form_Load()
    Dim I As Long
    Dim T As Single
    
    
    Randomize Timer
    Me.Show
    Print "Loading..."
    DoEvents
    InitD3D
    InitGeometry
    
    'D3DXMatrixIdentity Earth
    Earth.m11 = 0.9976597     ' Set it for the lower UK
    Earth.m12 = 0.05910983
    Earth.m13 = 0.0350053
    Earth.m21 = -0.006674567
    Earth.m22 = 0.5905674
    Earth.m23 = -0.8069918
    Earth.m31 = -0.06837339
    Earth.m32 = 0.80485
    Earth.m33 = 0.5895666
    Earth.m44 = 1
    ZoomZ = -120
    
    ' Sets GPS square highlight material/colour
    HighlightMaterial.Ambient.r = 1
    HighlightMaterial.Ambient.b = 1
    For I = 0 To 4
        Vertices(I).Z = -50.2
    Next I
    
    ReDim LoggerBuffer(5, 0) As String
    
    
    frmInfo.Show , Me
    frmInfo.Move Me.Left + Me.Width, Me.Top
    frmPOI.Show , Me
    frmPOI.Move Me.Left - frmPOI.Width, Me.Top
    
    
    Me.SetFocus
    ' Zooms in a bit when it starts
    T = Timer
    For I = -200 To -120 Step 5
        ZoomZ = I
        Tmr_Timer
        DoEvents
    Next I
    Tmr.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
    MouseDown = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MouseDown = 1 Then
        RotX = RotX + (ButX * rad)
        RotY = RotY + (ButY * rad)
        RotZ = RotZ + (ButZ * rad)
        Dim matWorld(2) As D3DMATRIX
        D3DXMatrixRotationAxis matWorld(0), vec3(1, 0, 0), ButX * rad
        D3DXMatrixRotationAxis matWorld(1), vec3(0, 1, 0), ButY * rad
        D3DXMatrixRotationAxis matWorld(2), vec3(0, 0, 1), ButZ * rad
        D3DXMatrixMultiply Earth, Earth, matWorld(0)
        D3DXMatrixMultiply Earth, Earth, matWorld(1)
        D3DXMatrixMultiply Earth, Earth, matWorld(2)
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReDim g_MeshTextures(0)
    ReDim g_MeshMaterials(0)
    Set g_Mesh = Nothing
    Set g_D3DDevice = Nothing
    Set g_D3D = Nothing
    End
End Sub

Function InitD3D()
    On Error Resume Next
    Set g_D3D = g_DX.Direct3DCreate()
    If g_D3D Is Nothing Then MsgBox "No Direct3DCreate": End
    
    Dim mode As D3DDISPLAYMODE
    g_D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, mode
    Dim d3dpp As D3DPRESENT_PARAMETERS
    'd3dpp.BackBufferWidth = 1024
    'd3dpp.BackBufferHeight = 768
    d3dpp.Windowed = 1
    d3dpp.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    d3dpp.BackBufferFormat = mode.Format
    d3dpp.BackBufferCount = 1
    d3dpp.EnableAutoDepthStencil = 1
    d3dpp.AutoDepthStencilFormat = D3DFMT_D16
    Set g_D3DDevice = g_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    If g_D3DDevice Is Nothing Then MsgBox "Switching 2 Software", , "Ouch": Set g_D3DDevice = g_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    If g_D3DDevice Is Nothing Then Set g_D3DDevice = g_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_REF, Me.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    
    'g_D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_WIREFRAME
    'g_D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE ' Front & Back
    'g_D3DDevice.SetRenderState D3DRS_ZENABLE, 0
    g_D3DDevice.SetRenderState D3DRS_LIGHTING, 1
    g_D3DDevice.SetRenderState D3DRS_AMBIENT, &HFFFFFFFF
    
    'fnt.Name = "arial"
    'fnt.Size = 12
    'Set MainFontDesc = fnt
    'Set MainFont = g_D3DX.CreateFont(g_D3DDevice, MainFontDesc.hFont)
    
End Function

Sub SetupMatrices()
    'Dim matWorld(2) As D3DMATRIX
    '''D3DXMatrixRotationAxis matWorld, vec3(0, 1, 0), ((-Timer / 86400 * 360) + 180) * (g_pi / 180) ' -Timer / 8
    'RotX = RotX + (ButX * rad)
    'RotY = RotY + (ButY * rad)
    'RotZ = RotZ + (ButZ * rad)
    'D3DXMatrixRotationAxis matWorld(0), vec3(1, 0, 0), RotX
    'D3DXMatrixRotationAxis matWorld(1), vec3(0, 1, 0), RotY
    'D3DXMatrixRotationAxis matWorld(2), vec3(0, 0, 1), RotZ
    'D3DXMatrixMultiply matWorld(0), matWorld(0), matWorld(1)
    'D3DXMatrixMultiply matWorld(0), matWorld(0), matWorld(2)
    'g_D3DDevice.SetTransform D3DTS_WORLD, matWorld(0)
    g_D3DDevice.SetTransform D3DTS_WORLD, Earth
    
    
    
    If Tmr.Enabled = True Then
        ' Not used when it zooms in on the start
        Select Case ZoomZ
        Case Is < -121.1031
            ZoomZ = -121.1031
        Case Is > -51.50161
            ZoomZ = -51.50161
        End Select
        ZoomZ = ZoomZ - ((ButZo * ZoomZ) / 100) ' Makes the zoom a bit logarithmic
    End If
    
    Dim matView As D3DMATRIX
    D3DXMatrixLookAtLH matView, vec3(0, 0, ZoomZ), vec3(0, 0, 0), vec3(0, 1, 0)
    g_D3DDevice.SetTransform D3DTS_VIEW, matView
    
    
    
    Dim matProj As D3DMATRIX
    D3DXMatrixPerspectiveFovLH matProj, 1, Me.Height / Me.Width, 1, 500
    g_D3DDevice.SetTransform D3DTS_PROJECTION, matProj
End Sub

Function vec3(X As Single, Y As Single, Z As Single) As D3DVECTOR
    vec3.X = X: vec3.Y = Y: vec3.Z = Z
End Function

Function InitGeometry()
    Dim MtrlBuffer As D3DXBuffer ' a d3dxbuffer is a generic chunk of memory
    Dim I As Long
    
    
    Set g_Mesh = g_D3DX.LoadMeshFromX("earf.x", D3DXMESH_MANAGED, g_D3DDevice, Nothing, MtrlBuffer, g_NumMaterials)
    If g_Mesh Is Nothing Then Exit Function
    'allocate space for our materials and textures
    ReDim g_MeshMaterials(g_NumMaterials)
    ReDim g_MeshTextures(g_NumMaterials)
    Dim strTexName As String
    ' We need to extract the material properties and texture names
    ' from the MtrlBuffer
    For I = 0 To g_NumMaterials - 1
        ' Copy the material using the d3dx helper function
        g_D3DX.BufferGetMaterial MtrlBuffer, I, g_MeshMaterials(I)
        ' Set the ambient color for the material (D3DX does not do this)
        g_MeshMaterials(I).Ambient = g_MeshMaterials(I).diffuse
        ' Create the texture
        strTexName = g_D3DX.BufferGetTextureName(MtrlBuffer, I)
        ' Earth picture copied from ESA, thank you - http://www.esa.int/esaEO/SEMGSY2IU7E_index_1.html
        If strTexName <> "" Then Set g_MeshTextures(I) = g_D3DX.CreateTextureFromFile(g_D3DDevice, strTexName)
    Next
    Set MtrlBuffer = Nothing
End Function

Private Sub MSComm1_OnComm()
    ' For GPS
    If MSComm1.CommEvent = 2 Then tmrComm1.Enabled = False: tmrComm1.Enabled = True
End Sub

Private Sub MSComm2_OnComm()
    ' For Program 1
    If MSComm2.CommEvent = 2 Then tmrComm2.Enabled = False: tmrComm2.Enabled = True
End Sub

Private Sub MSComm3_OnComm()
    ' For Program 2
    If MSComm3.CommEvent = 2 Then tmrComm3.Enabled = False: tmrComm3.Enabled = True
End Sub

Private Sub Tmr_Timer()
    Dim I As Long, Sng As Single
    Dim XX As Single, YY As Single, ZZ As Single
    
    
    ' Clear the z buffer to 1
    g_D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &HF, 1, 0
    ' Setup the world, view, and projection matrices
    SetupMatrices
    ' Begin the scene
    g_D3DDevice.BeginScene
    ' Meshes are divided into subsets, one for each material.
    ' Render them in a loop
    For I = 0 To g_NumMaterials - 1
        ' Set the material and texture for this subset
        g_D3DDevice.SetMaterial g_MeshMaterials(I)
        g_D3DDevice.SetTexture 0, g_MeshTextures(I)
        ' Draw the mesh subset
        g_Mesh.DrawSubset I
    Next
    
    
    
    
    ' GPS Location Square
    On Error Resume Next
    Dim PosGPS(1) As D3DMATRIX
    D3DXMatrixRotationAxis PosGPS(0), vec3(1, 0, 0), (Left(frmInfo.txtLat, Len(frmInfo.txtLat) - 2) + 2.8) * rad    ' Plus a minor correction
    D3DXMatrixRotationAxis PosGPS(1), vec3(0, 1, 0), -(Left(frmInfo.txtLong, Len(frmInfo.txtLong) - 2) - 3.2) * rad
    D3DXMatrixMultiply PosGPS(0), PosGPS(0), PosGPS(1)
    D3DXMatrixMultiply PosGPS(0), PosGPS(0), Earth ' Earth afterwards
    g_D3DDevice.SetTransform D3DTS_WORLD, PosGPS(0)
    
    Set g_VB = g_D3DDevice.CreateVertexBuffer(12 * 5, 0, D3DFVF_XYZ Or D3DFVF_DIFFUSE, D3DPOOL_DEFAULT)
    g_D3DDevice.SetMaterial HighlightMaterial
    g_D3DDevice.SetTexture 0, Nothing
    For Sng = -0.01 To 0.01 Step 0.0016
        Vertices(0).X = -0.1 - Sng: Vertices(0).Y = -0.1 - Sng
        Vertices(1).X = 0.1 + Sng: Vertices(1).Y = -0.1 - Sng
        Vertices(2).X = 0.1 + Sng: Vertices(2).Y = 0.1 + Sng
        Vertices(3).X = -0.1 - Sng: Vertices(3).Y = 0.1 + Sng
        Vertices(4).X = -0.1 - Sng: Vertices(4).Y = -0.1 - Sng
        D3DVertexBuffer8SetData g_VB, 0, 12 * 5, 0, Vertices(0)
        g_D3DDevice.SetStreamSource 0, g_VB, 12
        g_D3DDevice.DrawPrimitive D3DPT_LINESTRIP, 0, 4
    Next Sng
    
    
    
    
    ' End the scene
    g_D3DDevice.EndScene
    ' Present the backbuffer contents to the front buffer (screen)
    g_D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    
    
    
    frmInfo.txtX = RotX
    frmInfo.txtY = RotY
    frmInfo.txtZ = RotZ
    frmInfo.txtZo = ZoomZ
    'Dim PP As POINTAPI
    'GetCursorPos PP
    'If PP.X <> 320 Then SetCursorPos 320, 240
    
    
    
    ' FPS
    Static TT As Single, MM As Long
    If MM = 0 Then TT = Timer
    MM = MM + 1
    If MM > 0 Then If TT + 1 <= Timer Then frmInfo.txtFPS = MM: MM = 0
End Sub

Private Sub tmrComm1_Timer()
    ' I should check the checksum, but due to lazyness I dont,
    '  if you check the 'View Stat Data' it creates a checksum there to view
    
    On Error Resume Next
    Dim str As String
    Dim I As Long
    Dim GPGGA() As String
    Dim GPRMC() As String
    Dim Deg As Integer, Min As Double
    Dim Comm1Data As String
    tmrComm1.Enabled = False
    
    
    ' Copy Data
    Comm1Data = MSComm1.Input
            frmInfo.txtRawData = frmInfo.txtRawData & "From First Comm:" & vbCrLf & Comm1Data
            If Len(frmInfo.txtRawData) > 5000 Then frmInfo.txtRawData = Right(frmInfo.txtRawData, 3000)
            frmInfo.txtRawData.SelStart = 5000
    If SaveRawData = 1 Then Print #3, Comm1Data
    MSComm2.Output = Comm1Data
    MSComm3.Output = Comm1Data
    
    
    
    ' Extract some data - move it into an array
    I = InStr(Comm1Data, "$GPGGA")
    If I = 0 Then Exit Sub
    str = Mid$(Comm1Data, I + 7, InStr(I, Comm1Data, vbCrLf) - I - 7)
    GPGGA = Split(str, ",")
    
    ' Latitude, thanks go to 'Mike Morrow' for the next few lines
    Deg = GPGGA(1) \ 100
    Min = GPGGA(1) - (Deg * 100)
    GPGGA(1) = Deg + (Min / 60)
    If GPGGA(2) = "S" Then GPGGA(1) = -GPGGA(1)
    ' Longatude
    Deg = GPGGA(3) \ 100
    Min = GPGGA(3) - (Deg * 100)
    GPGGA(3) = Deg + (Min / 60)
    If GPGGA(4) = "W" Then GPGGA(3) = -GPGGA(3)
    
    ' Display infomation
    frmInfo.txtLat = Format(GPGGA(1), "0.#########") & "," & GPGGA(2)
    frmInfo.txtLong = Format(GPGGA(3), "0.#########") & "," & GPGGA(4)
    frmInfo.txtSats = GPGGA(6)
    frmInfo.txtAlt = GPGGA(8)
    
    ' Gets Direction from a different sentance
    I = InStr(Comm1Data, "$GPRMC")
    If I = 0 Then Exit Sub
    str = Mid$(Comm1Data, I + 7, InStr(I, Comm1Data, vbCrLf) - I - 7)
    GPRMC = Split(str, ",")
    
    frmInfo.txtDeg = GPRMC(7)
    
    
    
    ' Check Logger
    If StatsRunning = 1 Then
        I = UBound(LoggerBuffer, 2)
        'If (Int(Val(LoggerBuffer(4, I))) <> Int(Val(LoggerBuffer(4, I - 1)))) And (Int(Val(LoggerBuffer(5, I))) <> Int(Val(LoggerBuffer(5, I - 1)))) Then
            I = I + 1
            ReDim Preserve LoggerBuffer(5, I)
            LoggerBuffer(0, I) = Date
            LoggerBuffer(1, I) = Time
            LoggerBuffer(2, I) = GPGGA(3)
            LoggerBuffer(3, I) = GPGGA(1)
            LoggerBuffer(4, I) = GPRMC(7)
            LoggerBuffer(5, I) = GPRMC(6)
        'End If
    End If
    
    
    
    ' Check the POI list
    frmPOI.CheckPOI CSng(GPGGA(1)), CSng(GPGGA(3))
End Sub

Private Sub tmrComm2_Timer()
    On Error Resume Next
    Dim Comm2Data As String
    tmrComm2.Enabled = False
    
    
    ' Copy Data
    Comm2Data = MSComm2.Input
            frmInfo.txtRawData = frmInfo.txtRawData & "From Second Comm:" & vbCrLf & Replace(Comm2Data, Chr(0), " ")
            frmInfo.txtRawData.SelStart = 5000
    If SaveRawData = 1 Then Print #3, Comm2Data
    MSComm1.Output = Comm2Data
End Sub

Private Sub tmrComm3_Timer()
    On Error Resume Next
    Dim Comm3Data As String
    tmrComm3.Enabled = False
    
    
    ' Copy Data
    Comm3Data = MSComm3.Input
            frmInfo.txtRawData = frmInfo.txtRawData & "From Third Comm:" & vbCrLf & Replace(Comm3Data, Chr(0), " ")
            frmInfo.txtRawData.SelStart = 5000
    If SaveRawData = 1 Then Print #3, Comm3Data
    MSComm1.Output = Comm3Data
End Sub
