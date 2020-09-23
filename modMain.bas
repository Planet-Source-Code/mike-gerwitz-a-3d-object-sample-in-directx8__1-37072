Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Const pi = 3.14159265

Public dx As DirectX8
Public d3d As Direct3D8
Public d3dx As D3DX8
Public d3dDevice As Direct3DDevice8
Public d3dWindow As D3DPRESENT_PARAMETERS
Public d3dDispMode As D3DDISPLAYMODE

Public FramesDrawn As Long
Public FrameRate As Long
Public LastTimeCheckFPS As Long
Public Rotation As Single
Public ViewPoint As D3DVECTOR
Public CameraPoint As D3DVECTOR
Public GameFont As D3DXFont

Public TeaPot As D3DXMesh
Public Box As D3DXMesh
Public Sphere As D3DXMesh
Public Torus As D3DXMesh

Public bRunning As Boolean
Public ShowFrameRate As Boolean
Public BackColor As Long
Public CurrentMesh As Integer

Public Sub InitGame()
    Dim Result As Boolean
    
    Result = InitGraphics
    If Result = False Then
        MsgBox "Unable to initialize DirectX graphics! Here are some tips:" & vbCrLf & _
                     "     *Make Sure DirectX8 is properly installed on your system" & vbCrLf & _
                     "     *You may download it off the microsoft web site at" & vbCrLf & _
                     "       http://msdn.microsoft.com/directx" & vbCrLf & _
                    "The program will now end...", vbCritical, "Fatil Error"
        End
    End If
    
    bRunning = True
    ShowFrameRate = True
    
    ViewPoint = vec(0, 0, 0)
    CameraPoint = vec(5, 5, 5)
    
    BackColor = &H404040
End Sub

Public Function InitGraphics() As Boolean
    Dim Result As Boolean
    
    Result = InitDX
    If Result = False Then
        Debug.Print "Error in InitGraphics Sub (Result = False):"
        Debug.Print Err.Number & " - " & Err.Description
        Debug.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        
        InitGraphics = False
    End If
    
    CreateTeapot
    InitGraphics = True
End Function

Public Function InitDX() As Boolean
    On Error Resume Next
    
    Dim Result As Boolean
    
    Result = InitDXObjects
    If Result = False Then
        Debug.Print "Error in InitDX Sub (Result [1] = False):"
        Debug.Print Err.Number & " - " & Err.Description
        Debug.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        
        InitDX = False
        Exit Function
    End If
    
    Result = InitDXDevice
    If Result = False Then
        Debug.Print "Error in InitDX Sub (Result [2] = False):"
        Debug.Print Err.Number & " - " & Err.Description
        Debug.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        
        InitDX = False
        Exit Function
    End If
    
    InitDXLights
    
    SetDeviceState
    
    Result = InitDXFonts
    If Result = False Then
        Debug.Print "Error in InitDX Sub (Result [3] = False):"
        Debug.Print Err.Number & " - " & Err.Description
        Debug.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        
        InitDX = False
        Exit Function
    End If
    
    InitDX = True
End Function

Public Function InitDXObjects() As Boolean
    On Error Resume Next
    Err.Clear
    
    Set dx = New DirectX8
    If Err.Number <> 0 Then
        Debug.Print "Error in InitDXObjects Sub [1]:"
        Debug.Print Err.Number & " - " & Err.Description
        Debug.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        
        InitDXObjects = False
        Exit Function
    End If
    
    Set d3d = dx.Direct3DCreate()
    If Err.Number <> 0 Then
        Debug.Print "Error in InitDXObjects Sub: [2]"
        Debug.Print Err.Number & " - " & Err.Description
        Debug.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        
        InitDXObjects = False
        Exit Function
    End If
    
    Set d3dx = New D3DX8
    
    InitDXObjects = True
End Function

Public Function InitDXDevice() As Boolean
    Dim Caps As D3DCAPS8
    Dim DevType As CONST_D3DDEVTYPE
    Dim DevBehaviorFlags As Long
    
    On Error Resume Next
    Err.Clear
    
    d3d.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Caps
    If Err.Number = D3DERR_NOTAVAILABLE Then
        DevType = D3DDEVTYPE_REF
        DevBehaviorFlags = D3DCREATE_SOFTWARE_VERTEXPROCESSING
        
    Else
         DevType = D3DDEVTYPE_HAL
    
         If Caps.VertexProcessingCaps = 0 Then
            DevBehaviorFlags = D3DCREATE_SOFTWARE_VERTEXPROCESSING
       
         ElseIf Caps.VertexProcessingCaps = &H4B Then
            DevBehaviorFlags = D3DCREATE_HARDWARE_VERTEXPROCESSING
       
         Else
            DevBehaviorFlags = D3DCREATE_MIXED_VERTEXPROCESSING
        End If
    End If
    
    d3d.GetAdapterDisplayMode D3DADAPTER_DEFAULT, d3dDispMode
    
    d3dWindow.Windowed = 1
    d3dWindow.SwapEffect = D3DSWAPEFFECT_DISCARD
    d3dWindow.BackBufferFormat = d3dDispMode.Format
    d3dWindow.BackBufferCount = 1
    d3dWindow.hDeviceWindow = frmMain.hWnd
    d3dWindow.EnableAutoDepthStencil = 1
    d3dWindow.AutoDepthStencilFormat = D3DFMT_D16
    
    Err.Clear
    Set d3dDevice = d3d.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, _
        DevBehaviorFlags, d3dWindow)
    InitDXDevice = True
End Function

Public Function vec(x As Single, y As Single, z As Single) As D3DVECTOR

vec.x = x
vec.y = y
vec.z = z
 
End Function

Public Sub InitDXLights()
    Dim d3dLight As D3DLIGHT8
    Dim material As D3DMATERIAL8
    Dim Col As D3DCOLORVALUE
    
    Col.a = 1
    Col.b = 1
    Col.g = 1
    Col.r = 1
    
    d3dLight.Type = D3DLIGHT_DIRECTIONAL
    d3dLight.diffuse = Col
    d3dLight.Direction = vec(-1, -1, -1)
    
    material.Ambient = Col
    material.diffuse = Col
    d3dDevice.SetMaterial material
    
    d3dDevice.SetLight 0, d3dLight
    d3dDevice.LightEnable 0, 1
End Sub

Public Sub SetDeviceState()
    d3dDevice.SetRenderState D3DRS_LIGHTING, 1
    d3dDevice.SetRenderState D3DRS_ZENABLE, 1
    d3dDevice.SetRenderState D3DRS_LIGHTING, 1
    d3dDevice.SetRenderState D3DRS_ZENABLE, 1
    d3dDevice.SetRenderState D3DRS_SHADEMODE, D3DSHADE_GOURAUD
    d3dDevice.SetRenderState D3DRS_AMBIENT, &H202020
    d3dDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    d3dDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
End Sub

Public Function InitDXFonts() As Boolean
    Dim fntDesc As IFont
    Dim fnt As StdFont
    
    Set fnt = New StdFont
    fnt.Name = "Arial"
    fnt.Size = 8
    fnt.Bold = True
    Set fntDesc = fnt
    Set GameFont = d3dx.CreateFont(d3dDevice, fntDesc.hFont)
    
    InitDXFonts = True
End Function

Public Sub CreateTeapot()
    Set TeaPot = d3dx.CreateTeapot(d3dDevice, Nothing)
    CurrentMesh = 0
End Sub

Public Sub PlayGame()
    Do While bRunning
        RenderAll
        DoEvents
    Loop
    
    EndGame
End Sub

Public Sub RenderAll()
    Dim matProj As D3DMATRIX
    Dim matView As D3DMATRIX
    Dim matWorld As D3DMATRIX
    Dim matTemp As D3DMATRIX
    Dim matTrans As D3DMATRIX
    
    On Error Resume Next
    
    D3DXMatrixIdentity matWorld
    d3dDevice.SetTransform D3DTS_WORLD, matWorld
    
    D3DXMatrixRotationY matView, Rotation
    D3DXMatrixLookAtLH matTemp, CameraPoint, ViewPoint, vec(0, 1, 0)
    D3DXMatrixMultiply matView, matView, matTemp
    
    d3dDevice.SetTransform D3DTS_VIEW, matView
    
    D3DXMatrixPerspectiveFovLH matProj, pi / 4, 1, 0.1, 500
    d3dDevice.SetTransform D3DTS_PROJECTION, matProj
    
    d3dDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, BackColor, 1, 0
    d3dDevice.BeginScene
        
        If CurrentMesh = 0 Then
            TeaPot.DrawSubset 0
        ElseIf CurrentMesh = 1 Then
            Box.DrawSubset 0
        ElseIf CurrentMesh = 2 Then
            Sphere.DrawSubset 0
        ElseIf CurrentMesh = 3 Then
            Torus.DrawSubset 0
        End If
        
        If ShowFrameRate Then
            DrawFrameRate
            DrawViewPoint
        End If
        
    d3dDevice.EndScene
    
    Err.Clear
    d3dDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    If Err <> 0 Then
        MsgBox "There was an error rendering DirectX objects to the screen!" & vbCrLf & _
                     "The program will now end...", vbCritical, "Fatil Error"
        End
    End If
End Sub

Public Sub DrawFrameRate()
    Dim TextRect As RECT
    
    If GetTickCount - LastTimeCheckFPS >= 1000 Then
       LastTimeCheckFPS = GetTickCount
       FrameRate = FramesDrawn
       FramesDrawn = 0
       
    End If
    FramesDrawn = FramesDrawn + 1
    
    TextRect.Top = 1
    TextRect.Left = 1
    TextRect.bottom = 16
    TextRect.Right = 200
    d3dx.DrawText GameFont, &HFFFF0000, "Current Frame Rate: " & FrameRate, TextRect, DT_TOP Or DT_LEFT
End Sub

Private Sub DrawViewPoint()

    Dim TextRect As RECT
    
    TextRect.Top = 17
    TextRect.Left = 1
    TextRect.bottom = 32
    TextRect.Right = 200
    d3dx.DrawText GameFont, &HFFFF0000, "ViewPoint: (" & FormatNumber(ViewPoint.x, 1) & ", " & FormatNumber(ViewPoint.y, 1) & ", " & FormatNumber(ViewPoint.z, 1) & ")", TextRect, DT_TOP Or DT_LEFT
    
    TextRect.Top = 33
    TextRect.Left = 1
    TextRect.bottom = 48
    TextRect.Right = 200
    d3dx.DrawText GameFont, &HFFFF0000, "CameraPoint: (" & FormatNumber(CameraPoint.x, 1) & ", " & FormatNumber(CameraPoint.y, 1) & ", " & FormatNumber(CameraPoint.z, 1) & ")", TextRect, DT_TOP Or DT_LEFT

    TextRect.Top = 49
    TextRect.Left = 1
    TextRect.bottom = 48
    TextRect.Right = 250
    d3dx.DrawText GameFont, &HFFFF0000, "[Right-Click to display the Popup Menu]", TextRect, DT_TOP Or DT_LEFT
End Sub

Public Sub EndGame()
    On Error Resume Next
    
    Set TeaPot = Nothing
    
    Set d3dx = Nothing
    Set d3d = Nothing
    Set dx = Nothing
    
    End
End Sub

Public Sub ChangeLightColor(Optional r As Integer = 1, Optional g As Integer = 1, Optional b As Integer = 1)
    Dim d3dLight As D3DLIGHT8
    Dim Col As D3DCOLORVALUE
    
    Col.b = b
    Col.g = g
    Col.r = r
    
    d3dLight.Type = D3DLIGHT_DIRECTIONAL
    d3dLight.diffuse = Col
    d3dLight.Direction = vec(-1, -1, -1)
    
    d3dDevice.SetLight 0, d3dLight
End Sub

Public Sub ChangeBackColor(Optional Num As Long = &H404040)
    On Error GoTo InvalidNum
    
    BackColor = Num
    Exit Sub
    
InvalidNum:
    MsgBox "That is not a valid number", vbCritical, "Unable To Change Background Color"
End Sub

Public Sub CreateBox(Optional Width As Long = 3, Optional Height As Long = 3, Optional Depth As Long = 3)
    On Error GoTo ErrOut
    
    Set Box = d3dx.CreateBox(d3dDevice, Width, Height, Depth, Nothing)
    CurrentMesh = 1
    Exit Sub
    
ErrOut:
    MsgBox "There was an error creating your box", vbCritical, "Unable to Create Box"
End Sub

Public Sub CreateSphere(Optional Radius As Long = 3, Optional Slices As Long = 1000, Optional Stacks As Long = 7)
    On Error GoTo ErrOut
    
    Set Sphere = d3dx.CreateSphere(d3dDevice, Radius, Slices, Stacks, Nothing)
    CurrentMesh = 2
    Exit Sub
    
ErrOut:
    MsgBox "There was an error creating your sphere", vbCritical, "Unable to Create Sphere"
End Sub

Public Sub CreateTorus(Optional IRadius As Long = 3, Optional ORadius As Long = 10, Optional Sides As Long = 7, Optional Rings As Long = 5)
    On Error GoTo ErrOut
    
    Set Torus = d3dx.CreateTorus(d3dDevice, IRadius, ORadius, Sides, Rings, Nothing)
    CurrentMesh = 3
    Exit Sub
    
ErrOut:
    MsgBox "There was an error creating your torus", vbCritical, "Unable to Create torus"
End Sub

Public Sub ChangeView(x As Single, y As Single, z As Single, Optional OnlyZoom As Boolean)
    If Not OnlyZoom Then ViewPoint = vec(ViewPoint.x + x, ViewPoint.y + y, ViewPoint.z + z)
    CameraPoint = vec(CameraPoint.x + x, CameraPoint.y + y, CameraPoint.z + z)
End Sub

Public Sub Rotate(angle As Single)
    Rotation = Rotation + angle
End Sub
