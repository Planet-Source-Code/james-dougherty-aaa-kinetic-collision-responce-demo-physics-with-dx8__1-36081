VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'
  '|¶¶             © 2001-2002 Ariel Productions          ¶¶|'
  '|¶¶                 All Rights Reserved                ¶¶|'
  '|¶¶¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¶¶|'
  '|¶¶             Programmer - James Dougherty           ¶¶|'
  '|¶¶             Source - XEngine.cls                   ¶¶|'
  '|¶¶             Object - UltimaX.dll                   ¶¶|'
  '|¶¶             Version - 2.1                          ¶¶|'
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'

'PSC - I am still NOT done with my new engine
'therefore, there is no commenting
'Sorry for any troubles
'Still needs optimized
Option Explicit

Public Enum FilterType
 None = 0
 Dark_Map = 1
 Light_Map = 2
 Detail_Map = 3
 Gloss_Map = 4
 Detail_Modulate = 5
 Detail_Add = 6
 Add_Smooth_Saturation = 7
 Linear_Blend = 8
 Inverse_Light_Map = 9
End Enum

Private Type TLVERTEX
 X As Single
 Y As Single
 Z As Single
 RHW As Single
 Color As Long
 Specular As Long
 tU As Single
 tV As Single
End Type
Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Private Type TexPoolEntry
 Name As String
 Texture As Direct3DTexture8
 NextDelNode As Long
End Type

Private Type tTexturePool
 TexPool() As TexPoolEntry
 MaxPool As Long
 NextEmpty As Long
 FirstDel As Long
End Type: Private TextPool As tTexturePool

Private tmpIn As New XInput
Private Math As New XMath
Private DX8 As New DirectX8
Private D3DX As New D3DX8
Private D3D As Direct3D8
Private HoldCaps As D3DCAPS8
Private HoldD3DPPWIN As D3DPRESENT_PARAMETERS
Private HoldD3DPPFS As D3DPRESENT_PARAMETERS
Private D3DPP As D3DPRESENT_PARAMETERS
Private HoldWidth As Long
Private HoldHeight As Long
Private HoldhWnd As Long
Private FullS As Boolean
Private InitSucessfull As Boolean
Private vbFont As IFont
Private fnt As New StdFont
Private xFont As D3DXFont
Private TextInitialized As Boolean
Private RenderColor As Long
Private FPS As Single
Private BackVerts(3) As TLVERTEX
Private Background_Picture As Direct3DTexture8
Private HasBackground As Boolean
Private DoDark As Boolean
Private DoDetail As Boolean
Private DoLight As Boolean
Private DoGloss As Boolean
Private DoDMod As Boolean
Private DoDAdd As Boolean
Private DoASat As Boolean
Private DoLBlend As Boolean
Private DoNone As Boolean
Private DisplayTexture As Direct3DTexture8
Private DisplayTextureSurface As Direct3DSurface8
Private DisplayZSurface As Direct3DSurface8
Private BackBuffer As Direct3DSurface8
Private ZBuffer As Direct3DSurface8
Private MapTexture As Direct3DTexture8
Private VertList(0 To 3) As TLVERTEX
Private BlurAmount As Long
Private BlurEnabled As Boolean
Private bShowMouse As Boolean
Public SystemFrame As Long
Private Tick As Long
Private OldSys As Long
Private Const kGrowSize = 10
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef RECT As RECT) As Long

'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|
'|œœ                     User Access                        œœ|
'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|

Public Function DirectX() As DirectX8
 Set DirectX = DX8
End Function

Public Function Direct3DX() As D3DX8
 Set Direct3DX = D3DX
End Function

Public Function Direct3D() As Direct3D8
 Set Direct3D = D3D
End Function

Public Function Direct3DDevice() As Direct3DDevice8
 Set Direct3DDevice = D3DD
End Function

'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|
'|œœ                    Setup Dialog                        œœ|
'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|

Public Function SetupDialog_InitializeAuto(Form As Object, FormsCaption As String)
 SetupDialog_Show
 DoEvents

 If SetupDialog_CancelWasClicked Then End

 If SetupDialog_OkWasClicked Then
  If SetupDialog_IsFullScreen Then
   Form.Caption = ""
   Form.Show
   Initialize_EngineX8 Form.hWnd, True, SetupDialog_FS_Width, SetupDialog_FS_Height
   If SetupDialog_ShowMouseWasChecked Then ShowCursor 0: bShowMouse = True
  ElseIf Not SetupDialog_IsFullScreen Then
   Form.Caption = FormsCaption
   Form.Width = SetupDialog_Window_Width
   Form.Height = SetupDialog_Window_Height
   Form.Show
   XE.Initialize_EngineX8 Form.hWnd, False
  End If
 End If

End Function

Public Function SetupDialog_Show() As Boolean
 On Local Error Resume Next
 Dim HasIt As Boolean
 Dim i As Long
 
 frmSetup.Show
 If D3D Is Nothing Then Set D3D = DX8.Direct3DCreate
 tmpIn.Initialize_Input_Engine frmSetup.hWnd
 frmSetup.Hide
 Enum_Display_Modes frmSetup.cmbModes
 
 For i = 0 To frmSetup.cmbModes.ListCount - 1
  If frmSetup.cmbModes.List(i) = "800 x 600" Then
   frmSetup.cmbModes.ListIndex = i
   Exit For
  End If
 Next

 frmSetup.lblVideo = Get_Default_Adapter

 HasIt = tmpIn.Joystick_HasJoystick
 If HasIt Then
  frmSetup.lblJoy2 = "Found"
  frmSetup.lblJoy2.ForeColor = &HFF00&
 Else
  frmSetup.lblJoy2 = "No Joystick Found"
  frmSetup.lblJoy2.ForeColor = &HFF&
 End If
 HasIt = tmpIn.Joystick_HasForceFeedback
 If HasIt Then
  frmSetup.lblFF2 = "Supported"
  frmSetup.lblFF2.ForeColor = &HFF00&
 Else
  frmSetup.lblFF2 = "Not Supported"
  frmSetup.lblFF2.ForeColor = &HFF&
 End If
 
 frmSetup.Show 1
End Function

Public Function SetupDialog_OkWasClicked() As Boolean
 SetupDialog_OkWasClicked = frmSetup.OkWasClicked
End Function

Public Function SetupDialog_CancelWasClicked() As Boolean
 SetupDialog_CancelWasClicked = frmSetup.CancelWasClicked
End Function

Public Function SetupDialog_Window_Width() As Long
 SetupDialog_Window_Width = (frmSetup.Get_Win_Width * 15)
End Function

Public Function SetupDialog_Window_Height() As Long
 SetupDialog_Window_Height = (frmSetup.Get_Win_Height * 15)
End Function

Public Function SetupDialog_FS_Width() As Long
 SetupDialog_FS_Width = frmSetup.Get_FS_Width
End Function

Public Function SetupDialog_FS_Height() As Long
 SetupDialog_FS_Height = frmSetup.Get_FS_Height
End Function

Public Function SetupDialog_IsFullScreen() As Boolean
 SetupDialog_IsFullScreen = frmSetup.Fullscreen
End Function

Public Function SetupDialog_ShowMouseWasChecked() As Boolean
 SetupDialog_ShowMouseWasChecked = frmSetup.Get_Show_Cursor
End Function

'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|
'|œœ                     Main Setup                         œœ|
'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|

Private Function Initialize_Win(hWnd As Long) As Boolean
 On Local Error GoTo ErrOut
 Dim DType As Long
 Dim Caps As D3DCAPS8
 Dim ScRect As RECT
 Dim Screen_Height As Long
 Dim Screen_Width As Long
 
 CausedAt = ""
 If DX8 Is Nothing Then Set DX8 = New DirectX8
 If D3DX Is Nothing Then Set D3DX = New D3DX8
 If D3D Is Nothing Then Set D3D = DX8.Direct3DCreate
 
 InitSucessfull = False
 
 Dim Mode As D3DDISPLAYMODE
 CausedAt = "GetAdapterDisplayMode"
 D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode
         
 CausedAt = "GetDeviceCaps"
 DType = D3DDEVTYPE_HAL
 D3D.GetDeviceCaps D3DADAPTER_DEFAULT, DType, Caps
 If Err.Number = D3DERR_INVALIDDEVICE Then
  DType = D3DDEVTYPE_REF
  D3D.GetDeviceCaps D3DADAPTER_DEFAULT, DType, Caps
  Err.Clear
 End If
  
 CausedAt = "CheckDeviceMultiSampleType"
 If D3D.CheckDeviceMultiSampleType(D3DADAPTER_DEFAULT, DType, Mode.Format, True, D3DMULTISAMPLE_4_SAMPLES) Then
  D3DPP.MultiSampleType = D3DMULTISAMPLE_4_SAMPLES
 ElseIf D3D.CheckDeviceMultiSampleType(D3DADAPTER_DEFAULT, DType, Mode.Format, True, D3DMULTISAMPLE_3_SAMPLES) Then
  D3DPP.MultiSampleType = D3DMULTISAMPLE_3_SAMPLES
 ElseIf D3D.CheckDeviceMultiSampleType(D3DADAPTER_DEFAULT, DType, Mode.Format, True, D3DMULTISAMPLE_2_SAMPLES) Then
  D3DPP.MultiSampleType = D3DMULTISAMPLE_2_SAMPLES
 Else
  D3DPP.MultiSampleType = D3DMULTISAMPLE_NONE
 End If
 
 D3DPP.Windowed = 1
 D3DPP.SwapEffect = D3DSWAPEFFECT_DISCARD
 D3DPP.BackBufferFormat = Mode.Format
 D3DPP.BackBufferCount = 1
 D3DPP.EnableAutoDepthStencil = 1
 D3DPP.AutoDepthStencilFormat = D3DFMT_D16
 D3DPP.hDeviceWindow = hWnd
  
 CausedAt = "CreateDevice"
 If (Caps.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT) Then
  Set D3DD = D3D.CreateDevice(D3DADAPTER_DEFAULT, DType, hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DPP)
 Else
  Set D3DD = D3D.CreateDevice(D3DADAPTER_DEFAULT, DType, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DPP)
 End If
 
 D3DD.SetRenderState D3DRS_LIGHTING, 0
 D3DD.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
 D3DD.SetRenderState D3DRS_MULTISAMPLE_ANTIALIAS, 1
 D3DD.SetRenderState D3DRS_DITHERENABLE, 1
 
 HoldWidth = Screen_Width
 HoldHeight = Screen_Height
 D3DD.GetDeviceCaps HoldCaps
 HoldD3DPPWIN = D3DPP
 Initialize_Win = True
 FullS = False
 InitSucessfull = True
 Exit Function
 
ErrOut:
 ErrorToFile "Error - WIN", CausedAt, "N/A"
 Initialize_Win = False
 InitSucessfull = False
End Function

Private Function Restore_Window_Device(hWnd As Long) As Boolean
 On Local Error GoTo ErrOut
 If DX8 Is Nothing Then Exit Function
 If D3DX Is Nothing Then Exit Function
 If D3D Is Nothing Then Exit Function
 
 D3DPP.BackBufferWidth = 0
 D3DPP.BackBufferHeight = 0
 D3DPP.hDeviceWindow = hWnd
 D3DPP.Windowed = 1
 Restore_Window_Device = True
 Exit Function
 
ErrOut:
 Restore_Window_Device = False
End Function

Private Function Initialize_FS(hWnd As Long, Width As Long, Height As Long) As Boolean
 On Local Error GoTo ErrOut
 Dim Mode As D3DDISPLAYMODE
 Dim rc As RECT
 Dim Ret As Long
 Dim DType As Long
 Dim Caps As D3DCAPS8
 Dim Caps1 As D3DCAPS8
 Dim Format As Long
 If DX8 Is Nothing Then Set DX8 = New DirectX8
 If D3DX Is Nothing Then Set D3DX = New D3DX8
 If D3D Is Nothing Then Set D3D = DX8.Direct3DCreate
    
 InitSucessfull = False
 
 CausedAt = "GetDeviceCaps"
 DType = D3DDEVTYPE_HAL
 D3D.GetDeviceCaps D3DADAPTER_DEFAULT, DType, Caps
 If Err.Number = D3DERR_INVALIDDEVICE Then
  DType = D3DDEVTYPE_REF
  D3D.GetDeviceCaps D3DADAPTER_DEFAULT, DType, Caps
  Err.Clear
 End If
 
 CausedAt = "GetAdapterDisplayMode"
 D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode
 CausedAt = "Get_Display_Mode"
 Format = Mode.Format
 Ret = Get_Display_Mode(Width, Height, Format)
   
 CausedAt = "CheckDeviceMultiSampleType"
 If D3D.CheckDeviceMultiSampleType(D3DADAPTER_DEFAULT, DType, Format, False, D3DMULTISAMPLE_4_SAMPLES) Then
  D3DPP.MultiSampleType = D3DMULTISAMPLE_4_SAMPLES
 ElseIf D3D.CheckDeviceMultiSampleType(D3DADAPTER_DEFAULT, DType, Format, False, D3DMULTISAMPLE_3_SAMPLES) Then
  D3DPP.MultiSampleType = D3DMULTISAMPLE_3_SAMPLES
 ElseIf D3D.CheckDeviceMultiSampleType(D3DADAPTER_DEFAULT, DType, Format, False, D3DMULTISAMPLE_2_SAMPLES) Then
  D3DPP.MultiSampleType = D3DMULTISAMPLE_2_SAMPLES
 Else
  D3DPP.MultiSampleType = D3DMULTISAMPLE_NONE
 End If
 
 D3DPP.Windowed = 0
 D3DPP.SwapEffect = D3DSWAPEFFECT_FLIP
 D3DPP.BackBufferFormat = Format
 D3DPP.BackBufferCount = 1
 D3DPP.BackBufferWidth = Width
 D3DPP.BackBufferHeight = Height
 D3DPP.EnableAutoDepthStencil = 1
 D3DPP.AutoDepthStencilFormat = D3DFMT_D16
 D3DPP.hDeviceWindow = hWnd
     
 CausedAt = "CreateDevice"
 If (Caps.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT) Then
  Set D3DD = D3D.CreateDevice(D3DADAPTER_DEFAULT, DType, hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DPP)
 Else
  Set D3DD = D3D.CreateDevice(D3DADAPTER_DEFAULT, DType, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DPP)
 End If
   
 D3DD.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
 D3DD.SetRenderState D3DRS_MULTISAMPLE_ANTIALIAS, 1
 D3DD.SetRenderState D3DRS_DITHERENABLE, 1
 D3DD.GetDeviceCaps HoldCaps
 
 HoldD3DPPFS = D3DPP
 Initialize_FS = True
 HoldWidth = Width
 HoldHeight = Height
 FullS = True
 InitSucessfull = True
 Exit Function
 
ErrOut:
 ErrorToFile "ErrorFS", CausedAt, "N/A"
 Initialize_FS = False
 InitSucessfull = False
End Function

Private Function Restore_FullScreen_Device(hWnd As Long) As Boolean
 On Local Error GoTo ErrOut
 
 Initialize_FS hWnd, HoldWidth, HoldHeight
 Setup_Default_Rendering_State
 Restore_FullScreen_Device = True
 Exit Function
    
ErrOut:
 Restore_FullScreen_Device = False
End Function

Private Function Get_Display_Mode(ByVal Width As Long, ByVal Height As Long, Format As Long) As Long
Dim i  As Long
Dim Count As Long
Dim d3ddm As D3DDISPLAYMODE
Dim FoundMode As Boolean
    
i = 0
Count = D3D.GetAdapterModeCount(D3DADAPTER_DEFAULT) - 1
If Err.Number Then Get_Display_Mode = D3DERR_INVALIDCALL: Exit Function
    
For i = 0 To Count
 Call D3D.EnumAdapterModes(D3DADAPTER_DEFAULT, i, d3ddm)
 If Err.Number Then Get_Display_Mode = Err.Number: Exit Function
 If d3ddm.Width = Width And d3ddm.Height = Height Then
  If d3ddm.Format = D3DFMT_A8R8G8B8 Or _
     d3ddm.Format = D3DFMT_R8G8B8 Or _
     d3ddm.Format = D3DFMT_R5G6B5 Then
   FoundMode = True
   Format = d3ddm.Format
   Exit For
  End If
 End If
Next
        
If FoundMode Then Exit Function Else Get_Display_Mode = D3DERR_INVALIDDEVICE
End Function

Public Function Get_Default_Adapter() As String
On Local Error Resume Next
Dim ID As D3DADAPTER_IDENTIFIER8
Dim tmpStr As String
Dim j As Long
Dim i As Long

D3D.GetAdapterIdentifier j, 0, ID
        
For i = 0 To 511
 tmpStr = tmpStr & (Chr(ID.Description(i)))
Next i
        
tmpStr = Replace$(tmpStr, Chr(0), "")
Get_Default_Adapter = tmpStr
End Function

Public Function Enum_Adapters(ComboBox As Object)
On Local Error Resume Next
Dim ID As D3DADAPTER_IDENTIFIER8
Dim numAdapters As Long
Dim tmpStr As String
Dim j As Long
Dim i As Long

ComboBox.Clear
numAdapters = D3D.GetAdapterCount
D3D.GetAdapterIdentifier j, 0, ID
        
For j = 0 To numAdapters - 1
 For i = 0 To 511
  tmpStr = tmpStr & (Chr(ID.Description(i)))
 Next i
 tmpStr = Replace$(tmpStr, Chr(0), "")
 ComboBox.AddItem tmpStr
Next j

End Function

Public Function Enum_Display_Modes(ComboBox As Object) As Long
On Local Error Resume Next
Dim i  As Long
Dim Z As Long
Dim Count As Long
Dim d3ddm As D3DDISPLAYMODE
Dim sMode As String
    
i = 0
ComboBox.Clear
Count = D3D.GetAdapterModeCount(D3DADAPTER_DEFAULT) - 1
If Err.Number Then Enum_Display_Modes = D3DERR_INVALIDCALL: Exit Function
    
For i = 0 To Count
 Call D3D.EnumAdapterModes(D3DADAPTER_DEFAULT, i, d3ddm)
 If Err.Number Then Enum_Display_Modes = Err.Number: Exit Function
 sMode = d3ddm.Width & " x "
 sMode = sMode & d3ddm.Height
 
 For Z = 0 To i
  If sMode <> CStr(ComboBox.List(Z)) Then
   sMode = sMode
  Else
   sMode = ""
  End If
 Next Z
 If sMode <> "" Then ComboBox.AddItem sMode
Next
        
End Function

Public Function Enum_Display_ModesEX(ComboBox As Object) As Long
On Local Error Resume Next
Dim i  As Long
Dim Count As Long
Dim d3ddm As D3DDISPLAYMODE
Dim sMode As String
    
i = 0
ComboBox.Clear
Count = D3D.GetAdapterModeCount(D3DADAPTER_DEFAULT) - 1
If Err.Number Then Enum_Display_ModesEX = D3DERR_INVALIDCALL: Exit Function
    
For i = 0 To Count
 Call D3D.EnumAdapterModes(D3DADAPTER_DEFAULT, i, d3ddm)
 If Err.Number Then Enum_Display_ModesEX = Err.Number: Exit Function
 sMode = d3ddm.Width & " x "
 sMode = sMode & d3ddm.Height & " x "
  If d3ddm.Format = D3DFMT_X1R5G5B5 Or _
     d3ddm.Format = D3DFMT_R5G6B5 Then
    sMode = sMode & "16"
  ElseIf d3ddm.Format = D3DFMT_X8R8G8B8 Or _
     d3ddm.Format = D3DFMT_A8R8G8B8 Then
   sMode = sMode & "32"
  Else
   sMode = sMode & "Error(BPP)"
  End If
  ComboBox.AddItem sMode
Next
        
End Function

Private Function Reset_Device(hWnd As Long) As Long
If FullS = True Then
 Restore_FullScreen_Device hWnd
Else
 Restore_Window_Device hWnd
End If
End Function

Public Function Initialize_EngineX8(hWnd As Long, Fullscreen As Boolean, Optional Width As Long = 640, Optional Height As Long = 480) As Boolean
 On Local Error GoTo ErrOut
 Dim Mode As D3DDISPLAYMODE
 If DX8 Is Nothing Then Set DX8 = New DirectX8
 If D3DX Is Nothing Then Set D3DX = New D3DX8
 
 CausedAt = "Direct3DCreate"
 Set D3D = DX8.Direct3DCreate()
 If D3D Is Nothing Then Exit Function
 
 D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode
 HasBackground = False
 
 If Fullscreen Then
  CausedAt = "Initialize_FS"
  Initialize_FS hWnd, Width, Height
 Else
  CausedAt = "Initialize_Win"
  Initialize_Win hWnd
 End If
  
 D3DD.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
 D3DD.SetRenderState D3DRS_DITHERENABLE, 1
 D3DD.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
 D3DD.SetRenderState D3DRS_ZENABLE, 1
 D3DD.SetRenderState D3DRS_ZFUNC, D3DCMP_GREATEREQUAL
 
 HoldhWnd = hWnd
 Initialize_EngineX8 = True
 Exit Function
 
ErrOut:
 ErrorToFile "ErrorInit", CausedAt, "N/A"
 Initialize_EngineX8 = False
End Function

Private Sub Cleanup_XEngine8()
 If bShowMouse Then ShowCursor 1
 Call TP_ReleaseAllTexturesFromPool
 Set Background_Picture = Nothing
 Set MapTexture = Nothing
 Set Math = Nothing
 Set xFont = Nothing
 Set D3DD = Nothing
 Set D3D = Nothing
 Set D3DX = Nothing
 Set DX8 = Nothing
End Sub

'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|
'|œœ                     Rendering                          œœ|
'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|

Public Sub Enable_Motion_Blur(Blur_Amount As Single)
 Dim Width As Single
 Dim Height As Single
 Dim rc As RECT
 Dim Viewport As D3DVIEWPORT8
 
 GetClientRect HoldhWnd, rc
 Width = rc.Right
 Height = rc.bottom
 
 If BlurAmount <= 3 Then BlurAmount = 3
 If BlurAmount >= 255 Then BlurAmount = 255
 BlurEnabled = True
 BlurAmount = Blur_Amount
     
 VertList(0).X = 0:     VertList(0).Y = 0
 VertList(1).X = Width: VertList(1).Y = 0
 VertList(2).X = 0:     VertList(2).Y = Height
 VertList(3).X = Width: VertList(3).Y = Height
 VertList(0).RHW = 1: VertList(1).RHW = 1: VertList(2).RHW = 1: VertList(3).RHW = 1
 VertList(0).Color = D3DColorXRGB(255, 255, 255)
 VertList(1).Color = D3DColorXRGB(255, 255, 255)
 VertList(2).Color = D3DColorXRGB(255, 255, 255)
 VertList(3).Color = D3DColorXRGB(255, 255, 255)
 
 If FullS Then
  VertList(0).tU = 0#:           VertList(0).tV = 0#
  VertList(1).tU = Width / 1024: VertList(1).tV = 0#
  VertList(2).tU = 0#:           VertList(2).tV = Height / 1024
  VertList(3).tU = Width / 1024: VertList(3).tV = Height / 1024
 Else
  VertList(0).tU = 0#:          VertList(0).tV = 0#
  VertList(1).tU = Width / 512: VertList(1).tV = 0#
  VertList(2).tU = 0#:          VertList(2).tV = Height / 512
  VertList(3).tU = Width / 512: VertList(3).tV = Height / 512
 End If
 
End Sub

Public Function Start_Engine_Render(Optional Clear_Color As Long)
 On Local Error Resume Next
 If D3DD Is Nothing Then Exit Function
 Dim Res As Long
    
 Res = D3DD.TestCooperativeLevel
    
 If Res = D3DERR_DEVICELOST Then
    Exit Function
 ElseIf Res = D3DERR_DEVICENOTRESET Then
    Res = 0
    Res = Reset_Device(HoldhWnd)
    If Res Then Exit Function
 End If
  
 D3DD.BeginScene
 D3DD.Clear ByVal 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, Clear_Color, 1, 0
End Function

Public Function End_Engine_Render()
 On Local Error Resume Next

 D3DD.EndScene
 D3DD.Present ByVal 0, ByVal 0, 0, ByVal 0
 
 UpdateGlobalTime
 Update_FPS
 
End Function

Public Sub Setup_Default_Rendering_State()
 D3DD.SetRenderState D3DRS_DITHERENABLE, 1
 D3DD.SetRenderState D3DRS_ZENABLE, 1
 D3DD.SetRenderState D3DRS_LIGHTING, 0
End Sub

Public Sub Set_Engines_Shade_Type(Shade_Type As CONST_D3DSHADEMODE)
 On Local Error Resume Next
 D3DD.SetRenderState D3DRS_SHADEMODE, Shade_Type
End Sub

Public Sub Set_Engines_Fill_Type(Fill_Type As CONST_D3DFILLMODE)
 On Local Error Resume Next
 D3DD.SetRenderState D3DRS_FILLMODE, Fill_Type
End Sub

Private Sub UpdateGlobalTime()
 Tick = GetTickCount()
 SystemFrame = Tick - OldSys
 If SystemFrame > 200 Then SystemFrame = 50
 OldSys = Tick
End Sub

'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|
'|œœ                         FPS                            œœ|
'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|

Private Function Update_FPS()
 Static i As Integer
 Static tEnd As Single
 Static tCurrent As Single
 Static HoldFPS As Single

 i = i + 1
 If i = 30 Then
  tCurrent = Timer
    If tCurrent <> tEnd Then
      HoldFPS = 30 / (Timer - tEnd)
      tEnd = Timer
      i = 0
      FPS = HoldFPS
    End If
 End If
End Function

Public Property Get Get_FPS(Optional Format_FPS_String As Boolean = True) As String
 If Format_FPS_String = True Then
  Get_FPS = Format$(FPS, "###.00")
 Else
  Get_FPS = FPS
 End If
End Property

Public Sub Show_FPS_Track(PosX As Integer, PosY As Integer, Red As Single, Green As Single, Blue As Single)
 Dim Output As String
 Output = String$(FPS \ 2, "Ž")
 Draw_Text Output, PosX, PosY, Red, Green, Blue
 DoEvents
End Sub

'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|
'|œœ                     Text Functions                     œœ|
'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|

Public Sub Initialize_Text(Optional Font_Name As String = "Ariel", Optional Font_Size As Single = 8, Optional Bold As Boolean = False, _
                           Optional Underline As Boolean = False, Optional Italic As Boolean = False, Optional Strikethrough As Boolean = False)
 On Local Error Resume Next
 fnt.Name = Font_Name
 fnt.Size = Font_Size
 fnt.Bold = Bold
 fnt.Strikethrough = Strikethrough
 fnt.Underline = Underline
 fnt.Italic = Italic
 Set vbFont = fnt
 Set fnt = Nothing
 Set xFont = D3DX.CreateFont(D3DD, vbFont.hFont)
 TextInitialized = True
End Sub

Public Function Draw_Text(Text_To_Draw As String, PosX As Integer, PosY As Integer, Red As Single, Green As Single, Blue As Single)
 On Local Error Resume Next
 Dim TextRect As RECT
 Dim Screen_Height As Long
 Dim Screen_Width As Long
 Dim col As D3DCOLORVALUE

 If TextInitialized = False Then Exit Function

 col.R = Red: col.G = Green: col.B = Blue: col.A = 1
 Call GetClientRect(HoldhWnd, TextRect)
 TextRect.Left = PosX
 TextRect.Top = PosY
 Screen_Height = TextRect.bottom - TextRect.Top
 Screen_Width = TextRect.Right - TextRect.Left

 xFont.Begin
 D3DX.DrawText xFont, Math.Color4_To_Long(col), Text_To_Draw, TextRect, 0
 xFont.End
End Function

'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|
'|œœ                   Texture Functions                    œœ|
'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|

Public Function Create_Texture(FileName As String) As Direct3DTexture8
 On Local Error Resume Next
 Dim TmpPic As IPictureDisp
 Set TmpPic = LoadPicture(FileName)
  
 Select Case UCase(Right(FileName, 3))
  Case "BMP"
  Case "JPG"
   FileName = Left(FileName, Len(FileName) - 3) & "bmp"
 End Select

 SavePicture TmpPic, FileName
 Set Create_Texture = D3DX.CreateTextureFromFile(D3DD, FileName)
 Set TmpPic = Nothing
End Function

Public Function Create_TextureEX(FileName As String, Optional ColorKey As Long = &HFF000000) As Direct3DTexture8
 On Local Error Resume Next
 Dim Mode As D3DDISPLAYMODE
 Dim TmpPic As IPictureDisp
 Set TmpPic = LoadPicture(FileName)
  
 Select Case UCase(Right(FileName, 3))
  Case "BMP"
  Case "JPG"
   FileName = Left(FileName, Len(FileName) - 3) & "bmp"
 End Select
 
 SavePicture TmpPic, FileName
 D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode

 If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Mode.Format, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) Then
  Set Create_TextureEX = D3DX.CreateTextureFromFileEx(D3DD, FileName, D3DX_DEFAULT, _
                                                      D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                      D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, _
                                                      D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, _
                                                      ColorKey, ByVal 0, ByVal 0)
 ElseIf D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Mode.Format, 0, D3DRTYPE_TEXTURE, D3DFMT_A4R4G4B4) Then
  Set Create_TextureEX = D3DX.CreateTextureFromFileEx(D3DD, FileName, D3DX_DEFAULT, _
                                                      D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                      D3DFMT_A4R4G4B4, D3DPOOL_MANAGED, _
                                                      D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, _
                                                      ColorKey, ByVal 0, ByVal 0)
 ElseIf D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Mode.Format, 0, D3DRTYPE_TEXTURE, D3DFMT_A1R5G5B5) Then
  Set Create_TextureEX = D3DX.CreateTextureFromFileEx(D3DD, FileName, D3DX_DEFAULT, _
                                                      D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                      D3DFMT_A1R5G5B5, D3DPOOL_MANAGED, _
                                                      D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, _
                                                      ColorKey, ByVal 0, ByVal 0)
 Else
  Set Create_TextureEX = D3DX.CreateTextureFromFileEx(D3DD, FileName, D3DX_DEFAULT, _
                                                      D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                      D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                      D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, _
                                                      ColorKey, ByVal 0, ByVal 0)
 End If

 Set TmpPic = Nothing
End Function

Public Function TP_CreateTextureInPool(FileName As String, Optional ColorKey As Long = &HFF000000) As Direct3DTexture8
 On Local Error GoTo ErrOut
 Dim Tex As Direct3DTexture8
    
 If FileName = "" Then Exit Function
 Set Tex = TP_FindTextureInPool(FileName)
 If Not Tex Is Nothing Then
  Set TP_CreateTextureInPool = Tex
  Exit Function
 End If
     
 Set Tex = Create_TextureEX(FileName, ColorKey)
 TP_AddTextureToPool Tex, FileName
 Set TP_CreateTextureInPool = Tex
 Exit Function
    
ErrOut:
 Set TP_CreateTextureInPool = Nothing
End Function

Private Function TP_AddTextureToPool(Texture As Direct3DTexture8, FileName As String)
    
 If TextPool.MaxPool = 0 Then
  ReDim TextPool.TexPool(kGrowSize)
  TextPool.MaxPool = kGrowSize
  TextPool.FirstDel = -1
  Set TextPool.TexPool(0).Texture = Texture
  TextPool.TexPool(0).Name = FileName
  TextPool.TexPool(0).NextDelNode = -1
  Exit Function
 End If
    
 If TextPool.FirstDel > -1 Then
  Set TextPool.TexPool(TextPool.FirstDel).Texture = Texture
  TextPool.TexPool(TextPool.FirstDel).Name = FileName
  TextPool.FirstDel = TextPool.TexPool(TextPool.FirstDel).NextDelNode
  TextPool.TexPool(TextPool.FirstDel).NextDelNode = -1
  Exit Function
 End If
    
 If TextPool.MaxPool <= TextPool.NextEmpty Then
  TextPool.MaxPool = TextPool.MaxPool + kGrowSize
  ReDim Preserve TextPool.TexPool(TextPool.MaxPool)
 End If
    
 Set TextPool.TexPool(TextPool.NextEmpty).Texture = Texture
 TextPool.TexPool(TextPool.NextEmpty).Name = FileName
 TextPool.TexPool(TextPool.NextEmpty).NextDelNode = -1
 TextPool.NextEmpty = TextPool.NextEmpty + 1
    
End Function

Public Function TP_FindTextureInPool(FileName As String) As Direct3DTexture8
 Dim i As Long
 
 For i = 0 To TextPool.NextEmpty - 1
  If (TextPool.TexPool(i).Name = FileName) Then
   Set TP_FindTextureInPool = TextPool.TexPool(i).Texture
   Exit Function
  End If
 Next
 
End Function

Public Sub TP_ReleaseTextureFromPool(FileName As String)
Dim i As Long
 For i = 0 To TextPool.NextEmpty - 1
  If (TextPool.TexPool(i).Name = FileName) Then
   TextPool.TexPool(i).NextDelNode = TextPool.FirstDel
   TextPool.FirstDel = i
   TextPool.TexPool(i).Name = ""
   Set TextPool.TexPool(i).Texture = Nothing
   Exit Sub
  End If
 Next
End Sub

Private Sub TP_ReleaseAllTexturesFromPool()
 ReDim TextPool.TexPool(0)
 TextPool.FirstDel = -1
 TextPool.NextEmpty = 0
 TextPool.MaxPool = 0
End Sub

'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|
'|œœ                     Misc. Functions                    œœ|
'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|

Public Property Get IsFullScreen() As Boolean
 IsFullScreen = FullS
End Property

Public Function CreateGuid() As String
 CreateGuid = DX8.CreateNewGuid
End Function

Private Sub Class_Terminate()
 Unload frmSetup
 Cleanup_XEngine8
End Sub
