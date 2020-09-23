VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'
  '|¶¶             © 2001-2002 Ariel Productions          ¶¶|'
  '|¶¶                 All Rights Reserved                ¶¶|'
  '|¶¶¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¶¶|'
  '|¶¶             Programmer - James Dougherty           ¶¶|'
  '|¶¶             Source - XInput.cls                    ¶¶|'
  '|¶¶             Object - UltimaX.dll                   ¶¶|'
  '|¶¶             Version - 2.1                          ¶¶|'
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'

'PSC - I am still NOT done with my new engine
'therefore, there is no commenting
'Sorry for any troubles
'Still needs optimized
Option Explicit

Public Enum XInputKey
 X_1 = 2
 X_2 = 3
 X_3 = 4
 X_4 = 5
 X_5 = 6
 X_6 = 7
 X_7 = 8
 X_8 = 9
 X_9 = 10
 X_0 = 11
 X_A = &H1E
 X_B = &H30
 X_C = &H2E
 X_D = &H20
 X_E = &H12
 X_F = &H21
 X_G = &H22
 X_H = &H23
 X_I = &H17
 X_J = &H24
 X_K = &H25
 X_L = &H26
 X_M = &H32
 X_N = &H31
 X_O = &H18
 X_P = &H19
 X_Q = &H10
 X_R = &H13
 X_S = &H1F
 X_T = &H14
 X_U = &H16
 X_V = &H2F
 X_W = &H11
 X_X = &H2D
 X_Y = &H15
 X_Z = &H2C
 X_F1 = &H3B
 X_F2 = &H3C
 X_F3 = &H3D
 X_F4 = &H3E
 X_F5 = &H3F
 X_F6 = &H40
 X_F7 = &H41
 X_F8 = &H42
 X_F9 = &H43
 X_F10 = &H44
 X_F11 = &H57
 X_F12 = &H58
 X_Num1 = &H4F
 X_Num2 = &H50
 X_Num3 = &H51
 X_Num4 = &H4B
 X_Num5 = &H4C
 X_Num6 = &H4D
 X_Num7 = &H47
 X_Num8 = &H48
 X_Num9 = &H49
 X_Num0 = &H52
 X_NumEnter = &H9C
 X_UP = &HC8
 X_Down = &HD0
 X_Left = &HCB
 X_Right = &HCD
 X_Escape = 1
 X_Enter = &H1C
 X_LShift = &H2A
 X_RShift = &H36
 X_LControl = &H1D
 X_RControl = &H9D
 X_Space = &H39
 X_Insert = &HD2
 X_Delete = &HD3
 X_Home = &HC7
 X_End = &HCF
 X_PageUp = &HC9
 X_PageDown = &HD1
 X_BackSpace = 14
 X_Add = &H4E
 X_Subtract = &H4A
 X_Period = &H34
 X_Tab = 15
End Enum

Public Enum FXDirection
 North = 0
 North_East = 1
 East = 2
 South_East = 3
 South = 4
 South_West = 5
 West = 6
 North_West = 7
End Enum

Public Enum FXType
 ShootArrow = 0
 Gun_44_Magnum = 1
 Gun_9MM = 2
 BB_Gun = 3
 Gatling_Gun = 4
End Enum

Private Type POINTAPI
 X As Long
 Y As Long
End Type

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

Private DX As New DirectX8
Private DI As DirectInput8
Private D3DX As D3DX8
Private D3D As Direct3D8
Private DIEnum As DirectInputEnumDevices8
Private DIDev As DirectInputDevice8
Private DIState As DIKEYBOARDSTATE
Private DIMouse As DirectInputDevice8
Private DIMState As DIMOUSESTATE
Private MouseVerts(3) As TLVERTEX
Private MouseSurface As Direct3DTexture8
Private CustMouse As Boolean
Private DIJoy As DirectInputDevice8
Private DIJOYSTATE As DIJOYSTATE
Private DIEffEnum As DirectInputEnumEffects
Private DIFFEffect() As DirectInputEffect
Private DIEffectType As Long
Private DIFFStaticParams As Long
Private DIJoyPropRange As DIPROPRANGE
Private DIJoyPropDead As DIPROPLONG
Private DIJoyPropSaturation As DIPROPLONG
Private EffectParams() As Long
Private AxisPresent(1 To 8) As Boolean
Private LastEffect As Long
Private FFEffect As DIEFFECT
Private IsJoystick As Boolean
Private KeyboardCaps As DIDEVCAPS
Private MouseCaps As DIDEVCAPS
Private JoyCaps As DIDEVCAPS
Private WasEnumerated As Boolean
Private mhWnd As Long
Private bShowMouse As Boolean
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Function XInput8() As DirectInput8
 Set XInput8 = DI
End Function

Public Function Initialize_Input_Engine(hWnd As Long) As Boolean
 On Local Error GoTo ErrOut
 bShowMouse = False
 
 CausedAt = "Create 3D/X8"
 If D3DX Is Nothing Then Set D3DX = New D3DX8
 If D3D Is Nothing Then Set D3D = DX.Direct3DCreate
 
 CausedAt = "DirectInputCreate"
 Set DI = DX.DirectInputCreate()
 
 CausedAt = "Create Keyboard"
 Set DIDev = DI.CreateDevice("GUID_SysKeyboard")
 DIDev.SetCommonDataFormat DIFORMAT_KEYBOARD
 DIDev.SetCooperativeLevel hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
 DIDev.GetCapabilities KeyboardCaps
 DIDev.Acquire
 
 CausedAt = "Create Mouse"
 Set DIMouse = DI.CreateDevice("GUID_SysMouse")
 DIMouse.SetCommonDataFormat DIFORMAT_MOUSE
 DIMouse.SetCooperativeLevel hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
 DIMouse.GetCapabilities MouseCaps
 DIMouse.Acquire
 
 CausedAt = "Create Joystick"
 Set DIEnum = DI.GetDIDevices(DI8DEVCLASS_GAMECTRL, DIEDFL_ATTACHEDONLY)
 If DIEnum.GetCount > 0 Then
  Set DIJoy = DI.CreateDevice(DIEnum.GetItem(1).GetGuidInstance)
  DIJoy.SetCommonDataFormat DIFORMAT_JOYSTICK
  DIJoy.SetCooperativeLevel hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
  SetDefaultJoystickProperties
  IdentifyAxes
  DIJoy.GetCapabilities JoyCaps
  DIJoy.Acquire
  IsJoystick = True
 End If
 
 mhWnd = hWnd
 Initialize_Input_Engine = True
 Exit Function
 
ErrOut:
 ErrorToFile "Error - InitDI", CausedAt, "N/A"
 Initialize_Input_Engine = False
End Function

Private Sub Cleanup_Input_Engine()
 If bShowMouse Then ShowCursor 1
 DIDev.Unacquire
 DIMouse.Unacquire
 If IsJoystick Then DIJoy.Unacquire: Set DIJoy = Nothing
 Set MouseSurface = Nothing
 Set DIDev = Nothing
 Set DIMouse = Nothing
 Set D3D = Nothing
 Set D3DX = Nothing
 Set DI = Nothing
 Set DX = Nothing
End Sub

Private Sub Class_Terminate()
 Cleanup_Input_Engine
End Sub

'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|
'|ºº                     Keyboard                           ºº|
'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|

Public Function Keyboard_KeyState(Key As XInputKey) As Boolean
 On Local Error Resume Next
 Keyboard_KeyState = DIState.Key(Key)
 DIDev.GetDeviceStateKeyboard DIState
End Function

Public Sub Keyboard_RunControlPanel(hWnd As Long)
 On Local Error Resume Next
 DIDev.RunControlPanel hWnd
End Sub

'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|
'|ºº                       Mouse                            ºº|
'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|

Public Sub Mouse_RunControlPanel(hWnd As Long)
 On Local Error Resume Next
 DIMouse.RunControlPanel hWnd
End Sub

Public Sub Mouse_Update(Direct3DDevice As Direct3DDevice8)
 On Local Error GoTo ErrOut
 Dim PT As POINTAPI
 
 CausedAt = "Getting Cursor Position"
 GetCursorPos PT
 CausedAt = "Getting Mouse State"
 DIMouse.GetDeviceStateMouse DIMState
 
 If CustMouse = True Then
  CausedAt = "Setting Mouse Vertices"
  MouseVerts(0).X = (PT.X - 54):      MouseVerts(0).Y = (PT.Y - 73) + 32: MouseVerts(0).Z = 0
  MouseVerts(1).X = (PT.X - 54):      MouseVerts(1).Y = (PT.Y - 73):      MouseVerts(1).Z = 0
  MouseVerts(2).X = (PT.X - 54) + 32: MouseVerts(2).Y = (PT.Y - 73) + 32: MouseVerts(2).Z = 0
  MouseVerts(3).X = (PT.X - 54) + 32: MouseVerts(3).Y = (PT.Y - 73):      MouseVerts(3).Z = 0
  CausedAt = "Setting Alphablendable"
  Direct3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
  CausedAt = "Blending"
  Direct3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
  Direct3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
  CausedAt = "SetVertexShader"
  Direct3DDevice.SetVertexShader FVF
  Direct3DDevice.SetTexture 0, MouseSurface
  CausedAt = "Rendering Mouse Cursor"
  Direct3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MouseVerts(0), Len(MouseVerts(0))
  Direct3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
 End If
 
 Exit Sub
 
ErrOut:
 ErrorToFile "Error - MouseUpdate", CausedAt, "N/A"
End Sub

Public Function Mouse_InputX() As Long
 On Local Error Resume Next
 Mouse_InputX = DIMState.lX
End Function

Public Function Mouse_InputY() As Long
 On Local Error Resume Next
 Mouse_InputY = DIMState.lY
End Function

Public Function Mouse_InputZ() As Long
 On Local Error Resume Next
 Mouse_InputZ = DIMState.lZ
End Function

Public Function Mouse_LeftClick() As Long
 On Local Error Resume Next
 Mouse_LeftClick = DIMState.Buttons(0)
End Function

Public Function Mouse_RightClick() As Long
 On Local Error Resume Next
 Mouse_RightClick = DIMState.Buttons(1)
End Function

Public Function Mouse_WheelClick() As Long
 On Local Error Resume Next
 Mouse_WheelClick = DIMState.Buttons(2)
End Function

Public Function Mouse_CursorVisible(Visible As Boolean)
 If Visible = False Then ShowCursor 0: bShowMouse = True
End Function

Public Function Mouse_Custom_Cursor(FileName As String, Direct3DDevice As Direct3DDevice8) As Boolean
 On Local Error GoTo ErrOut
 Mouse_Custom_Cursor = False

 MouseVerts(0).X = 0:  MouseVerts(0).Y = 32: MouseVerts(0).Z = 0
 MouseVerts(1).X = 0:  MouseVerts(1).Y = 0:  MouseVerts(1).Z = 0
 MouseVerts(2).X = 32: MouseVerts(2).Y = 32: MouseVerts(2).Z = 0
 MouseVerts(3).X = 32: MouseVerts(3).Y = 0:  MouseVerts(3).Z = 0
 MouseVerts(0).tU = 0: MouseVerts(0).tV = 1: MouseVerts(0).RHW = 1: MouseVerts(0).Color = &HFFFFFF
 MouseVerts(1).tU = 0: MouseVerts(1).tV = 0: MouseVerts(1).RHW = 1: MouseVerts(1).Color = &HFFFFFF
 MouseVerts(2).tU = 1: MouseVerts(2).tV = 1: MouseVerts(2).RHW = 1: MouseVerts(2).Color = &HFFFFFF
 MouseVerts(3).tU = 1: MouseVerts(3).tV = 0: MouseVerts(3).RHW = 1: MouseVerts(3).Color = &HFFFFFF

 If FileName <> "" Then
  Set MouseSurface = Create_TextureEX(Direct3DDevice, FileName)
 End If
 
 ShowCursor 0: bShowMouse = True
 Mouse_Custom_Cursor = True
 CustMouse = True
 Exit Function

ErrOut:
 Mouse_Custom_Cursor = False
End Function

'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|
'|ºº                      JoyStick                          ºº|
'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|

Public Sub Joystick_RunControlPanel(hWnd As Long)
 On Local Error Resume Next
 If Not IsJoystick Then Exit Sub
 DIJoy.RunControlPanel hWnd
End Sub

Public Function Joystick_GetDriverVersion() As Long
 On Local Error Resume Next
 If Not IsJoystick Then Exit Function
 Joystick_GetDriverVersion = JoyCaps.lDriverVersion
End Function

Public Function Joystick_GetFirmwareRevision() As Long
 On Local Error Resume Next
 If Not IsJoystick Then Exit Function
 Joystick_GetFirmwareRevision = JoyCaps.lFirmwareRevision
End Function

Public Function Joystick_GetHardwareRevision() As Long
 On Local Error Resume Next
 If Not IsJoystick Then Exit Function
 Joystick_GetHardwareRevision = JoyCaps.lHardwareRevision
End Function

Public Function Joystick_HasJoystick() As Boolean
Joystick_HasJoystick = IsJoystick
End Function

Public Function Joystick_HasForceFeedback() As Boolean
Set DIEnum = Nothing
Set DIEnum = DI.GetDIDevices(DI8DEVCLASS_GAMECTRL, DIEDFL_ATTACHEDONLY Or DIEDFL_FORCEFEEDBACK)
If DIEnum.GetCount = 0 Then
 Joystick_HasForceFeedback = False
Else
 Joystick_HasForceFeedback = True
End If
End Function

Private Sub SetDefaultJoystickProperties()
On Local Error Resume Next

With DIJoyPropDead
 .lData = 1000
 .lHow = DIPH_BYOFFSET
 .lObj = DIJOFS_X
 DIJoy.SetProperty "DIPROP_DEADZONE", DIJoyPropDead
 .lObj = DIJOFS_Y
 DIJoy.SetProperty "DIPROP_DEADZONE", DIJoyPropDead
End With
    
With DIJoyPropSaturation
 .lData = 9500
 .lHow = DIPH_BYOFFSET
 .lObj = DIJOFS_X
 DIJoy.SetProperty "DIPROP_SATURATION", DIJoyPropSaturation
 .lObj = DIJOFS_Y
 DIJoy.SetProperty "DIPROP_SATURATION", DIJoyPropSaturation
End With
    
With DIJoyPropRange
 .lHow = DIPH_DEVICE
 .lMin = 0
 .lMax = 10000
 DIJoy.SetProperty "DIPROP_RANGE", DIJoyPropRange
End With

End Sub

Public Function Joystick_EnumJoysticks(ListBox As Object) As Long
On Local Error Resume Next
If Not IsJoystick Then Exit Function
Dim i As Integer

Set DIEnum = Nothing
Set DIEnum = DI.GetDIDevices(DI8DEVCLASS_GAMECTRL, DIEDFL_ATTACHEDONLY)
If DIEnum.GetCount = 0 Then
 ListBox.AddItem "No avialable devices attached"
 Joystick_EnumJoysticks = 0
 Exit Function
Else
 For i = 1 To DIEnum.GetCount
  Call ListBox.AddItem(DIEnum.GetItem(i).GetInstanceName)
 Next
 Joystick_EnumJoysticks = 1
End If

End Function

Public Sub Joystick_EnumEffects(ListBox As Object)
On Local Error Resume Next
If Not IsJoystick Then Exit Sub
Dim i As Integer
Set DIEffEnum = DIJoy.GetEffectsEnum(DIEFT_ALL)

For i = 1 To DIEffEnum.GetCount
 DIEffectType = DIEffEnum.GetType(i) And &HFF
 DIFFStaticParams = DIEffEnum.GetStaticParams(i)
                                                   
 If (DIEffectType = DIEFT_HARDWARE) And (DIFFStaticParams And DIEP_TYPESPECIFICPARAMS) <> 0 Then
  GoTo Ignore
 ElseIf DIEffectType = DIEFT_CUSTOMFORCE Then
  GoTo Ignore
 End If

 ListBox.AddItem DIEffEnum.GetName(i)
 ReDim Preserve EffectParams(ListBox.ListCount - 1)
 EffectParams(ListBox.ListCount - 1) = DIEffectType
 ReDim Preserve DIFFEffect(ListBox.ListCount - 1)
 On Local Error GoTo ErrorHandler
 Set DIFFEffect(UBound(DIFFEffect)) = DIJoy.CreateEffect(DIEffEnum.GetEffectGuid(i), CreateDefaultFX(i))
 DIFFEffect(UBound(DIFFEffect)).Unload
Ignore:
Next
WasEnumerated = True
Exit Sub

ErrorHandler:
    
If Err.Number = 5 Then
 ListBox.RemoveItem ListBox.ListCount - 1
 ReDim Preserve DIFFEffect(ListBox.ListCount - 1)
 Resume Next
ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
 DIJoy.Unacquire
 DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
 DIJoy.Acquire
 Resume
End If

End Sub

Private Function CreateDefaultFX(Index As Integer) As DIEFFECT
    
With CreateDefaultFX
 .lDuration = &HFFFFFFFF
 .lGain = 10000
 .lSamplePeriod = 0
 .lTriggerButton = DIEB_NOTRIGGER
 .lTriggerRepeatInterval = -1
 .constantForce.lMagnitude = 10000
 .rampForce.lRangeStart = 0
 .rampForce.lRangeEnd = 0
 .conditionFlags = DICONDITION_USE_BOTH_AXES
 With .conditionX
  .lDeadBand = 0
  .lNegativeSaturation = 10000
  .lOffset = 0
  .lPositiveSaturation = 10000
 End With
 With .conditionY
  .lDeadBand = 0
  .lNegativeSaturation = 10000
  .lOffset = 0
  .lPositiveSaturation = 10000
 End With
 With .periodicForce
  .lMagnitude = 10000
  .lOffset = 0
  .lPeriod = 1
  .lPhase = 0
 End With
End With
    
End Function

Private Sub IdentifyAxes()
Dim didoEnum As DirectInputEnumDeviceObjects
Dim dido As DirectInputDeviceObjectInstance
Dim sGuid As String
Dim i As Integer
   
For i = 1 To 8
 AxisPresent(i) = False
Next
   
Set didoEnum = DIJoy.GetDeviceObjectsEnum(DIDFT_AXIS)
For i = 1 To didoEnum.GetCount
 Set dido = didoEnum.GetItem(i)
 sGuid = dido.GetGuidType
 Select Case sGuid
  Case "GUID_XAxis"
   AxisPresent(1) = True
  Case "GUID_YAxis"
   AxisPresent(2) = True
  Case "GUID_ZAxis"
   AxisPresent(3) = True
  Case "GUID_RxAxis"
   AxisPresent(4) = True
  Case "GUID_RyAxis"
   AxisPresent(5) = True
  Case "GUID_RzAxis"
   AxisPresent(6) = True
  Case "GUID_Slider"
   AxisPresent(7) = True
   AxisPresent(8) = True
 End Select
Next
End Sub

Private Sub CheckLoss()
 On Local Error Resume Next
 If Err.Number = DIERR_NOTACQUIRED Or Err.Number = DIERR_INPUTLOST Then
  DIJoy.Acquire
  Exit Sub
 End If
End Sub

Public Sub Joystick_Update()
 On Local Error Resume Next
 If Not IsJoystick Then Exit Sub
 DIJoy.Poll
 CheckLoss
End Sub

Public Function Joystick_NumberOfButtons() As Long
 On Local Error Resume Next
 If Not IsJoystick Then Exit Function
 Joystick_NumberOfButtons = JoyCaps.lButtons
End Function

Public Function Joystick_NumberOfAxis() As Long
 On Local Error Resume Next
 If Not IsJoystick Then Exit Function
 Joystick_NumberOfAxis = JoyCaps.lAxes
End Function

Public Function Joystick_NumberOfPOVs() As Long
 On Local Error Resume Next
 If Not IsJoystick Then Exit Function
 Joystick_NumberOfPOVs = JoyCaps.lPOVs
End Function

Public Function Joystick_Button(Button As Long) As Long
 On Local Error Resume Next
 If Not IsJoystick Then Exit Function
 If Button > Joystick_NumberOfButtons Then Exit Function
 If Button <= 0 Then Exit Function
 Select Case DIJOYSTATE.Buttons(Button - 1)
  Case 0
   Joystick_Button = 0
  Case Else
   Joystick_Button = 1
 End Select
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
End Function

Public Function Joystick_POV() As Long
 On Local Error Resume Next
 If Not IsJoystick Then Exit Function
 Dim i As Long
 For i = 0 To JoyCaps.lPOVs - 1
  Joystick_POV = DIJOYSTATE.POV(i)
 Next
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Function Joystick_XAxis() As Long
 On Local Error Resume Next
 If Not IsJoystick Then Exit Function
 If Not AxisPresent(1) Then Exit Function
 Joystick_XAxis = DIJOYSTATE.X
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Function Joystick_YAxis() As Long
 On Local Error Resume Next
 If Not IsJoystick Then Exit Function
 If Not AxisPresent(2) Then Exit Function
 Joystick_YAxis = DIJOYSTATE.Y
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Function Joystick_ZAxis() As Long
 On Local Error Resume Next
 If Not IsJoystick Then Exit Function
 If Not AxisPresent(3) Then Exit Function
 Joystick_ZAxis = DIJOYSTATE.Z
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Function Joystick_RotXAxis() As Long
 On Local Error Resume Next
 If Not IsJoystick Then Exit Function
 If Not AxisPresent(4) Then Exit Function
 Joystick_RotXAxis = DIJOYSTATE.rx
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Function Joystick_RotYAxis() As Long
 On Local Error Resume Next
 If Not IsJoystick Then Exit Function
 If Not AxisPresent(5) Then Exit Function
 Joystick_RotYAxis = DIJOYSTATE.ry
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Function Joystick_RotZAxis() As Long
 On Local Error Resume Next
 If Not IsJoystick Then Exit Function
 If Not AxisPresent(6) Then Exit Function
 Joystick_RotZAxis = DIJOYSTATE.rz
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Function Joystick_Slider0() As Long
 On Local Error Resume Next
 If Not IsJoystick Then Exit Function
 If Not AxisPresent(7) Then Exit Function
 Joystick_Slider0 = DIJOYSTATE.slider(0)
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Function Joystick_Slider1() As Long
 On Local Error Resume Next
 If Not IsJoystick Then Exit Function
 If Not AxisPresent(8) Then Exit Function
 Joystick_Slider1 = DIJOYSTATE.slider(1)
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Sub Joystick_TurnOffAutocenter()
On Local Error Resume Next
If Not IsJoystick Then Exit Sub
If Not Joystick_HasForceFeedback Then Exit Sub
Dim Prop As DIPROPLONG

Prop.lData = 0
Prop.lHow = DIPH_DEVICE
Prop.lObj = 0
DIJoy.SetProperty "DIPROP_AUTOCENTER", Prop
End Sub

Public Sub Joystick_SetFXStart(FXIndex As Long)
On Local Error Resume Next
If Not IsJoystick Then Exit Sub
If Not Joystick_HasForceFeedback Then Exit Sub
If Not WasEnumerated Then Exit Sub
DIFFEffect(FXIndex).Start 1, 0
End Sub

Public Sub Joystick_SetFXUnload(FXIndex As Long)
On Local Error Resume Next
If Not IsJoystick Then Exit Sub
If Not Joystick_HasForceFeedback Then Exit Sub
If Not WasEnumerated Then Exit Sub
DIFFEffect(FXIndex).Unload
End Sub

Public Sub Joystick_SetFXStop(FXIndex As Long)
On Local Error Resume Next
If Not IsJoystick Then Exit Sub
If Not Joystick_HasForceFeedback Then Exit Sub
If Not WasEnumerated Then Exit Sub
DIFFEffect(FXIndex).Stop
End Sub

Public Sub Joystick_SetFXEnvelopeEffect(FXIndex As Long, AttackLevel As Long, AttackTime As Long, _
                                      FadeLevel As Long, FadeTime As Long)
On Local Error GoTo ErrorHandler
If Not IsJoystick Then Exit Sub
If Not Joystick_HasForceFeedback Then Exit Sub
If Not WasEnumerated Then Exit Sub

FFEffect.bUseEnvelope = True
With FFEffect.envelope
 .lAttackLevel = AttackLevel
 .lAttackTime = AttackTime
 .lFadeLevel = FadeLevel
 .lFadeTime = FadeTime
End With
DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_ENVELOPE
Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXDuration(FXIndex As Long, Duration As Long, Optional Infinite As Boolean = False)
On Local Error GoTo ErrorHandler
If Not IsJoystick Then Exit Sub
If Not Joystick_HasForceFeedback Then Exit Sub
If Not WasEnumerated Then Exit Sub

If Infinite Then
 FFEffect.lDuration = -1
Else
 FFEffect.lDuration = Duration * 100
End If
DIFFEffect(FXIndex).Stop
DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_DURATION
DIFFEffect(FXIndex).Start 1, 0
Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXGain(FXIndex As Long, Gain As Long)
On Local Error GoTo ErrorHandler
If Not IsJoystick Then Exit Sub
If Not Joystick_HasForceFeedback Then Exit Sub
If Not WasEnumerated Then Exit Sub

FFEffect.lGain = Gain
DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_GAIN
Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXSampleRate(FXIndex As Long, Rate As Long)
On Local Error GoTo ErrorHandler
If Not IsJoystick Then Exit Sub
If Not Joystick_HasForceFeedback Then Exit Sub
If Not WasEnumerated Then Exit Sub

FFEffect.lSamplePeriod = Rate
DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_SAMPLEPERIOD
Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXConstantForce(FXIndex As Long, Force As Long)
On Local Error GoTo ErrorHandler
If Not IsJoystick Then Exit Sub
If Not Joystick_HasForceFeedback Then Exit Sub
If Not WasEnumerated Then Exit Sub

FFEffect.constantForce.lMagnitude = Force
DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_TYPESPECIFICPARAMS
Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXDirection(FXIndex As Long, Direction As FXDirection)
On Local Error GoTo ErrorHandler
If Not IsJoystick Then Exit Sub
If Not Joystick_HasForceFeedback Then Exit Sub
If Not WasEnumerated Then Exit Sub

With FFEffect
 If Direction = North Then
  .X = 0
 ElseIf Direction = North_East Then
  .X = 4500
 ElseIf Direction = East Then
  .X = 9000
 ElseIf Direction = South_East Then
  .X = 13500
 ElseIf Direction = South Then
  .X = 18000
 ElseIf Direction = South_West Then
  .X = 22500
 ElseIf Direction = West Then
  .X = 27000
 ElseIf Direction = North_West Then
  .X = 31500
 End If
End With

DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_DIRECTION
Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXConditionX(FXIndex As Long, DeadBand As Long, _
                                    NegCoeff As Long, NegSat As Long, _
                                    PosCoeff As Long, PosSat As Long, Offset As Long)
On Local Error GoTo ErrorHandler
If Not IsJoystick Then Exit Sub
If Not Joystick_HasForceFeedback Then Exit Sub
If Not WasEnumerated Then Exit Sub

With FFEffect.conditionX
 .lDeadBand = DeadBand
 .lNegativeCoefficient = NegCoeff
 .lNegativeSaturation = NegSat
 .lOffset = Offset
 .lPositiveCoefficient = PosCoeff
 .lPositiveSaturation = PosSat
End With
DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_TYPESPECIFICPARAMS
Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXConditionY(FXIndex As Long, DeadBand As Long, _
                                    NegCoeff As Long, NegSat As Long, _
                                    PosCoeff As Long, PosSat As Long, Offset As Long)
On Local Error GoTo ErrorHandler
If Not IsJoystick Then Exit Sub
If Not Joystick_HasForceFeedback Then Exit Sub
If Not WasEnumerated Then Exit Sub

With FFEffect.conditionY
 .lDeadBand = DeadBand
 .lNegativeCoefficient = NegCoeff
 .lNegativeSaturation = NegSat
 .lOffset = Offset
 .lPositiveCoefficient = PosCoeff
 .lPositiveSaturation = PosSat
End With
DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_TYPESPECIFICPARAMS
Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXRampForce(FXIndex As Long, StartRange As Long, EndRange As Long)
On Local Error GoTo ErrorHandler
If Not IsJoystick Then Exit Sub
If Not Joystick_HasForceFeedback Then Exit Sub
If Not WasEnumerated Then Exit Sub

With FFEffect.rampForce
 .lRangeStart = StartRange
 .lRangeEnd = EndRange
End With
DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_TYPESPECIFICPARAMS
Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXPeriodicForce(FXIndex As Long, Magnitude As Long, Offset As Long, Period As Long, Phase As Long)
On Local Error GoTo ErrorHandler
If Not IsJoystick Then Exit Sub
If Not Joystick_HasForceFeedback Then Exit Sub
If Not WasEnumerated Then Exit Sub

With FFEffect.periodicForce
 .lMagnitude = Magnitude
 .lOffset = Offset
 .lPeriod = Period
 .lPhase = Phase
End With
DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_TYPESPECIFICPARAMS
Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_PlayPredefinedFX(Effect As FXType)
On Local Error GoTo ErrorHandler
If Not IsJoystick Then Exit Sub
If Not Joystick_HasForceFeedback Then Exit Sub
If Not WasEnumerated Then Exit Sub
Dim HIndex As Long

If Effect = ShootArrow Then
 Joystick_SetFXConstantForce 0, 7000
 Joystick_SetFXDirection 0, North
 Joystick_SetFXDuration 0, 1000
 HIndex = 0
ElseIf Effect = Gun_44_Magnum Then
 Joystick_SetFXDirection 0, South
 Joystick_SetFXConstantForce 0, 10000
 Joystick_SetFXDuration 0, 1500
 Joystick_SetFXDirection 0, North
 Joystick_SetFXConstantForce 0, 8000
 Joystick_SetFXDuration 0, 1000
 HIndex = 0
ElseIf Effect = Gun_9MM Then
 Joystick_SetFXDirection 0, South
 Joystick_SetFXConstantForce 0, 5000
 Joystick_SetFXDuration 0, 1000
 Joystick_SetFXDirection 0, North
 Joystick_SetFXConstantForce 0, 5000
 Joystick_SetFXDuration 0, 500
 HIndex = 0
ElseIf Effect = BB_Gun Then
 Joystick_SetFXDirection 0, South
 Joystick_SetFXConstantForce 0, 2000
 Joystick_SetFXDuration 0, 500
 Joystick_SetFXDirection 0, North
 Joystick_SetFXConstantForce 0, 2000
 Joystick_SetFXDuration 0, 500
 HIndex = 0
ElseIf Effect = Gatling_Gun Then
 Joystick_SetFXDirection 2, North
 Joystick_SetFXPeriodicForce 2, 10000, 0, 40000, 0
 Joystick_SetFXDuration 2, 5000
 HIndex = 2
End If
Joystick_SetFXUnload HIndex
Joystick_SetFXStart HIndex
DoEvents
Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Private Function Create_TextureEX(Direct3DDevice As Direct3DDevice8, FileName As String, Optional ColorKey As Long = &HFF000000) As Direct3DTexture8
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
  Set Create_TextureEX = D3DX.CreateTextureFromFileEx(Direct3DDevice, FileName, D3DX_DEFAULT, _
                                                      D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                      D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, _
                                                      D3DX_FILTER_POINT, D3DX_FILTER_POINT, _
                                                      ColorKey, ByVal 0, ByVal 0)
 ElseIf D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Mode.Format, 0, D3DRTYPE_TEXTURE, D3DFMT_A4R4G4B4) Then
  Set Create_TextureEX = D3DX.CreateTextureFromFileEx(Direct3DDevice, FileName, D3DX_DEFAULT, _
                                                      D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                      D3DFMT_A4R4G4B4, D3DPOOL_MANAGED, _
                                                      D3DX_FILTER_POINT, D3DX_FILTER_POINT, _
                                                      ColorKey, ByVal 0, ByVal 0)
 ElseIf D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Mode.Format, 0, D3DRTYPE_TEXTURE, D3DFMT_A1R5G5B5) Then
  Set Create_TextureEX = D3DX.CreateTextureFromFileEx(Direct3DDevice, FileName, D3DX_DEFAULT, _
                                                      D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                      D3DFMT_A1R5G5B5, D3DPOOL_MANAGED, _
                                                      D3DX_FILTER_POINT, D3DX_FILTER_POINT, _
                                                      ColorKey, ByVal 0, ByVal 0)
 Else
  Set Create_TextureEX = D3DX.CreateTextureFromFileEx(Direct3DDevice, FileName, D3DX_DEFAULT, _
                                                      D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                      D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                      D3DX_FILTER_POINT, D3DX_FILTER_POINT, _
                                                      ColorKey, ByVal 0, ByVal 0)
 End If

 Set TmpPic = Nothing
End Function

