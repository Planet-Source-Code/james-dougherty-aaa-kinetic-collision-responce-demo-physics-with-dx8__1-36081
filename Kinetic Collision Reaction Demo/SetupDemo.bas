Attribute VB_Name = "SetupDemo"
Option Explicit 'General rule, MAKE SURE EVERYTHING GETS DEFINED, WE DO NOT WANT VARIABLES THEY ARE SLOOWWWWWWWW!!!!!


'NOTE ABOUT Coefficient Of Restitution
' If the value is 1 the balls maintains const velocity
' If the value is less than 1 it applys friction and heat, therefore the balls looses energy
' If the value is greater than 1 the balls gains energy(which is impossible in real life)


Public Const Min_X As Single = 20     'MinX collision boundry, kind of like frmMain.Left
Public Const Max_X As Single = 700    'MaxX collision boundry, kind of like frmMain.ScaleWidth
Public Const Min_Y As Single = 15     'MinY collision boundry, kind of like frmMain.Top
Public Const Max_Y As Single = 500    'MaxY collision boundry, kind of like frmMain.ScaleHeight

Public Const Num_Balls As Single = 10 'Number of balls
'Public Const Mass As Single = 1       'The mass of the balls
Public Gravity As Single              'Used to hold the gravity value
Public Wind As Single                 'Used to hold the wind value
Public Coef As Single                 'Used to hold the Coefficient Of Restitution

Private Type tBall
 Ball As New XSprite                  'Our main balls surface
 Pos As D3DVECTOR                     'Used to hold the position
 Vel As D3DVECTOR                     'Used to hold the velocity
 Radius As Single                     'Used to hold the radius
 Mass As Single                       'The mass of the balls
End Type: Public Balls(Num_Balls) As tBall 'Our balls.... Hey Now

Public XE As New XEngine              'Main Engine
Public XI As New XInput               'Input Engine
Public XM As New XMath                'Math Engine
Public FPSLimit As New clsFrameLimiter 'Geoffrey Hazen's Class to limit the FPS
Public Background As New XSprite      'The background image
Public TotalEnergy As Single          'Used to hold the systems(all the balls) kinetic energy
Public SlowDown As Boolean            'Use Geoffrey Hazen's Class to limit the FPS
Public ShowHelp As Boolean            'Used to show help (look in rendering loop and frmMain)
Public ShowTrails As Boolean          'Used to show trails of the ball

Public NFO(10) As String              'This is to hold the information we are going
                                      'to print out to the screen at rutime

Public EndLoop As Boolean             'This will stop our rendering loop
Private Declare Function GetInputState Lib "user32" () As Long

Public Sub Init_Engine()
 'Initialize the engine using the setup form
 'If you use a caption (the "") then it will add a border to the
 'form, leaving it "" makes it borderless
 XE.SetupDialog_InitializeAuto frmKCR, ""
 XE.Initialize_Text , 10, True
 XI.Initialize_Input_Engine frmKCR.hWnd
 DoEvents
 
 'Setup initial value to be used
 Coef = 1.001
 Wind = 0
 Gravity = 0.1
 SlowDown = True
 EndLoop = False
 
 'Make something to hold our information thats going to be printed out
 NFO(10) = "Press (H) To Hide Help"
  NFO(9) = "Press (T) To Show Trails"
  NFO(8) = "Press (G) For Ghost Balls"
  NFO(7) = "Press (PageUp) To Increase Gravity"
  NFO(6) = "Press (PageDown) To Decrease Gravity"
  NFO(5) = "Press (Up) To Increase Wind"
  NFO(4) = "Press (Down) To Decrease Wind"
  NFO(3) = "Press (Left) To Increase Coefficient"
  NFO(2) = "Press (Right) To Decrease Coefficient"
  NFO(1) = "Press (S) To Speed Up"
  NFO(0) = "Press (ESC) To Quit"
   
 'CALLED FROM BUILDDEMO
 SetupObjects
    
 'Apply the "smooth" filter
 Set_Bilinear_Filter
 
 'Finally we can render
 Render
End Sub

Private Sub Render()
 Dim i As Long
 
 Do 'Start our render loop
  XE.Start_Engine_Render vbBlack 'Clears the viewport
    
  Background.Render_Sprite 'Render the background picture
  
  UpdateBalls 'Update the balls physics
  For i = 0 To Num_Balls
   Balls(i).Ball.Render_Sprite 'Render all of the balls
  Next
  
  If ShowTrails Then 'Do trail effect by updating the position a little and rendering it agian
   UpdateBalls 'Update the balls physics
   For i = 0 To Num_Balls
    Balls(i).Ball.Render_Sprite 'Render all of the balls
   Next
  End If
  
  'Draw data
  XE.Draw_Text "Kinetic Collision Reaction Demo", frmKCR.ScaleWidth - 230, 15, 1, 1, 1
  XE.Draw_Text XE.Get_FPS & " FPS", frmKCR.ScaleWidth - 155, 30, 1, 1, 1
  XE.Draw_Text "Coefficient Of Restitution - " & Format(Coef, "0.0000"), 20, 15, 1, 1, 1
  XE.Draw_Text "Total Kinetic Energy - " & TotalEnergy, 20, 30, 1, 1, 1
  XE.Draw_Text "Wind - " & Format(Wind, "0.00"), 20, 45, 1, 1, 1
  XE.Draw_Text "Gravity - " & Format(Gravity, "0.00"), 20, 60, 1, 1, 1
    
  'Do you want to display the help?
  'WILL SLOW DOWN THE FPS
  'THE MORE STUFF NO YOU DISPLAY THE SLOWER IT GETS
  If ShowHelp Then
   For i = 1 To 11
    XE.Draw_Text NFO(i - 1), 20, frmKCR.ScaleHeight - ((i * 15) + 13), 1, 1, 1
   Next
  Else
   'If not just display this
   XE.Draw_Text "Press (H) To Display Help", 20, frmKCR.ScaleHeight - 28, 1, 1, 1
  End If
  
  XE.End_Engine_Render 'Display our scene
  GetUserInput
  If SlowDown Then FPSLimit.LimitFrames 30 'Slow down the FPS
  If GetInputState() Then DoEvents 'Let windows do its thing
 Loop Until EndLoop = True 'Keep looping until the user presses the escape key
 
 Cleanup 'After the loop cleanup then end
End Sub

Public Sub GetUserInput()
 Dim i As Long
 
 If XI.Keyboard_KeyState(X_Escape) <> 0 Then EndLoop = True 'Do we quit now?
 If XI.Keyboard_KeyState(X_PageUp) <> 0 Then Gravity = Gravity + 0.01 'increase gravity
 If XI.Keyboard_KeyState(X_PageDown) <> 0 Then Gravity = Gravity - 0.01 'decrease gravity
 If XI.Keyboard_KeyState(X_UP) <> 0 Then Wind = Wind + 0.01 'increase wind
 If XI.Keyboard_KeyState(X_Down) <> 0 Then Wind = Wind - 0.01 'decrease wind
 If XI.Keyboard_KeyState(X_Left) <> 0 Then Coef = Coef + 0.001 'increase Coefficient
 If XI.Keyboard_KeyState(X_Right) <> 0 Then Coef = Coef - 0.001 'decrease Coefficient
   
 'Keep values resonable
 If Coef > 1.001 Then Coef = 1.001
 If Coef < 0 Then Coef = 0
 DoEvents
End Sub


Public Sub Set_Bilinear_Filter()
 On Local Error Resume Next
 'We use this to smooth our scene
 'kinda like anti-aliasing
 With D3DD
  .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
  .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
  .SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_LINEAR
 End With
End Sub

'Deallocate Memory(Clean it, whatever you wanna call it :))
Public Sub Cleanup()
 Dim i As Long
 
 'Release the balls
 For i = 0 To Num_Balls
  Set Balls(i).Ball = Nothing
 Next
 
 Set FPSLimit = Nothing     'Release the frame limiter
 Set Background = Nothing   'Release the background image
 Set XM = Nothing           'Release Math Engine
 Set XI = Nothing           'Release Input Engine
 Set XE = Nothing           'release the main engine
 Unload frmSetup            'make sure even the forms are unloaded
 Unload frmKCR             'make sure even the forms are unloaded
End Sub
