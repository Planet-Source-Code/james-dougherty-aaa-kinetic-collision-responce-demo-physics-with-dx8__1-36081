Attribute VB_Name = "BuildDemo"
Option Explicit

Public Sub SetupObjects()
 Dim i As Long   'For the standard loop
 
 'Create Our Background Sprite
 Background.Create_Sprite App.Path & "\Back1.jpg", frmKCR.ScaleWidth, frmKCR.ScaleHeight, 0, 0, 1, True, , , , , vbYellow, , , vbRed
 
 'Create Our Balls
 'We setup the balls as follows
 '   -Random range far from the boundries(so no balls get "stuck")
 '   -Random Velocities
 '   -The balls radius, in this case it is 20
 '   -And the mass, I used a random range
 
 For i = 0 To Num_Balls
  Balls(i).Ball.Create_Sprite App.Path & "\MB.bmp", 30, 30, 0, 0, 1, True, 1.04, 1.04, 0.015, 0.015, vbYellow, vbRed
  Balls(i).Pos = XM.Vector3(RndRange(100, 400), RndRange(100, 400), 1)
  Balls(i).Ball.Position_Sprite Balls(i).Pos.X, Balls(i).Pos.Y, 1
  Balls(i).Vel = XM.Vector3((RndRange(-100, 100) / 15), (RndRange(-100, 100) / 15), 0)
  Balls(i).Radius = 20
  Balls(i).Mass = RndRange(1, 2)
 Next

End Sub

Public Function RndRange(LowerValue As Single, UpperValue As Single) As Integer
 'Computes a random value between to given values
 Randomize
 RndRange = Int((UpperValue - LowerValue + 1) * Rnd + LowerValue)
End Function

Public Sub UpdateBalls()
 Dim i As Long        'For the standard loop
 Dim TotEX As Single  'Helper for computing the total system(Every balls) kinetic energy
 Dim TotEY As Single  'Helper for computing the total system(Every balls) kinetic energy
  
 'Reset Helpers every frame
 TotEX = 0: TotEY = 0
 
 For i = 0 To Num_Balls
  'Update balls position due to its velocity
  Balls(i).Pos.X = Balls(i).Pos.X + Balls(i).Vel.X
  Balls(i).Pos.Y = Balls(i).Pos.Y + Balls(i).Vel.Y
   
  'Update velocity with the wind and gravity
  Balls(i).Vel.X = Balls(i).Vel.X + Wind
  Balls(i).Vel.Y = Balls(i).Vel.Y + Gravity
  
  'Keep the ball from going out of bounds.(X)
  'If the ball hits a boundry, simply reverse its velocity
  If ((Balls(i).Pos.X > Max_X - Balls(i).Radius) Or (Balls(i).Pos.X < Min_X + Balls(i).Radius)) Then
   Balls(i).Vel.X = -Balls(i).Vel.X
   Balls(i).Pos.X = Balls(i).Pos.X + Balls(i).Vel.X
   Balls(i).Pos.Y = Balls(i).Pos.Y + Balls(i).Vel.Y
  End If
  
  'Keep the ball from going out of bounds.(Y)
  'If the ball hits a boundry, simply reverse its velocity
  If ((Balls(i).Pos.Y > Max_Y - Balls(i).Radius) Or (Balls(i).Pos.Y < Min_Y + Balls(i).Radius)) Then
   Balls(i).Vel.Y = -Balls(i).Vel.Y
   Balls(i).Pos.X = Balls(i).Pos.X + Balls(i).Vel.X
   Balls(i).Pos.Y = Balls(i).Pos.Y + Balls(i).Vel.Y
  End If
  
  'Now update the actual balls
  Balls(i).Ball.Position_Sprite Balls(i).Pos.X + 0.5 - Balls(i).Radius, Balls(i).Pos.Y + 0.5 - Balls(i).Radius, 1
  'While keeping them in bound
  KeepInBounds
  
  'The next 3 lines compute the kinetic energy
  TotEX = TotEX + (Balls(i).Vel.X * Balls(i).Vel.X * Balls(i).Radius)
  TotEY = TotEY + (Balls(i).Vel.Y * Balls(i).Vel.Y * Balls(i).Radius)
  TotalEnergy = 0.5 * Sqr(TotEX * TotEX + TotEY * TotEY)
  DoEvents
 Next
 
 'Now we apply our physics to the balls
 ApplyPhysics
 DoEvents
End Sub

Public Sub KeepInBounds()
 Dim i As Long
 
 'If we go out of our boundry, try to throw it back inside
 'the boundry.(X)
 If ((Balls(i).Pos.X > Max_X - Balls(i).Radius) Or (Balls(i).Pos.X < Min_X + Balls(i).Radius)) Then
  Balls(i).Pos = XM.Vector3(RndRange(Min_X, Max_X), RndRange(Min_Y, Max_Y), 1)
 End If
  
 'If we go out of our boundry, try to throw it back inside
 'the boundry.(Y)
 If ((Balls(i).Pos.Y > Max_Y - Balls(i).Radius) Or (Balls(i).Pos.Y < Min_Y + Balls(i).Radius)) Then
  Balls(i).Pos = XM.Vector3(RndRange(Min_X, Max_X), RndRange(Min_Y, Max_Y), 1)
 End If
 DoEvents
 
End Sub

Public Sub ApplyPhysics()
 Dim BallA As Long       'Ball A
 Dim BallB As Long       'Ball B
 Dim nX As Single        'Normal X
 Dim nY As Single        'Normal Y
 Dim SqrnX As Single     'nX * nX For Optimizing
 Dim SqrnY As Single     'nY * nY For Optimizing
 Dim tX As Single        'Tangent X
 Dim tY As Single        'Tangent Y
 Dim Length As Single    'The Length Between The Two Balls
 Dim OpLength As Single  'For Optimizing
 Dim BAIT As Single      'Ball A Initial Tangent
 Dim BAIN As Single      'Ball A Initial Normal
 Dim BAFN As Single      'Ball A Final Normal
 Dim BAFT As Single      'Ball A Final Tangent
 Dim BBIT As Single      'Ball B Initial Tangent
 Dim BBIN As Single      'Ball B Initial Normal
 Dim BBFN As Single      'Ball B Final Normal
 Dim BBFT As Single      'Ball B Final Tangent
 Dim mA As Single        'Ball A Mass
 Dim mB As Single        'Ball B Mass
 Dim CombMass As Single  'Combined Mass Of Both Balls (For Optimizing)
 Dim TranAX As Single    'Translate Back To X Coordinate System(A)
 Dim TranAY As Single    'Translate Back To Y Coordinate System(A)
 Dim TranBX As Single    'Translate Back To X Coordinate System(B)
 Dim TranBY As Single    'Translate Back To Y Coordinate System(B)
 
 'This will actually test not only 2 balls colliding, but it can
 'test if more than 2 balls collide :)
 For BallA = 0 To Num_Balls
  For BallB = BallA + 1 To Num_Balls
  
   If BallA = BallB Then Exit Sub
   'Compute the normal vectr from A->B
   nX = (Balls(BallB).Pos.X - Balls(BallA).Pos.X)
   nY = (Balls(BallB).Pos.Y - Balls(BallA).Pos.Y)
   
   SqrnX = nX * nX
   SqrnY = nY * nY
   Length = Sqr(SqrnX + SqrnY)
   
   'Check for actual collisions
   'If there is a collision compute the collision responce
   OpLength = 2# * ((Balls(BallA).Radius * 0.75) + 0.1)
   If Length <= OpLength Then
    
    'Compute system coordinates and normalize the normal vector
    nX = nX / Length
    nY = nY / Length
    
    'Compute the tangential vector perpendicular from the normal
    tX = -nY
    tY = nX
    
    'Show Tangent and Normal Axis
    'Normal(White)
    DrawLine XM.Vector3(Balls(BallA).Pos.X + 0.5, Balls(BallA).Pos.Y + 0.5, 1), _
             XM.Vector3(Balls(BallA).Pos.X + 20 * nX + 0.5, Balls(BallA).Pos.Y + 20 * nY + 0.5, 1), vbWhite
    
    'Tangent(Yellow)
    DrawLine XM.Vector3(Balls(BallA).Pos.X + 0.5, Balls(BallA).Pos.Y + 0.5, 1), _
             XM.Vector3(Balls(BallA).Pos.X + 20 * tX + 0.5, Balls(BallA).Pos.Y + 20 * tY + 0.5, 1), vbYellow
             
    'Next we compute the initial velocities
    BAIT = (Balls(BallA).Vel.X * tX + Balls(BallA).Vel.Y * tY)
    BAIN = (Balls(BallA).Vel.X * nX + Balls(BallA).Vel.Y * nY)
    BBIT = (Balls(BallB).Vel.X * tX + Balls(BallB).Vel.Y * tY)
    BBIN = (Balls(BallB).Vel.X * nX + Balls(BallB).Vel.Y * nY)
    
    'And get them to get the velocities into terms of n and t axis
    mA = Balls(BallA).Mass
    mB = Balls(BallB).Mass
    CombMass = (mA + mB)
    BAFN = (mB * BBIN * (Coef + 1) + BAIN * (mA - Coef * mB)) / CombMass
    BBFN = (mA * BAIN * (Coef + 1) - BBIN * (mA - Coef * mB)) / CombMass
    'Tangent stays the same
    BAFT = BAIT
    BBFT = BBIT
    
    'Now comes the problem, we are still in system coorinates
    'We need to translate back to X and Y coordinates
    TranAX = BAFN * nX + BAFT * tX
    TranAY = BAFN * nY + BAFT * tY
    TranBX = BBFN * nX + BBFT * tX
    TranBY = BBFN * nY + BBFT * tY
    
    'Store the values
    Balls(BallA).Vel.X = TranAX
    Balls(BallA).Vel.Y = TranAY
    Balls(BallB).Vel.X = TranBX
    Balls(BallB).Vel.Y = TranBY
    
    'Then update the balls
    Balls(BallA).Pos.X = Balls(BallA).Pos.X + Balls(BallA).Vel.X
    Balls(BallA).Pos.Y = Balls(BallA).Pos.Y + Balls(BallA).Vel.Y
    Balls(BallB).Pos.X = Balls(BallB).Pos.X + Balls(BallB).Vel.X
    Balls(BallB).Pos.Y = Balls(BallB).Pos.Y + Balls(BallB).Vel.Y
   End If
   
  Next
 Next

End Sub

Private Sub DrawLine(Pos1 As D3DVECTOR, Pos2 As D3DVECTOR, Color As Long)
 'We are going to use this to draw the tangent and normal lines
 frmKCR.Line (Pos1.X, Pos1.Y)-(Pos2.X, Pos2.Y), Color
End Sub
