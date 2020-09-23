VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'
  '|¶¶             © 2001-2002 Ariel Productions          ¶¶|'
  '|¶¶                 All Rights Reserved                ¶¶|'
  '|¶¶¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¶¶|'
  '|¶¶             Programmer - James Dougherty           ¶¶|'
  '|¶¶             Source - XSprite.cls                   ¶¶|'
  '|¶¶             Object - UltimaX.dll                   ¶¶|'
  '|¶¶             Version - 2.3                          ¶¶|'
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'

'PSC - I am still NOT done with my new engine
'therefore, there is no commenting
'Sorry for any troubles
'Still needs optimized
Option Explicit

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

Private Type tSprite
 SpriteVerts(3) As TLVERTEX
 Position As D3DVECTOR
 Transparent As Boolean
 Visible As Boolean
 sHeight As Single
 sWidth As Single
End Type: Private Sprite As tSprite

Private Eng As New XEngine
Private Math As New XMath
Private SpriteSurface As Direct3DTexture8
Private HasSurface As Boolean
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef RECT As RECT) As Long

Public Property Get Get_Sprite_Position_X() As Single
 On Local Error Resume Next
 Get_Sprite_Position_X = Sprite.Position.X
End Property

Public Property Get Get_Sprite_Position_Y() As Single
 On Local Error Resume Next
 Get_Sprite_Position_Y = Sprite.Position.Y
End Property

Public Property Get Get_Sprite_Position_Z() As Single
 On Local Error Resume Next
 Get_Sprite_Position_Z = Sprite.Position.Z
End Property

Public Property Get Get_Sprite_PositionEX() As D3DVECTOR
 On Local Error Resume Next
 Get_Sprite_PositionEX = Sprite.Position
End Property

Public Property Get Get_Sprites_Width() As Single
 On Local Error Resume Next
 Get_Sprites_Width = Sprite.sWidth
End Property

Public Property Get Get_Sprites_Height() As Single
 On Local Error Resume Next
 Get_Sprites_Height = Sprite.sHeight
End Property

Public Sub SetVisible(Visible As Boolean)
 Sprite.Visible = Visible
End Sub

Public Function Is_Transparent() As Boolean
 Is_Transparent = Sprite.Transparent
End Function

Public Sub Enable_Transparency(Enabled As Boolean)
 If Enabled Then Sprite.Transparent = True Else Sprite.Transparent = False
End Sub

Public Sub ShadeSprite(NormalColor As Long, HighLightColor As Long)
 
 With Sprite
  .SpriteVerts(0).Color = NormalColor
  .SpriteVerts(1).Color = HighLightColor
  .SpriteVerts(2).Color = NormalColor
  .SpriteVerts(3).Color = HighLightColor
 End With
 
End Sub

Public Sub Position_Sprite(X As Single, Y As Single, Z As Single)
 On Local Error Resume Next
 Dim W, H As Single
 W = Sprite.sWidth: H = Sprite.sHeight
 Sprite.Position = Math.Vector3(X, Y, Z)

 With Sprite
  .SpriteVerts(0).X = X:     .SpriteVerts(0).Y = Y + H: .SpriteVerts(0).Z = Z
  .SpriteVerts(1).X = X:     .SpriteVerts(1).Y = Y:     .SpriteVerts(1).Z = Z
  .SpriteVerts(2).X = X + W: .SpriteVerts(2).Y = Y + H: .SpriteVerts(2).Z = Z
  .SpriteVerts(3).X = X + W: .SpriteVerts(3).Y = Y:     .SpriteVerts(3).Z = Z
 End With

End Sub

Public Sub Scale_Sprite(Scale_X As Single, Scale_Y As Single)
 On Local Error Resume Next
 Dim W As Single, H As Single
 Dim X As Single, Y As Single, Z As Single
 W = Scale_X: H = Scale_Y
 Sprite.sWidth = Scale_X: Sprite.sHeight = Scale_Y
 X = Sprite.Position.X: Y = Sprite.Position.Y: Z = Sprite.Position.Z

 With Sprite
  .SpriteVerts(0).X = X:     .SpriteVerts(0).Y = Y + H: .SpriteVerts(0).Z = Z
  .SpriteVerts(1).X = X:     .SpriteVerts(1).Y = Y:     .SpriteVerts(1).Z = Z
  .SpriteVerts(2).X = X + W: .SpriteVerts(2).Y = Y + H: .SpriteVerts(2).Z = Z
  .SpriteVerts(3).X = X + W: .SpriteVerts(3).Y = Y:     .SpriteVerts(3).Z = Z
 End With

End Sub

Public Function Create_Sprite(FileName As String, Sprite_Width As Single, Sprite_Height As Single, _
                              StartX As Single, StartY As Single, StartZ As Single, Optional UseTexturePool As Boolean = True, _
                              Optional tU As Single = 1, Optional tV As Single = 1, Optional MintU As Single = 0, Optional MintV As Single = 0, _
                              Optional Color1 As Long = vbWhite, Optional Color2 As Long = vbWhite, Optional Color3 As Long = vbWhite, Optional Color4 As Long = vbWhite) As Boolean
 On Local Error GoTo ErrOut
 Create_Sprite = False

 With Sprite
  .SpriteVerts(0).X = StartX:                .SpriteVerts(0).Y = StartY + Sprite_Height: .SpriteVerts(0).Z = StartZ
  .SpriteVerts(1).X = StartX:                .SpriteVerts(1).Y = StartY:                 .SpriteVerts(1).Z = StartZ
  .SpriteVerts(2).X = StartX + Sprite_Width: .SpriteVerts(2).Y = StartY + Sprite_Height: .SpriteVerts(2).Z = StartZ
  .SpriteVerts(3).X = StartX + Sprite_Width: .SpriteVerts(3).Y = StartY:                 .SpriteVerts(3).Z = StartZ
  .SpriteVerts(0).tU = MintU: .SpriteVerts(0).tV = tV:    .SpriteVerts(0).RHW = 1: .SpriteVerts(0).Color = Color1
  .SpriteVerts(1).tU = MintU: .SpriteVerts(1).tV = MintV: .SpriteVerts(1).RHW = 1: .SpriteVerts(1).Color = Color2
  .SpriteVerts(2).tU = tU:    .SpriteVerts(2).tV = tV:    .SpriteVerts(2).RHW = 1: .SpriteVerts(2).Color = Color3
  .SpriteVerts(3).tU = tU:    .SpriteVerts(3).tV = MintV: .SpriteVerts(3).RHW = 1: .SpriteVerts(3).Color = Color4
 
  .Position = Math.Vector3(StartX, StartY, StartZ)
  .sHeight = Sprite_Height: .sWidth = Sprite_Width
 End With

 If FileName <> "" Then
  If UseTexturePool Then
   Set SpriteSurface = Eng.TP_CreateTextureInPool(FileName, &HFF000000)
  Else
   Set SpriteSurface = Eng.Create_TextureEX(FileName, &HFF000000)
  End If
 End If
 
 HasSurface = True
 Create_Sprite = True
 Sprite.Visible = True
 Exit Function

ErrOut:
 Create_Sprite = False
End Function

Public Sub Render_Sprite()
 On Local Error Resume Next
 Dim HoldRender1 As Long
  
 HoldRender1 = D3DD.GetRenderState(D3DRS_CULLMODE)
 If HasSurface = True And Sprite.Visible = True Then
  D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
  D3DD.SetRenderState D3DRS_ALPHATESTENABLE, 1
  D3DD.SetRenderState D3DRS_ALPHAREF, &H8&
  D3DD.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL
  Eng.Direct3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
  
  D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
  If Sprite.Transparent = True Then
   'D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
   D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
  End If

  D3DD.SetVertexShader FVF
  D3DD.SetTexture 0, SpriteSurface
  D3DD.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Sprite.SpriteVerts(0), Len(Sprite.SpriteVerts(0))
  
  D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 0
  D3DD.SetRenderState D3DRS_ALPHATESTENABLE, 0
  D3DD.SetRenderState D3DRS_CULLMODE, HoldRender1
 Else
  Exit Sub
 End If
  
End Sub

Private Sub Cleanup_Sprite_Engine()
 On Local Error Resume Next
 Set SpriteSurface = Nothing
 Set Math = Nothing
 Set Eng = Nothing
End Sub

Private Sub Class_Terminate()
 On Local Error Resume Next
 Cleanup_Sprite_Engine
End Sub
