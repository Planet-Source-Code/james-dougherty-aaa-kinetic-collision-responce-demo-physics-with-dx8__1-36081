VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'
  '|¶¶             © 2001-2002 Ariel Productions          ¶¶|'
  '|¶¶                 All Rights Reserved                ¶¶|'
  '|¶¶¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¶¶|'
  '|¶¶             Programmer - James Dougherty           ¶¶|'
  '|¶¶             Source - XMath.cls                     ¶¶|'
  '|¶¶             Object - UltimaX.dll                   ¶¶|'
  '|¶¶             Version - 2.1                          ¶¶|'
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'

'PSC - I am still NOT done with my new engine
'therefore, there is no commenting
'Sorry for any troubles
'Still needs optimized
Option Explicit

Private D3DX As New D3DX8
Private FSys As New FileSystemObject
Private OutStream As TextStream
Private InputStream As TextStream

Private Const mPI As Single = 3.14159265358979
Private Const Inv_PI As Single = 0.318309886183791
Private Const g_90d As Single = 1.5707963267949
Private Const g_180d As Single = 3.14159265358979
Private Const g_270d As Single = 4.71238898038469
Private Const g_360d As Single = 6.28318530717959
Private Const mRadiant As Single = 1.74532925199433E-02
Private Const DegToRad As Single = 1.74532925199433E-02
Private Const RadToDeg As Single = 57.2957795130823
Private Const mHugeSingle As Single = 1E+38
Private Const mEPSILON As Single = 0.0001

'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|
'|þþ                    Math Constants                       þþ|
'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|

Public Function PI() As Single
 PI = mPI
End Function

Public Function Inverse_PI() As Single
 Inverse_PI = Inv_PI
End Function

Public Function Degrees_90() As Single
 Degrees_90 = g_90d
End Function

Public Function Degrees_180() As Single
 Degrees_180 = g_180d
End Function

Public Function Degrees_270() As Single
 Degrees_270 = g_270d
End Function

Public Function Degrees_360() As Single
 Degrees_360 = g_360d
End Function

Public Function Epsilon() As Single
 Epsilon = mEPSILON
End Function

Public Function Radiant() As Single
 Radiant = mRadiant
End Function

Public Function HugeSingle() As Single
 HugeSingle = mHugeSingle
End Function

Public Function Radian_to_Degree(Radians As Single) As Single
 Radian_to_Degree = ((Radians * 180) / mPI)
End Function

Public Function Degrees_To_Radian(Degrees As Single) As Single
 Degrees_To_Radian = ((Degrees * mPI) / 180)
End Function

'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|
'|þþ                    Matrix Functions                     þþ|
'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|

Public Function Matrix_Copy(Matrix As D3DMATRIX) As D3DMATRIX
 Matrix_Copy.m11 = Matrix.m11
 Matrix_Copy.m12 = Matrix.m12
 Matrix_Copy.m13 = Matrix.m13
 Matrix_Copy.m14 = Matrix.m14
 Matrix_Copy.m21 = Matrix.m21
 Matrix_Copy.m22 = Matrix.m22
 Matrix_Copy.m23 = Matrix.m23
 Matrix_Copy.m24 = Matrix.m24
 Matrix_Copy.m31 = Matrix.m31
 Matrix_Copy.m32 = Matrix.m32
 Matrix_Copy.m33 = Matrix.m33
 Matrix_Copy.m34 = Matrix.m34
 Matrix_Copy.m41 = Matrix.m41
 Matrix_Copy.m42 = Matrix.m42
 Matrix_Copy.m43 = Matrix.m43
 Matrix_Copy.m44 = Matrix.m44
End Function

Public Function Matrix_Identity(Matrix As D3DMATRIX)
 Matrix.m11 = 1: Matrix.m12 = 0: Matrix.m13 = 0: Matrix.m14 = 0
 Matrix.m21 = 0: Matrix.m22 = 1: Matrix.m23 = 0: Matrix.m24 = 0
 Matrix.m31 = 0: Matrix.m32 = 0: Matrix.m33 = 1: Matrix.m34 = 0
 Matrix.m41 = 0: Matrix.m42 = 0: Matrix.m43 = 0: Matrix.m44 = 1
End Function

Public Function Matrix_Translate(SrcMatrix As D3DMATRIX, PositionX As Single, PositionY As Single, PositionZ As Single) As D3DMATRIX
 Dim TempMatrix As D3DMATRIX
 D3DXMatrixTranslation TempMatrix, PositionX, PositionY, PositionZ
 D3DXMatrixMultiply Matrix_Translate, SrcMatrix, TempMatrix
End Function

Public Function Matrix_Scale(SrcMatrix As D3DMATRIX, ScaleX As Single, ScaleY As Single, ScaleZ As Single) As D3DMATRIX
 Dim TempMatrix As D3DMATRIX
 D3DXMatrixScaling TempMatrix, ScaleX, ScaleY, ScaleZ
 D3DXMatrixMultiply Matrix_Scale, SrcMatrix, TempMatrix
End Function

Public Function Matrix_Affine_Transformation(Scaling As Single, Rotation_Center As D3DVECTOR, _
                                             Rotation As D3DQUATERNION, Translation As D3DVECTOR) As D3DMATRIX
 D3DXMatrixAffineTransformation Matrix_Affine_Transformation, Scaling, Rotation_Center, Rotation, Translation
End Function

Public Function Matrix_Ortho_LH(Width As Single, Height As Single, Near As Single, Far As Single) As D3DMATRIX
 D3DXMatrixOrthoLH Matrix_Ortho_LH, Width, Height, Near, Far
End Function

Public Function Matrix_Ortho_RH(Width As Single, Height As Single, Near As Single, Far As Single) As D3DMATRIX
 D3DXMatrixOrthoRH Matrix_Ortho_RH, Width, Height, Near, Far
End Function

Public Function Matrix_Ortho_OffCenter_LH(XMin As Single, XMax As Single, YMin As Single, YMax As Single, Near As Single, Far As Single) As D3DMATRIX
 D3DXMatrixOrthoOffCenterLH Matrix_Ortho_OffCenter_LH, XMin, XMax, YMax, YMin, Near, Far
End Function

Public Function Matrix_Ortho_OffCenter_RH(XMin As Single, XMax As Single, YMin As Single, YMax As Single, Near As Single, Far As Single) As D3DMATRIX
 D3DXMatrixOrthoOffCenterRH Matrix_Ortho_OffCenter_RH, XMin, XMax, YMax, YMin, Near, Far
End Function

Public Function Matrix_Perspective_LH(Width As Single, Height As Single, Near As Single, Far As Single) As D3DMATRIX
 D3DXMatrixPerspectiveLH Matrix_Perspective_LH, Width, Height, Near, Far
End Function

Public Function Matrix_Perspective_RH(Width As Single, Height As Single, Near As Single, Far As Single) As D3DMATRIX
 D3DXMatrixPerspectiveRH Matrix_Perspective_RH, Width, Height, Near, Far
End Function

Public Function Matrix_Perspective_FOV_LH(Feild_Of_View As Single, Aspect As Single, Near As Single, Far As Single) As D3DMATRIX
 D3DXMatrixPerspectiveFovLH Matrix_Perspective_FOV_LH, Feild_Of_View, Aspect, Near, Far
End Function

Public Function Matrix_Perspective_FOV_RH(Feild_Of_View As Single, Aspect As Single, Near As Single, Far As Single) As D3DMATRIX
 D3DXMatrixPerspectiveFovRH Matrix_Perspective_FOV_RH, Feild_Of_View, Aspect, Near, Far
End Function

Public Function Matrix_Perspective_OffCenter_LH(XMin As Single, XMax As Single, YMin As Single, YMax As Single, Near As Single, Far As Single) As D3DMATRIX
 D3DXMatrixPerspectiveOffCenterLH Matrix_Perspective_OffCenter_LH, XMin, XMax, YMax, YMin, Near, Far
End Function

Public Function Matrix_Perspective_OffCenter_RH(XMin As Single, XMax As Single, YMin As Single, YMax As Single, Near As Single, Far As Single) As D3DMATRIX
 D3DXMatrixPerspectiveOffCenterRH Matrix_Perspective_OffCenter_RH, XMin, XMax, YMax, YMin, Near, Far
End Function

Public Function Matrix_LookAt_LH(Eye As D3DVECTOR, At As D3DVECTOR, Up As D3DVECTOR) As D3DMATRIX
 D3DXMatrixLookAtLH Matrix_LookAt_LH, Eye, At, Up
End Function

Public Function Matrix_LookAt_RH(Eye As D3DVECTOR, At As D3DVECTOR, Up As D3DVECTOR) As D3DMATRIX
 D3DXMatrixLookAtRH Matrix_LookAt_RH, Eye, At, Up
End Function

Public Function Matrix_Transformation(Scale_Center As D3DVECTOR, Scale_Rotation As D3DQUATERNION, _
                                      VScale As D3DVECTOR, Rotation_Center As D3DVECTOR, QRotation As D3DQUATERNION, VTranslation As D3DVECTOR) As D3DMATRIX
 D3DXMatrixTransformation Matrix_Transformation, Scale_Center, Scale_Rotation, VScale, Rotation_Center, QRotation, VTranslation
End Function

Public Function Matrix_Transform_Coord(Matrix As D3DMATRIX, Vector As D3DVECTOR) As D3DVECTOR
 D3DXVec3TransformCoord Matrix_Transform_Coord, Vector, Matrix
End Function

Public Function Matrix_Transpose(Destination_Matrix As D3DMATRIX, Matrix As D3DMATRIX) As Long
 Matrix_Transpose = D3DXMatrixTranspose(Destination_Matrix, Matrix)
End Function

Public Function Matrix_Rotation_X(Destination_Matrix As D3DMATRIX, Angle As Single) As Long
 Matrix_Rotation_X = D3DXMatrixRotationX(Destination_Matrix, Angle)
End Function

Public Function Matrix_Rotation_Y(Destination_Matrix As D3DMATRIX, Angle As Single) As Long
 Matrix_Rotation_Y = D3DXMatrixRotationY(Destination_Matrix, Angle)
End Function

Public Function Matrix_Rotation_Z(Destination_Matrix As D3DMATRIX, Angle As Single) As Long
 Matrix_Rotation_Z = D3DXMatrixRotationZ(Destination_Matrix, Angle)
End Function

Public Function Matrix_RotationEX(SrcMatrix As D3DMATRIX, X As Single, Y As Single, Z As Single) As D3DMATRIX
 Dim TempMatrix As D3DMATRIX
 D3DXMatrixRotationYawPitchRoll TempMatrix, Y * mPI / 180, X * mPI / 180, Z * mPI / 180
 D3DXMatrixMultiply Matrix_RotationEX, SrcMatrix, TempMatrix
End Function

Public Function Matrix_Rotation_Quaternion(Destination_Matrix As D3DMATRIX, Quaternion As D3DQUATERNION) As Long
 Matrix_Rotation_Quaternion = D3DXMatrixRotationQuaternion(Destination_Matrix, Quaternion)
End Function

Public Function Matrix_Rotation_Axis(Destination_Matrix As D3DMATRIX, Axis As D3DVECTOR, Angle As Single) As Long
 Matrix_Rotation_Axis = D3DXMatrixRotationAxis(Destination_Matrix, Axis, Angle)
End Function

Public Function Matrix_Yaw_Pitch_Roll(Destination_Matrix As D3DMATRIX, Yaw As Single, Pitch As Single, Roll As Single) As Long
 Matrix_Identity Destination_Matrix
 Matrix_Yaw_Pitch_Roll = D3DXMatrixRotationYawPitchRoll(Destination_Matrix, Yaw, Pitch, Roll)
End Function

Public Function Matrix_Determinant(Matrix As D3DMATRIX) As Single
 Matrix_Determinant = D3DXMatrixfDeterminant(Matrix)
End Function

Public Function Matrix_Inverse(Destination_Matrix As D3DMATRIX, Determinant As Single, Matrix As D3DMATRIX) As Long
 Matrix_Inverse = D3DXMatrixInverse(Destination_Matrix, Determinant, Matrix)
End Function

Public Sub Matrix_Zero(Matrix As D3DMATRIX)
 Matrix.m11 = 0: Matrix.m12 = 0: Matrix.m13 = 0: Matrix.m14 = 0
 Matrix.m21 = 0: Matrix.m22 = 0: Matrix.m23 = 0: Matrix.m24 = 0
 Matrix.m31 = 0: Matrix.m32 = 0: Matrix.m33 = 0: Matrix.m34 = 0
 Matrix.m41 = 0: Matrix.m42 = 0: Matrix.m43 = 0: Matrix.m44 = 0
End Sub

Public Function Matrix_ConcatTransform(Matrix_Input1 As D3DMATRIX, Matrix_Input2 As D3DMATRIX) As D3DMATRIX
 With Matrix_ConcatTransform
  .m11 = Matrix_Input1.m11 * Matrix_Input2.m11 + Matrix_Input1.m21 * Matrix_Input2.m12 + Matrix_Input1.m31 * Matrix_Input2.m13
  .m12 = Matrix_Input1.m12 * Matrix_Input2.m11 + Matrix_Input1.m22 * Matrix_Input2.m12 + Matrix_Input1.m32 * Matrix_Input2.m13
  .m13 = Matrix_Input1.m13 * Matrix_Input2.m11 + Matrix_Input1.m23 * Matrix_Input2.m12 + Matrix_Input1.m33 * Matrix_Input2.m13
  .m21 = Matrix_Input1.m11 * Matrix_Input2.m21 + Matrix_Input1.m21 * Matrix_Input2.m22 + Matrix_Input1.m31 * Matrix_Input2.m23
  .m22 = Matrix_Input1.m12 * Matrix_Input2.m21 + Matrix_Input1.m22 * Matrix_Input2.m22 + Matrix_Input1.m32 * Matrix_Input2.m23
  .m23 = Matrix_Input1.m13 * Matrix_Input2.m21 + Matrix_Input1.m23 * Matrix_Input2.m22 + Matrix_Input1.m33 * Matrix_Input2.m23
  .m31 = Matrix_Input1.m11 * Matrix_Input2.m31 + Matrix_Input1.m21 * Matrix_Input2.m32 + Matrix_Input1.m31 * Matrix_Input2.m33
  .m32 = Matrix_Input1.m12 * Matrix_Input2.m31 + Matrix_Input1.m22 * Matrix_Input2.m32 + Matrix_Input1.m32 * Matrix_Input2.m33
  .m33 = Matrix_Input1.m13 * Matrix_Input2.m31 + Matrix_Input1.m23 * Matrix_Input2.m32 + Matrix_Input1.m33 * Matrix_Input2.m33
  .m41 = Matrix_Input1.m11 * Matrix_Input2.m41 + Matrix_Input1.m21 * Matrix_Input2.m42 + Matrix_Input1.m31 * Matrix_Input2.m43 + Matrix_Input1.m41
  .m42 = Matrix_Input1.m12 * Matrix_Input2.m41 + Matrix_Input1.m22 * Matrix_Input2.m42 + Matrix_Input1.m32 * Matrix_Input2.m43 + Matrix_Input1.m42
  .m43 = Matrix_Input1.m13 * Matrix_Input2.m41 + Matrix_Input1.m23 * Matrix_Input2.m42 + Matrix_Input1.m33 * Matrix_Input2.m42 + Matrix_Input1.m43
 End With
End Function

Public Sub Matrix_Quaternion(Quaternion As D3DQUATERNION, Matrix As D3DMATRIX)
 With Matrix
  .m11 = 1 - 2 * Quaternion.Y * Quaternion.Y - 2 * Quaternion.Z * Quaternion.Z
  .m12 = 2 * Quaternion.X * Quaternion.Y + 2 * Quaternion.W * Quaternion.Z
  .m13 = 2 * Quaternion.X * Quaternion.Z - 2 * Quaternion.W * Quaternion.Y
  .m21 = 2 * Quaternion.X * Quaternion.Y - 2 * Quaternion.W * Quaternion.Z
  .m22 = 1# - 2# * Quaternion.X * Quaternion.X - 2# * Quaternion.Z * Quaternion.Z
  .m23 = 2# * Quaternion.Y * Quaternion.Z + 2# * Quaternion.W * Quaternion.X
  .m31 = 2# * Quaternion.X * Quaternion.Z + 2# * Quaternion.W * Quaternion.Y
  .m32 = 2# * Quaternion.Y * Quaternion.Z - 2# * Quaternion.W * Quaternion.X
  .m33 = 1# - 2# * Quaternion.X * Quaternion.X - 2# * Quaternion.Y * Quaternion.Y
 End With
End Sub

Public Function Matrix_Is_Identity(Matrix As D3DMATRIX) As Long
 Matrix_Is_Identity = D3DXMatrixIsIdentity(Matrix)
End Function

Public Function Matrix_Multiply(Destination_Matrix As D3DMATRIX, Matrix1 As D3DMATRIX, Matrix2 As D3DMATRIX) As Long
 Matrix_Multiply = D3DXMatrixMultiply(Destination_Matrix, Matrix1, Matrix2)
End Function

Public Function Matrix_Reflect(Destination_Matrix As D3DMATRIX, Plane As D3DPLANE) As Long
 Matrix_Reflect = D3DXMatrixReflect(Destination_Matrix, Plane)
End Function

Public Function Matrix_Shadow(Destination_Matrix As D3DMATRIX, Vector_Light As D3DVECTOR4, Plane As D3DPLANE) As Long
 Matrix_Shadow = D3DXMatrixShadow(Destination_Matrix, Vector_Light, Plane)
End Function

Public Function Matrix_Read_From_File(MatrixOut As D3DMATRIX, Input_Filenme As String)
 Set InputStream = FSys.OpenTextFile(Input_Filenme & ".txt")

 MatrixOut.m11 = InputStream.ReadLine
 MatrixOut.m12 = InputStream.ReadLine
 MatrixOut.m13 = InputStream.ReadLine
 MatrixOut.m14 = InputStream.ReadLine
 MatrixOut.m21 = InputStream.ReadLine
 MatrixOut.m22 = InputStream.ReadLine
 MatrixOut.m23 = InputStream.ReadLine
 MatrixOut.m24 = InputStream.ReadLine
 MatrixOut.m31 = InputStream.ReadLine
 MatrixOut.m32 = InputStream.ReadLine
 MatrixOut.m33 = InputStream.ReadLine
 MatrixOut.m34 = InputStream.ReadLine
 MatrixOut.m41 = InputStream.ReadLine
 MatrixOut.m42 = InputStream.ReadLine
 MatrixOut.m43 = InputStream.ReadLine
 MatrixOut.m44 = InputStream.ReadLine

 Set InputStream = Nothing
End Function

Public Function Matrix_Print_To_File(Matrix As D3DMATRIX, Output_Filenme As String)
 On Local Error Resume Next
 Dim Results As Long
 Results = 0
 Set OutStream = FSys.CreateTextFile(Output_Filenme & ".txt", True, False)

 OutStream.WriteLine "Matrix - " & CStr(Output_Filenme)
 OutStream.WriteLine ""
 OutStream.WriteLine "Matrix.m11 - " & Matrix.m11 & " " & "Matrix.m12 - " & Matrix.m12 & " " & "Matrix.m13 - " & Matrix.m13 & " " & "Matrix.m14 - " & Matrix.m14
 OutStream.WriteLine ""
 OutStream.WriteLine "Matrix.m21 - " & Matrix.m21 & " " & "Matrix.m22 - " & Matrix.m22 & " " & "Matrix.m23 - " & Matrix.m23 & " " & "Matrix.m24 - " & Matrix.m24
 OutStream.WriteLine ""
 OutStream.WriteLine "Matrix.m31 - " & Matrix.m31 & " " & "Matrix.m32 - " & Matrix.m32 & " " & "Matrix.m33 - " & Matrix.m33 & " " & "Matrix.m34 - " & Matrix.m34
 OutStream.WriteLine ""
 OutStream.WriteLine "Matrix.m41 - " & Matrix.m41 & " " & "Matrix.m42 - " & Matrix.m42 & " " & "Matrix.m43 - " & Matrix.m43 & " " & "Matrix.m44 - " & Matrix.m44
 OutStream.WriteLine ""
 OutStream.WriteLine ""

 If Matrix.m11 = 1 And Matrix.m22 = 1 And Matrix.m33 = 1 And Matrix.m44 = 1 Then
  OutStream.WriteLine CStr(Output_Filenme) & " has been identified as an Identity Matrix"
  OutStream.WriteLine ""
  Results = Results + 1
 End If
 If Matrix.m11 = 0 And Matrix.m12 = 0 And Matrix.m13 = 0 And Matrix.m14 = 0 And _
        Matrix.m21 = 0 And Matrix.m22 = 0 And Matrix.m23 = 0 And Matrix.m24 = 0 And _
        Matrix.m31 = 0 And Matrix.m32 = 0 And Matrix.m33 = 0 And Matrix.m34 = 0 And _
        Matrix.m41 = 0 And Matrix.m42 = 0 And Matrix.m43 = 0 And Matrix.m44 = 0 Then
  If Results <> 0 Then
   OutStream.WriteLine CStr(Output_Filenme) & " has also been identified as an Zero Matrix"
   OutStream.WriteLine ""
  Else
   OutStream.WriteLine CStr(Output_Filenme) & " has been identified as an Zero Matrix"
   OutStream.WriteLine ""
  End If
  Results = Results + 1
 End If
 If Matrix.m41 <> 0 And Matrix.m42 <> 0 And Matrix.m43 <> 0 Then
  If Results <> 0 Then
   OutStream.WriteLine CStr(Output_Filenme) & " has also been identified as an Translation Matrix"
   OutStream.WriteLine ""
  Else
   OutStream.WriteLine CStr(Output_Filenme) & " has been identified as an Translation Matrix"
   OutStream.WriteLine ""
  End If
  Results = Results + 1
 End If
 If Matrix.m11 <> 1 And Matrix.m22 <> 1 And Matrix.m33 <> 1 Then
  If Results <> 0 Then
   OutStream.WriteLine CStr(Output_Filenme) & " has also been identified as an Scale Matrix"
   OutStream.WriteLine ""
  Else
   OutStream.WriteLine CStr(Output_Filenme) & " has been identified as an Scale Matrix"
   OutStream.WriteLine ""
  End If
  Results = Results + 1
 End If

 Set OutStream = Nothing
End Function

'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|
'|þþ                    Vector 2 Functions                   þþ|
'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|

Public Function Vector2(X As Single, Y As Single) As D3DVECTOR2
 Vector2.X = X
 Vector2.Y = Y
End Function

Public Function Vector2_Add(Vector1 As D3DVECTOR2, Vector2 As D3DVECTOR2) As D3DVECTOR2
 Vector2_Add.X = Vector1.X + Vector2.X
 Vector2_Add.Y = Vector1.Y + Vector2.Y
End Function

Public Function Vector2_Nearly_Equal(Vector1 As D3DVECTOR2, Vector2 As D3DVECTOR2, Radius As Single) As Boolean
 Dim Diff As D3DVECTOR2
 Diff.X = Vector1.X - Vector2.X
 Diff.Y = Vector1.Y - Vector2.Y
 Vector2_Nearly_Equal = (Vector2_Dot_Product(Vector1, Vector2) < Radius * Radius)
End Function

Public Function Vector2_Swap(Vector1 As D3DVECTOR2, Vector2 As D3DVECTOR2)
 Dim TmpVec As D3DVECTOR2
 TmpVec = Vector1
 Vector1 = Vector2
 Vector2 = TmpVec
End Function

Public Function Vector2_Multiply(Vector1 As D3DVECTOR2, Vector2 As D3DVECTOR2) As D3DVECTOR2
 Vector2_Multiply.X = Vector1.X * Vector2.X
 Vector2_Multiply.Y = Vector1.Y * Vector2.Y
End Function

Public Function Vector2_Divide(Vector1 As D3DVECTOR2, Vector2 As D3DVECTOR2) As D3DVECTOR2
 Vector2_Divide.X = Vector1.X / Vector2.X
 Vector2_Divide.Y = Vector1.Y / Vector2.Y
End Function

Public Function Vector2_Average(Vector1 As D3DVECTOR2, Vector2 As D3DVECTOR2) As D3DVECTOR2
 Vector2_Average = Vector2_Add(Vector1, Vector2)
 Vector2_Average.X = Vector2_Average.X / 2
 Vector2_Average.Y = Vector2_Average.Y / 2
End Function

Public Function Vector2_Subtract(Vector1 As D3DVECTOR2, Vector2 As D3DVECTOR2) As D3DVECTOR2
 Vector2_Subtract.X = Vector1.X - Vector2.X
 Vector2_Subtract.Y = Vector1.Y - Vector2.Y
End Function

Public Sub Vector2_Negative(Vector As D3DVECTOR2)
 Vector.X = -Vector.X
 Vector.Y = -Vector.Y
End Sub

Public Sub Vector2_Normalize(Vector As D3DVECTOR2)
 On Local Error Resume Next
 Dim Value As Double
 Value = Vector.X * Vector.X + Vector.Y * Vector.Y
 Value = Sqr(Value)

 If Value = 0 Then
  Vector.X = 0
  Vector.Y = 0
  Exit Sub
 End If

 Vector.X = Vector.X / Value
 Vector.Y = Vector.Y / Value
End Sub
  
Public Function Vector2_Bary_Centric(Destination_Vector As D3DVECTOR2, Vector1 As D3DVECTOR2, Vector2 As D3DVECTOR2, Vector3 As D3DVECTOR2, Weight_Factor_F As Single, Weight_Factor_G As Single) As Long
 Vector2_Bary_Centric = D3DXVec2BaryCentric(Destination_Vector, Vector1, Vector2, Vector3, Weight_Factor_F, Weight_Factor_G)
End Function

Public Function Vector2_Catmull_Rom(Destination_Vector As D3DVECTOR2, Vector1 As D3DVECTOR2, Vector2 As D3DVECTOR2, Vector3 As D3DVECTOR2, Vector4 As D3DVECTOR2, Weight_Factor_S As Single) As Long
 Vector2_Catmull_Rom = D3DXVec2CatmullRom(Destination_Vector, Vector1, Vector2, Vector3, Vector4, Weight_Factor_S)
End Function

Public Function Vector2_Lerp(Destination_Vector As D3DVECTOR2, Vector1 As D3DVECTOR2, Vector2 As D3DVECTOR2, S As Single) As Long
 Vector2_Lerp = D3DXVec2Lerp(Destination_Vector, Vector1, Vector2, S)
End Function

Public Function Vector2_Maximize(Destination_Vector As D3DVECTOR2, Vector1 As D3DVECTOR2, Vector2 As D3DVECTOR2) As Long
 Vector2_Maximize = D3DXVec2Maximize(Destination_Vector, Vector1, Vector2)
End Function

Public Function Vector2_Transform_Coord(Destination_Vector As D3DVECTOR2, Vector1 As D3DVECTOR2, Matrix As D3DMATRIX) As Long
 Vector2_Transform_Coord = D3DXVec2TransformCoord(Destination_Vector, Vector1, Matrix)
End Function

Public Function Vector2_Transform_Normal(Destination_Vector As D3DVECTOR2, Vector1 As D3DVECTOR2, Matrix As D3DMATRIX) As Long
 Vector2_Transform_Normal = D3DXVec2TransformNormal(Destination_Vector, Vector1, Matrix)
End Function

Public Function Vector2_Minimize(Destination_Vector As D3DVECTOR2, Vector1 As D3DVECTOR2, Vector2 As D3DVECTOR2) As Long
 Vector2_Minimize = D3DXVec2Minimize(Destination_Vector, Vector1, Vector2)
End Function

Public Function Vector2_CCW(Vector1 As D3DVECTOR2, Vector2 As D3DVECTOR2) As Single
 Vector2_CCW = D3DXVec2CCW(Vector1, Vector2)
End Function

Public Function Vector2_Hermit(Destination_Vector As D3DVECTOR2, Vector1 As D3DVECTOR2, Tangent_Vector1 As D3DVECTOR2, Vector2 As D3DVECTOR2, Tangent_Vector2 As D3DVECTOR2, Weight_Factor_S As Single) As Long
 Vector2_Hermit = D3DXVec2Hermite(Destination_Vector, Vector1, Tangent_Vector1, Vector2, Tangent_Vector2, Weight_Factor_S)
End Function

Public Function Vector2_Dot_Product(Vector1 As D3DVECTOR2, Vector2 As D3DVECTOR2) As Single
 Vector2_Dot_Product = Vector1.X * Vector2.X + Vector1.Y * Vector2.Y
End Function

Public Sub Vector2_Add_And_Scale(Destination_Vector As D3DVECTOR2, Vector1 As D3DVECTOR2, Scale_Amount1 As Single, Vector2 As D3DVECTOR2, Scale_Amount2 As Single)
 Destination_Vector.X = Scale_Amount1 * Vector1.X + Scale_Amount2 * Vector2.X
 Destination_Vector.Y = Scale_Amount1 * Vector1.Y + Scale_Amount2 * Vector2.Y
End Sub

Public Sub Vector2_Subtract_And_Scale(Destination_Vector As D3DVECTOR2, Vector1 As D3DVECTOR2, Scale_Amount1 As Single, Vector2 As D3DVECTOR2, Scale_Amount2 As Single)
 Destination_Vector.X = Scale_Amount1 * Vector1.X - Scale_Amount2 * Vector2.X
 Destination_Vector.Y = Scale_Amount1 * Vector1.Y - Scale_Amount2 * Vector2.Y
End Sub

Public Sub Vector2_Copy(Destination_Vector As D3DVECTOR2, Source_Vector As D3DVECTOR2)
 Destination_Vector.X = Source_Vector.X
 Destination_Vector.Y = Source_Vector.Y
End Sub

Public Sub Vector2_Scale(Destination_Vector As D3DVECTOR2, Source_Vector As D3DVECTOR2, Scale_Amount As Single)
 Destination_Vector.X = Source_Vector.X * Scale_Amount
 Destination_Vector.Y = Source_Vector.Y * Scale_Amount
End Sub

Public Function Vector2_Length(Vector As D3DVECTOR2) As Single
 Vector2_Length = D3DXVec2Length(Vector)
End Function

Public Function Vector2_Length_Squared(Vector As D3DVECTOR2) As Single
 Vector2_Length_Squared = D3DXVec2LengthSq(Vector)
End Function

'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|
'|þþ                    Vector 3 Functions                   þþ|
'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|

Public Function Vector3(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
 Vector3.X = X
 Vector3.Y = Y
 Vector3.Z = Z
End Function

Public Function Vector3_Add(Vector1 As D3DVECTOR, Vector2 As D3DVECTOR) As D3DVECTOR
 Vector3_Add.X = Vector1.X + Vector2.X
 Vector3_Add.Y = Vector1.Y + Vector2.Y
 Vector3_Add.Z = Vector1.Z + Vector2.Z
End Function

Public Function Vector3_Nearly_Equal(Vector1 As D3DVECTOR, Vector2 As D3DVECTOR, Radius As Single) As Boolean
 Dim Diff As D3DVECTOR
 Diff.X = Vector1.X - Vector2.X
 Diff.Y = Vector1.Y - Vector2.Y
 Diff.Z = Vector1.Z - Vector2.Z
 Vector3_Nearly_Equal = (Vector3_Dot_Product(Vector1, Vector2) < Radius * Radius)
End Function

Public Function Vector3_Multiply(Vector1 As D3DVECTOR, Vector2 As D3DVECTOR) As D3DVECTOR
 Vector3_Multiply.X = Vector1.X * Vector2.X
 Vector3_Multiply.Y = Vector1.Y * Vector2.Y
 Vector3_Multiply.Z = Vector1.Z * Vector2.Z
End Function

Public Function Vector3_Divide(Vector1 As D3DVECTOR, Vector2 As D3DVECTOR) As D3DVECTOR
 Vector3_Divide.X = Vector1.X / Vector2.X
 Vector3_Divide.Y = Vector1.Y / Vector2.Y
 Vector3_Divide.Z = Vector1.Z / Vector2.Z
End Function

Public Function Vector3_Swap(Vector1 As D3DVECTOR, Vector2 As D3DVECTOR)
 Dim TmpVec As D3DVECTOR
 TmpVec = Vector1
 Vector1 = Vector2
 Vector2 = TmpVec
End Function

Public Function Vector3_Bary_Centric(Destination_Vector As D3DVECTOR, Vector1 As D3DVECTOR, Vector2 As D3DVECTOR, Vector3 As D3DVECTOR, Weight_Factor_F As Single, Weight_Factor_G As Single) As Long
 Vector3_Bary_Centric = D3DXVec3BaryCentric(Destination_Vector, Vector1, Vector2, Vector3, Weight_Factor_F, Weight_Factor_G)
End Function

Public Function Vector3_Catmull_Rom(Destination_Vector As D3DVECTOR, Vector1 As D3DVECTOR, Vector2 As D3DVECTOR, Vector3 As D3DVECTOR, Vector4 As D3DVECTOR, Weight_Factor_S As Single) As Long
 Vector3_Catmull_Rom = D3DXVec3CatmullRom(Destination_Vector, Vector1, Vector2, Vector3, Vector4, Weight_Factor_S)
End Function

Public Function Vector3_Lerp(Destination_Vector As D3DVECTOR, Vector1 As D3DVECTOR, Vector2 As D3DVECTOR, S As Single) As Long
 Vector3_Lerp = D3DXVec3Lerp(Destination_Vector, Vector1, Vector2, S)
End Function

Public Function Vector3_Maximize(Destination_Vector As D3DVECTOR, Vector1 As D3DVECTOR, Vector2 As D3DVECTOR) As Long
 Vector3_Maximize = D3DXVec3Maximize(Destination_Vector, Vector1, Vector2)
End Function

Public Function Vector3_Transform_Coord(Destination_Vector As D3DVECTOR, Vector1 As D3DVECTOR, Matrix As D3DMATRIX) As Long
 Vector3_Transform_Coord = D3DXVec3TransformCoord(Destination_Vector, Vector1, Matrix)
End Function

Public Function Vector3_Transform(Destination_Vector As D3DVECTOR4, Vector1 As D3DVECTOR, Matrix As D3DMATRIX) As Long
 Vector3_Transform = D3DXVec3Transform(Destination_Vector, Vector1, Matrix)
End Function

Public Function Vector3_Transform_Normal(Destination_Vector As D3DVECTOR, Vector1 As D3DVECTOR, Matrix As D3DMATRIX) As Long
 Vector3_Transform_Normal = D3DXVec3TransformNormal(Destination_Vector, Vector1, Matrix)
End Function

Public Function Vector3_Minimize(Destination_Vector As D3DVECTOR, Vector1 As D3DVECTOR, Vector2 As D3DVECTOR) As Long
 Vector3_Minimize = D3DXVec3Minimize(Destination_Vector, Vector1, Vector2)
End Function

Public Function Vector3_Hermit(Destination_Vector As D3DVECTOR, Vector1 As D3DVECTOR, Tangent_Vector1 As D3DVECTOR, Vector2 As D3DVECTOR, Tangent_Vector2 As D3DVECTOR, Weight_Factor_S As Single) As Long
 Vector3_Hermit = D3DXVec3Hermite(Destination_Vector, Vector1, Tangent_Vector1, Vector2, Tangent_Vector2, Weight_Factor_S)
End Function

Public Function Vector3_Average(Vector1 As D3DVECTOR, Vector2 As D3DVECTOR) As D3DVECTOR
 Vector3_Average = Vector3_Add(Vector1, Vector2)
 Vector3_Average.X = Vector3_Average.X / 2
 Vector3_Average.Y = Vector3_Average.Y / 2
 Vector3_Average.Z = Vector3_Average.Z / 2
End Function

Public Function Vector3_Subtract(Vector1 As D3DVECTOR, Vector2 As D3DVECTOR) As D3DVECTOR
 Vector3_Subtract.X = Vector1.X - Vector2.X
 Vector3_Subtract.Y = Vector1.Y - Vector2.Y
 Vector3_Subtract.Z = Vector1.Z - Vector2.Z
End Function

Public Sub Vector3_Negative(Vector As D3DVECTOR)
 Vector.X = -Vector.X
 Vector.Y = -Vector.Y
 Vector.Z = -Vector.Z
End Sub

Public Function Vector3_Cross_Product(Vector1 As D3DVECTOR, Vector2 As D3DVECTOR) As D3DVECTOR
 Dim vTemp As D3DVECTOR
 Dim vTemp2 As D3DVECTOR
 vTemp.X = Vector1.Y * Vector2.Z - Vector1.Z * Vector2.Y
 vTemp.Y = Vector1.Z * Vector2.X - Vector1.X * Vector2.Z
 vTemp.Z = Vector1.X * Vector2.Y - Vector1.Y * Vector2.X
 vTemp2 = Vector3_Normalize(vTemp)
 Vector3_Cross_Product = vTemp2
End Function
  
Public Function Vector3_Normalize(Vector As D3DVECTOR) As D3DVECTOR
 On Local Error Resume Next
 Dim Value As Double
 Value = Vector.X * Vector.X + Vector.Y * Vector.Y + Vector.Z * Vector.Z
 Value = Sqr(Value)

 If Value = 0 Then Vector3_Normalize = Vector3(0, 0, 0): Exit Function
 Vector3_Normalize = Vector3(Vector.X / Value, Vector.Y / Value, Vector.Z / Value)
End Function
  
Public Function Vector3_Dot_Product(Vector1 As D3DVECTOR, Vector2 As D3DVECTOR) As Single
 Vector3_Dot_Product = Vector1.X * Vector2.X + Vector1.Y * Vector2.Y + Vector1.Z * Vector2.Z
End Function

Public Sub Vector3_Add_And_Scale(Destination_Vector As D3DVECTOR, Vector1 As D3DVECTOR, Scale_Amount1 As Single, Vector2 As D3DVECTOR, Scale_Amount2 As Single)
 Destination_Vector.X = Scale_Amount1 * Vector1.X + Scale_Amount2 * Vector2.X
 Destination_Vector.Y = Scale_Amount1 * Vector1.Y + Scale_Amount2 * Vector2.Y
 Destination_Vector.Z = Scale_Amount1 * Vector1.Z + Scale_Amount2 * Vector2.Z
End Sub

Public Sub Vector3_Subtract_And_Scale(Destination_Vector As D3DVECTOR, Vector1 As D3DVECTOR, Scale_Amount1 As Single, Vector2 As D3DVECTOR, Scale_Amount2 As Single)
 Destination_Vector.X = Scale_Amount1 * Vector1.X - Scale_Amount2 * Vector2.X
 Destination_Vector.Y = Scale_Amount1 * Vector1.Y - Scale_Amount2 * Vector2.Y
 Destination_Vector.Z = Scale_Amount1 * Vector1.Z - Scale_Amount2 * Vector2.Z
End Sub

Public Sub Vector3_Copy(Destination_Vector As D3DVECTOR, Source_Vector As D3DVECTOR)
 Destination_Vector.X = Source_Vector.X
 Destination_Vector.Y = Source_Vector.Y
 Destination_Vector.Z = Source_Vector.Z
End Sub

Public Function Vector3_Scale(Vector As D3DVECTOR, Scale_Amount As Single) As D3DVECTOR
 Vector3_Scale.X = Vector.X * Scale_Amount
 Vector3_Scale.Y = Vector.Y * Scale_Amount
 Vector3_Scale.Z = Vector.Z * Scale_Amount
End Function

Public Function Vector3_Length(Vector As D3DVECTOR) As Single
 Vector3_Length = D3DXVec3Length(Vector)
End Function

Public Function Vector3_Length_Squared(Vector As D3DVECTOR) As Single
 Vector3_Length_Squared = D3DXVec3LengthSq(Vector)
End Function

Public Function Vector3_Project(Destination_Vector As D3DVECTOR, Vector As D3DVECTOR, Viewport As D3DVIEWPORT8, _
                                  Projection As D3DMATRIX, View As D3DMATRIX, World As D3DMATRIX) As Long
 Vector3_Project = D3DXVec3Project(Destination_Vector, Vector, Viewport, Projection, View, World)
End Function

Public Function Vector3_Unproject(Destination_Vector As D3DVECTOR, Vector As D3DVECTOR, Viewport As D3DVIEWPORT8, _
                                  Projection As D3DMATRIX, View As D3DMATRIX, World As D3DMATRIX) As Long
 Vector3_Unproject = D3DXVec3Unproject(Destination_Vector, Vector, Viewport, Projection, View, World)
End Function

Public Function Vector3_Max(V1 As Single, V2 As Single, V3 As Single) As Single
 If V1 > V2 And V1 > V3 Then Vector3_Max = V1
 If V2 > V1 And V2 > V3 Then Vector3_Max = V2
 If V3 > V2 And V3 > V1 Then Vector3_Max = V3
      
 If Vector3_Max = 0 Then
  If V1 > V2 And V1 = V3 Then Vector3_Max = V1
  If V2 > V3 And V2 = V1 Then Vector3_Max = V2
  If V3 > V2 And V3 = V1 Then Vector3_Max = V3
 End If
    
 If Vector3_Max = 0 Then If V1 = V2 And V1 = V3 Then Vector3_Max = V1
End Function

'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|
'|þþ                    Vector 4 Functions                   þþ|
'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|

Public Function Vector4(X As Single, Y As Single, Z As Single, W As Single) As D3DVECTOR4
 Vector4.X = X
 Vector4.Y = Y
 Vector4.Z = Z
 Vector4.W = W
End Function

Public Function Vector4_Nearly_Equal(Vector1 As D3DVECTOR4, Vector2 As D3DVECTOR4, Radius As Single) As Boolean
 Dim Diff As D3DVECTOR4
 Diff.X = Vector1.X - Vector2.X
 Diff.Y = Vector1.Y - Vector2.Y
 Diff.Z = Vector1.Z - Vector2.Z
 Diff.W = Vector1.W - Vector2.W
 Vector4_Nearly_Equal = (Vector4_Dot_Product(Vector1, Vector2) < Radius * Radius)
End Function

Public Function Vector4_Add(Vector1 As D3DVECTOR4, Vector2 As D3DVECTOR4) As D3DVECTOR4
 Vector4_Add.X = Vector1.X + Vector2.X
 Vector4_Add.Y = Vector1.Y + Vector2.Y
 Vector4_Add.Z = Vector1.Z + Vector2.Z
 Vector4_Add.W = Vector1.W + Vector2.W
End Function

Public Function Vector4_Multiply(Vector1 As D3DVECTOR4, Vector2 As D3DVECTOR4) As D3DVECTOR4
 Vector4_Multiply.X = Vector1.X * Vector2.X
 Vector4_Multiply.Y = Vector1.Y * Vector2.Y
 Vector4_Multiply.Z = Vector1.Z * Vector2.Z
 Vector4_Multiply.W = Vector1.W * Vector2.W
End Function

Public Function Vector4_Divide(Vector1 As D3DVECTOR4, Vector2 As D3DVECTOR4) As D3DVECTOR4
 Vector4_Divide.X = Vector1.X / Vector2.X
 Vector4_Divide.Y = Vector1.Y / Vector2.Y
 Vector4_Divide.Z = Vector1.Z / Vector2.Z
 Vector4_Divide.W = Vector1.W / Vector2.W
End Function

Public Function Vector4_Swap(Vector1 As D3DVECTOR4, Vector2 As D3DVECTOR4)
 Dim TmpVec As D3DVECTOR4
 TmpVec = Vector1
 Vector1 = Vector2
 Vector2 = TmpVec
End Function

Public Function Vector4_Average(Vector1 As D3DVECTOR4, Vector2 As D3DVECTOR4) As D3DVECTOR4
 Vector4_Average = Vector4_Add(Vector1, Vector2)
 Vector4_Average.X = Vector4_Average.X / 2
 Vector4_Average.Y = Vector4_Average.Y / 2
 Vector4_Average.Z = Vector4_Average.Z / 2
 Vector4_Average.W = Vector4_Average.W / 2
End Function

Public Function Vector4_Subtract(Vector1 As D3DVECTOR4, Vector2 As D3DVECTOR4) As D3DVECTOR4
 Vector4_Subtract.X = Vector1.X - Vector2.X
 Vector4_Subtract.Y = Vector1.Y - Vector2.Y
 Vector4_Subtract.Z = Vector1.Z - Vector2.Z
 Vector4_Subtract.W = Vector1.W - Vector2.W
End Function

Public Sub Vector4_Negative(Vector As D3DVECTOR4)
 Vector.X = -Vector.X
 Vector.Y = -Vector.Y
 Vector.Z = -Vector.Z
 Vector.W = -Vector.W
End Sub

Public Function Vector4_Bary_Centric(Destination_Vector As D3DVECTOR4, Vector1 As D3DVECTOR4, Vector2 As D3DVECTOR4, Vector3 As D3DVECTOR4, Weight_Factor_F As Single, Weight_Factor_G As Single) As Long
 Vector4_Bary_Centric = D3DXVec4BaryCentric(Destination_Vector, Vector1, Vector2, Vector3, Weight_Factor_F, Weight_Factor_G)
End Function

Public Function Vector4_Catmull_Rom(Destination_Vector As D3DVECTOR4, Vector1 As D3DVECTOR4, Vector2 As D3DVECTOR4, Vector3 As D3DVECTOR4, Vector4 As D3DVECTOR4, Weight_Factor_S As Single) As Long
 Vector4_Catmull_Rom = D3DXVec4CatmullRom(Destination_Vector, Vector1, Vector2, Vector3, Vector4, Weight_Factor_S)
End Function

Public Function Vector4_Lerp(Destination_Vector As D3DVECTOR4, Vector1 As D3DVECTOR4, Vector2 As D3DVECTOR4, S As Single) As Long
 Vector4_Lerp = D3DXVec4Lerp(Destination_Vector, Vector1, Vector2, S)
End Function

Public Function Vector4_Hermit(Destination_Vector As D3DVECTOR4, Vector1 As D3DVECTOR, Tangent_Vector1 As D3DVECTOR4, Vector2 As D3DVECTOR, Tangent_Vector2 As D3DVECTOR4, Weight_Factor_S As Single) As Long
 Vector4_Hermit = D3DXVec4Hermite(Destination_Vector, Vector1, Tangent_Vector1, Vector2, Tangent_Vector2, Weight_Factor_S)
End Function

Public Function Vector4_Maximize(Destination_Vector As D3DVECTOR4, Vector1 As D3DVECTOR4, Vector2 As D3DVECTOR4) As Long
 Vector4_Maximize = D3DXVec4Maximize(Destination_Vector, Vector1, Vector2)
End Function

Public Function Vector4_Minimize(Destination_Vector As D3DVECTOR4, Vector1 As D3DVECTOR4, Vector2 As D3DVECTOR4) As Long
 Vector4_Minimize = D3DXVec4Minimize(Destination_Vector, Vector1, Vector2)
End Function

Public Function Vector4_Transform(Destination_Vector As D3DVECTOR4, Vector1 As D3DVECTOR4, Matrix As D3DMATRIX) As Long
 Vector4_Transform = D3DXVec4Transform(Destination_Vector, Vector1, Matrix)
End Function

Public Function Vector4_Cross_Product(Destination_Vector As D3DVECTOR4, Vector1 As D3DVECTOR4, Vector2 As D3DVECTOR4, Vector3 As D3DVECTOR4) As Long
 Vector4_Cross_Product = D3DXVec4Cross(Destination_Vector, Vector1, Vector2, Vector3)
End Function
  
Public Sub Vector4_Normalize(Vector As D3DVECTOR4)
 On Local Error Resume Next
 Dim Value As Double
 Value = Vector.X * Vector.X + Vector.Y * Vector.Y + Vector.Z * Vector.Z + Vector.W * Vector.W
 Value = Sqr(Value)

 If Value = 0 Then
  Vector.X = 0: Vector.Y = 0
  Vector.Z = 0: Vector.W = 0
  Exit Sub
 End If

 Vector.X = Vector.X / Value
 Vector.Y = Vector.Y / Value
 Vector.Z = Vector.Z / Value
 Vector.W = Vector.W / Value
End Sub
  
Public Function Vector4_Dot_Product(Vector1 As D3DVECTOR4, Vector2 As D3DVECTOR4) As Single
 Vector4_Dot_Product = Vector1.X * Vector2.X + Vector1.Y * Vector2.Y + Vector1.Z * Vector2.Z + Vector1.W * Vector2.W
End Function

Public Sub Vector4_Add_And_Scale(Destination_Vector As D3DVECTOR4, Vector1 As D3DVECTOR4, Scale_Amount1 As Single, Vector2 As D3DVECTOR4, Scale_Amount2 As Single)
 Destination_Vector.X = Scale_Amount1 * Vector1.X + Scale_Amount2 * Vector2.X
 Destination_Vector.Y = Scale_Amount1 * Vector1.Y + Scale_Amount2 * Vector2.Y
 Destination_Vector.Z = Scale_Amount1 * Vector1.Z + Scale_Amount2 * Vector2.Z
 Destination_Vector.W = Scale_Amount1 * Vector1.W + Scale_Amount2 * Vector2.W
End Sub

Public Sub Vector4_Subtract_And_Scale(Destination_Vector As D3DVECTOR4, Vector1 As D3DVECTOR4, Scale_Amount1 As Single, Vector2 As D3DVECTOR4, Scale_Amount2 As Single)
 Destination_Vector.X = Scale_Amount1 * Vector1.X - Scale_Amount2 * Vector2.X
 Destination_Vector.Y = Scale_Amount1 * Vector1.Y - Scale_Amount2 * Vector2.Y
 Destination_Vector.Z = Scale_Amount1 * Vector1.Z - Scale_Amount2 * Vector2.Z
 Destination_Vector.W = Scale_Amount1 * Vector1.W - Scale_Amount2 * Vector2.W
End Sub

Public Sub Vector4_Copy(Destination_Vector As D3DVECTOR4, Source_Vector As D3DVECTOR4)
 Destination_Vector.X = Source_Vector.X
 Destination_Vector.Y = Source_Vector.Y
 Destination_Vector.Z = Source_Vector.Z
 Destination_Vector.W = Source_Vector.W
End Sub

Public Sub Vector4_Scale(Destination_Vector As D3DVECTOR4, Source_Vector As D3DVECTOR4, Scale_Amount As Single)
 Destination_Vector.X = Source_Vector.X * Scale_Amount
 Destination_Vector.Y = Source_Vector.Y * Scale_Amount
 Destination_Vector.Z = Source_Vector.Z * Scale_Amount
 Destination_Vector.W = Source_Vector.W * Scale_Amount
End Sub

Public Function Vector4_Length(Vector As D3DVECTOR4) As Single
 Vector4_Length = D3DXVec4Length(Vector)
End Function

Public Function Vector4_Length_Squared(Vector As D3DVECTOR4) As Single
 Vector4_Length_Squared = D3DXVec4LengthSq(Vector)
End Function

'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|
'|þþ                     Vertex Functions                    þþ|
'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|

Public Function Vertex(V As D3DVERTEX2) As D3DVECTOR
 With Vertex
  .X = V.X: .Y = V.Y: .Z = V.Z
 End With
End Function

Public Function Vertex2(X As Single, Y As Single, Z As Single, _
                        Normal_X As Single, Normal_Y As Single, Normal_Z As Single, _
                        tU As Single, tV As Single) As D3DVERTEX2
 With Vertex2
  .X = X: .Y = Y: .Z = Z
  .nX = Normal_X: .nY = Normal_Y: .nz = Normal_Z
  .tu1 = tU: .tv1 = tV
 End With

End Function

Public Function Vertex3(X As Single, Y As Single, Z As Single, _
                        Normal_X As Single, Normal_Y As Single, Normal_Z As Single, _
                        tU As Single, tV As Single) As D3DVERTEX
 Dim tmpVert As D3DVERTEX
 tmpVert.X = X: tmpVert.Y = Y: tmpVert.Z = Z
 tmpVert.nX = Normal_X: tmpVert.nY = Normal_Y: tmpVert.nz = Normal_Z
 tmpVert.tU = tU: tmpVert.tV = tV
 Vertex3 = tmpVert
End Function

Public Function VertexL2(X As Single, Y As Single, Z As Single, _
                               Color As Long, tU As Single, tV As Single) As D3DLVERTEX2
 Dim tmpVert As D3DLVERTEX2
 tmpVert.X = X: tmpVert.Y = Y: tmpVert.Z = Z
 tmpVert.tu1 = tU: tmpVert.tv1 = tV
 tmpVert.Color = Color: tmpVert.Specular = Color
 VertexL2 = tmpVert
End Function

Public Function VertexL3(X As Single, Y As Single, Z As Single, _
                               Color As Long, tU As Single, tV As Single) As D3DLVERTEX
 Dim tmpVert As D3DLVERTEX
 tmpVert.X = X: tmpVert.Y = Y: tmpVert.Z = Z
 tmpVert.tU = tU: tmpVert.tV = tV
 tmpVert.Color = Color
 tmpVert.Specular = Color
 VertexL3 = tmpVert
End Function

Public Function VertexTL2(Sx As Single, Sy As Single, Sz As Single, _
                                RHW As Single, tu1 As Single, tU2 As Single, tv1 As Single, tV2 As Single, _
                                Color As Long, Specular As Long) As D3DTLVERTEX2
 Dim tmpVert As D3DTLVERTEX2
 tmpVert.Sx = Sx: tmpVert.Sy = Sy: tmpVert.Sz = Sz
 tmpVert.tu1 = tu1: tmpVert.tv1 = tv1
 tmpVert.tU2 = tU2: tmpVert.tV2 = tV2
 tmpVert.RHW = RHW
 tmpVert.Color = Color
 tmpVert.Specular = Specular
 VertexTL2 = tmpVert
End Function

Public Function VertexTL3(Sx As Single, Sy As Single, Sz As Single, _
                                RHW As Single, tU As Single, tV As Single, Color As Long, Specular As Long) As D3DTLVERTEX
 Dim tmpVert As D3DTLVERTEX
 tmpVert.Sx = Sx: tmpVert.Sy = Sy: tmpVert.Sz = Sz
 tmpVert.tU = tU: tmpVert.tV = tV
 tmpVert.RHW = RHW
 tmpVert.Color = Color
 tmpVert.Specular = Specular
 VertexTL3 = tmpVert
End Function

'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|
'|þþ                     Color Functions                     þþ|
'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|

Public Function Color4(Red As Single, Green As Single, Blue As Single, Alpha As Single) As D3DCOLORVALUE
 Dim Color As D3DCOLORVALUE
 Color.R = Red: Color.G = Green: Color.B = Blue: Color.A = Alpha
 Color4 = Color
End Function

Public Function ColorRGBA(Red As Single, Green As Single, Blue As Single, Alpha As Single) As Long
 ColorRGBA = Color4_To_Long(Color4(Red, Green, Blue, Alpha))
End Function

Public Function Color4_To_Long(Color As D3DCOLORVALUE) As Long
 Dim R As Long
 Dim G As Long
 Dim B As Long
 Dim A As Long
 Dim C As Long
    
 R = Color.R * 255: G = Color.G * 255
 B = Color.B * 255: A = Color.A * 255
    
 If A > 127 Then
  A = A - 128
  C = A * 2 ^ 24 Or &H80000000
  C = C Or R * 2 ^ 16
  C = C Or G * 2 ^ 8
  C = C Or B
 Else
  C = A * 2 ^ 24
  C = C Or R * 2 ^ 16
  C = C Or G * 2 ^ 8
  C = C Or B
 End If
    
 Color4_To_Long = C
End Function

Public Function Color_Long_To_Color4(Color As Long) As D3DCOLORVALUE
 Dim A As Long, R As Long, G As Long, B As Long
        
 If Color < 0 Then
  A = ((Color And (&H7F000000)) / (2 ^ 24)) Or &H80&
 Else
  A = Color / (2 ^ 24)
 End If

 R = (Color And &HFF0000) / (2 ^ 16)
 G = (Color And &HFF00&) / (2 ^ 8)
 B = (Color And &HFF&)
    
 Color_Long_To_Color4.A = A / 255: Color_Long_To_Color4.R = R / 255
 Color_Long_To_Color4.G = G / 255: Color_Long_To_Color4.B = B / 255
End Function

Public Sub Color_Long_To_RGB(LongCol As Long, Red As Single, Green As Single, Blue As Single)
 Red = LongCol And 255
 Green = (LongCol And 65280) \ 256&
 Blue = (LongCol And 16711680) \ 65535
End Sub

Public Function Color_GreyScale(Color As Long) As Single
 Dim R As Single
 Dim G As Single
 Dim B As Single
 Color_Long_To_RGB Color, R, G, B
 Color_GreyScale = (R + B + G) / 765
End Function

Public Function Color_Add(Color_Out As D3DCOLORVALUE, Color1 As D3DCOLORVALUE, Color2 As D3DCOLORVALUE)
 D3DXColorAdd Color_Out, Color1, Color2
End Function

Public Function Color_Adjust_Contrast(Color_Out As D3DCOLORVALUE, Color1 As D3DCOLORVALUE, Contrast As Single)
 D3DXColorAdjustContrast Color_Out, Color1, Contrast
End Function

Public Function Color_Adjust_Saturation(Color_Out As D3DCOLORVALUE, Color1 As D3DCOLORVALUE, Saturation As Single)
 D3DXColorAdjustSaturation Color_Out, Color1, Saturation
End Function

Public Function Color_Lerp(Color_Out As D3DCOLORVALUE, Color1 As D3DCOLORVALUE, Color2 As D3DCOLORVALUE, S As Single)
 D3DXColorLerp Color_Out, Color1, Color2, S
End Function

Public Function Color_Modulate(Color_Out As D3DCOLORVALUE, Color1 As D3DCOLORVALUE, Color2 As D3DCOLORVALUE)
 D3DXColorModulate Color_Out, Color1, Color2
End Function

Public Function Color_Negative(Color_Out As D3DCOLORVALUE, Color1 As D3DCOLORVALUE) As D3DCOLORVALUE
 D3DXColorNegative Color_Out, Color1
End Function

Public Function Color_Scale(Color_Out As D3DCOLORVALUE, Color1 As D3DCOLORVALUE, Scale_Value As Single)
 D3DXColorScale Color_Out, Color1, Scale_Value
End Function

Public Function Color_Subtract(Color_Out As D3DCOLORVALUE, Color1 As D3DCOLORVALUE, Color2 As D3DCOLORVALUE)
 D3DXColorSubtract Color_Out, Color1, Color2
End Function

'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|
'|þþ                    Plane Functions                      þþ|
'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|

Public Function Plane(A As Single, B As Single, C As Single, D As Single) As D3DPLANE
 Plane.A = A: Plane.B = B: Plane.C = C: Plane.D = D
End Function

Public Function Plane_Dot_Product(Plane1 As D3DPLANE, Plane2 As D3DPLANE) As Single
 Plane_Dot_Product = D3DXPlaneDot(Plane1, Plane1)
End Function

Public Function Plane_Dot_Coord(Plane1 As D3DPLANE, Vector As D3DVECTOR) As Single
 Plane_Dot_Coord = D3DXPlaneDotCoord(Plane1, Vector)
End Function

Public Function Plane_Dot_Normal(Plane1 As D3DPLANE, Vector As D3DVECTOR) As Single
 Plane_Dot_Normal = D3DXPlaneDotNormal(Plane1, Vector)
End Function

Public Function Plane_Normalize(Detination_Plane As D3DPLANE, Plane1 As D3DPLANE) As Long
 Plane_Normalize = D3DXPlaneNormalize(Detination_Plane, Plane1)
End Function

Public Function Plane_Transform(Detination_Plane As D3DPLANE, Plane1 As D3DPLANE, Matrix As D3DMATRIX) As Long
 Plane_Transform = D3DXPlaneTransform(Detination_Plane, Plane1, Matrix)
End Function

Public Function Plane_Intersect_Line(Detination_Plane As D3DPLANE, Plane1 As D3DPLANE, Starting_Point As D3DVECTOR, Ending_Point As D3DVECTOR) As Long
 Plane_Intersect_Line = D3DXPlaneIntersectLine(Detination_Plane, Plane1, Starting_Point, Ending_Point)
End Function

Public Function Plane_From_Points(Detination_Plane As D3DPLANE, Point1 As D3DVECTOR, Point2 As D3DVECTOR, Point3 As D3DVECTOR) As Long
 Plane_From_Points = D3DXPlaneFromPoints(Detination_Plane, Point1, Point2, Point3)
End Function

Public Function Plane_From_Point_Normal(Detination_Plane As D3DPLANE, Vector_Point As D3DVECTOR, Vector_Normal As D3DVECTOR) As Long
 Plane_From_Point_Normal = D3DXPlaneFromPointNormal(Detination_Plane, Vector_Point, Vector_Normal)
End Function

'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|
'|þþ                   Quaternion Functions                  þþ|
'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|

Public Function Quaternion(X As Single, Y As Single, Z As Single, W As Single) As D3DQUATERNION
 Quaternion.X = X: Quaternion.Y = Y
 Quaternion.Z = Z: Quaternion.W = W
End Function

Public Function Quaternion_Bary_Centric(Quaternion_Out As D3DQUATERNION, Quaternion1 As D3DQUATERNION, Quaternion2 As D3DQUATERNION, Quaternion3 As D3DQUATERNION, Wieght_Factor_F As Single, Wieght_Factor_G As Single) As Long
 Quaternion_Bary_Centric = D3DXQuaternionBaryCentric(Quaternion_Out, Quaternion1, Quaternion2, Quaternion3, Wieght_Factor_F, Wieght_Factor_G)
End Function

Public Function Quaternion_Conjugate(Quaternion_Out As D3DQUATERNION, Quaternion As D3DQUATERNION) As Long
 Quaternion_Conjugate = D3DXQuaternionConjugate(Quaternion_Out, Quaternion)
End Function

Public Function Quaternion_Dot_Product(Quaternion1 As D3DQUATERNION, Quaternion2 As D3DQUATERNION) As Single
 Quaternion_Dot_Product = D3DXQuaternionDot(Quaternion1, Quaternion2)
End Function

Public Function Quaternion_Exp(Quaternion_Out As D3DQUATERNION, Quaternion As D3DQUATERNION) As Long
 Quaternion_Exp = D3DXQuaternionExp(Quaternion_Out, Quaternion)
End Function

Public Function Quaternion_Identity(Quaternion_Out As D3DQUATERNION) As Long
 Quaternion_Identity = D3DXQuaternionIdentity(Quaternion_Out)
End Function

Public Function Quaternion_Is_Identity(Quaternion As D3DQUATERNION) As Long
 Quaternion_Is_Identity = D3DXQuaternionIsIdentity(Quaternion)
End Function

Public Function Quaternion_Inverse(Quaternion_Out As D3DQUATERNION, Quaternion As D3DQUATERNION) As Long
 Quaternion_Inverse = D3DXQuaternionInverse(Quaternion_Out, Quaternion)
End Function

Public Function Quaternion_Length(Quaternion As D3DQUATERNION) As Single
 Quaternion_Length = D3DXQuaternionLength(Quaternion)
End Function

Public Function Quaternion_Length_Squared(Quaternion As D3DQUATERNION) As Single
 Quaternion_Length_Squared = D3DXQuaternionLengthSq(Quaternion)
End Function

Public Function Quaternion_LN(Quaternion_Out As D3DQUATERNION, Quaternion As D3DQUATERNION) As Long
 Quaternion_LN = D3DXQuaternionLn(Quaternion_Out, Quaternion)
End Function

Public Function Quaternion_Multiply(Quaternion_Out As D3DQUATERNION, Quaternion1 As D3DQUATERNION, Quaternion2 As D3DQUATERNION) As Long
 Quaternion_Multiply = D3DXQuaternionMultiply(Quaternion_Out, Quaternion1, Quaternion2)
End Function

Public Function Quaternion_Normalize(Quaternion_Out As D3DQUATERNION, Quaternion As D3DQUATERNION) As Long
 Quaternion_Normalize = D3DXQuaternionNormalize(Quaternion_Out, Quaternion)
End Function

Public Function Quaternion_Rotation_Axis(Quaternion_Out As D3DQUATERNION, Axis As D3DVECTOR, Angle As Single) As Long
 Quaternion_Rotation_Axis = D3DXQuaternionRotationAxis(Quaternion_Out, Axis, Angle)
End Function

Public Function Quaternion_Rotation_Matrix(Quaternion_Out As D3DQUATERNION, Matrix As D3DMATRIX) As Long
 Quaternion_Rotation_Matrix = D3DXQuaternionRotationMatrix(Quaternion_Out, Matrix)
End Function

Public Function Quaternion_Yaw_Pitch_Roll(Quaternion_Out As D3DQUATERNION, Yaw As Single, Pitch As Single, Roll As Single) As Long
 Quaternion_Yaw_Pitch_Roll = D3DXQuaternionRotationYawPitchRoll(Quaternion_Out, Yaw, Pitch, Roll)
End Function

Public Function Quaternion_Slerp(Quaternion_Out As D3DQUATERNION, Quaternion1 As D3DQUATERNION, Quaternion2 As D3DQUATERNION, t As Single) As Long
 Quaternion_Slerp = D3DXQuaternionSlerp(Quaternion_Out, Quaternion1, Quaternion2, t)
End Function

Public Function Quaternion_Squad(Quaternion_Out As D3DQUATERNION, Quaternion1 As D3DQUATERNION, Quaternion2 As D3DQUATERNION, _
                                 Quaternion3 As D3DQUATERNION, Quaternion4 As D3DQUATERNION, t As Single) As Long
 Quaternion_Squad = D3DXQuaternionSquad(Quaternion_Out, Quaternion1, Quaternion2, Quaternion3, Quaternion4, t)
End Function

Public Function Quaternion_To_Axis_Angle(Quaternion_Out As D3DQUATERNION, Axis As D3DVECTOR, Angle As Single) As Long
 Quaternion_To_Axis_Angle = D3DXQuaternionToAxisAngle(Quaternion_Out, Axis, Angle)
End Function

'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|
'|þþ               Collision Detection Functions             þþ|
'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|

Public Function Get_Distance_2D(Position1 As D3DVECTOR2, Position2 As D3DVECTOR2) As Single
 Get_Distance_2D = (Position2.X - Position1.X) * _
                   (Position2.X - Position1.X) + _
                   (Position2.Y - Position1.Y) * _
                   (Position2.Y - Position1.Y)
End Function

Public Function Get_Distance_3D(Position1 As D3DVECTOR, Position2 As D3DVECTOR) As Single
 Get_Distance_3D = (Position2.X - Position1.X) * _
                   (Position2.X - Position1.X) + _
                   (Position2.Y - Position1.Y) * _
                   (Position2.Y - Position1.Y) + _
                   (Position2.Z - Position1.Z) * _
                   (Position2.Z - Position1.Z)
End Function

Public Function Get_Distance_4D(Position1 As D3DVECTOR4, Position2 As D3DVECTOR4) As Single
 Get_Distance_4D = (Position2.X - Position1.X) * _
                   (Position2.X - Position1.X) + _
                   (Position2.Y - Position1.Y) * _
                   (Position2.Y - Position1.Y) + _
                   (Position2.Z - Position1.Z) * _
                   (Position2.Z - Position1.Z) + _
                   (Position2.W - Position1.W) * _
                   (Position2.W - Position1.W)
End Function

Public Function Check_3D_Object_To_Object_Collision_YZ(Collider_Position As D3DVECTOR, Object_Position As D3DVECTOR, Object_Radius As Single, Optional Radius_OffsetY = 0, Optional Radius_OffsetZ = 0) As Boolean

  If Collider_Position.Y > (Object_Position.Y + Radius_OffsetY) - Object_Radius _
 And Collider_Position.Y < (Object_Position.Y + Radius_OffsetY) + Object_Radius _
 And Collider_Position.Z > (Object_Position.Z + Radius_OffsetZ) - Object_Radius _
 And Collider_Position.Z < (Object_Position.Z + Radius_OffsetZ) + Object_Radius Then
  Check_3D_Object_To_Object_Collision_YZ = True
 Else
  Check_3D_Object_To_Object_Collision_YZ = False
 End If

End Function

Public Function Check_3D_Object_To_Object_Collision_XY(Collider_Position As D3DVECTOR, Object_Position As D3DVECTOR, Object_Radius As Single, Optional Radius_OffsetX = 0, Optional Radius_OffsetY = 0) As Boolean

  If Collider_Position.X > (Object_Position.X + Radius_OffsetX) - Object_Radius _
 And Collider_Position.X < (Object_Position.X + Radius_OffsetX) + Object_Radius _
 And Collider_Position.Y > (Object_Position.Y + Radius_OffsetY) - Object_Radius _
 And Collider_Position.Y < (Object_Position.Y + Radius_OffsetY) + Object_Radius Then
  Check_3D_Object_To_Object_Collision_XY = True
 Else
  Check_3D_Object_To_Object_Collision_XY = False
 End If

End Function

Public Function Check_3D_Object_To_Object_Collision_XZ(Collider_Position As D3DVECTOR, Object_Position As D3DVECTOR, Object_Radius As Single, Optional Radius_OffsetX = 0, Optional Radius_OffsetZ = 0) As Boolean

  If Collider_Position.X > (Object_Position.X + Radius_OffsetX) - Object_Radius _
 And Collider_Position.X < (Object_Position.X + Radius_OffsetX) + Object_Radius _
 And Collider_Position.Z > (Object_Position.Z + Radius_OffsetZ) - Object_Radius _
 And Collider_Position.Z < (Object_Position.Z + Radius_OffsetZ) + Object_Radius Then
  Check_3D_Object_To_Object_Collision_XZ = True
 Else
  Check_3D_Object_To_Object_Collision_XZ = False
 End If

End Function

Public Function Check_3D_Object_To_Object_Collision_XYZ(Collider_Position As D3DVECTOR, Object_Position As D3DVECTOR, Object_Radius As Single, Optional Radius_OffsetX = 0, Optional Radius_OffsetY = 0, Optional Radius_OffsetZ = 0) As Boolean
 
  If Collider_Position.X > (Object_Position.X + Radius_OffsetX) - Object_Radius _
 And Collider_Position.X < (Object_Position.X + Radius_OffsetX) + Object_Radius _
 And Collider_Position.Y > (Object_Position.Y + Radius_OffsetY) - Object_Radius _
 And Collider_Position.Y < (Object_Position.Y + Radius_OffsetY) + Object_Radius _
 And Collider_Position.Z > (Object_Position.Z + Radius_OffsetZ) - Object_Radius _
 And Collider_Position.Z < (Object_Position.Z + Radius_OffsetZ) + Object_Radius Then
  Check_3D_Object_To_Object_Collision_XYZ = True
 Else
  Check_3D_Object_To_Object_Collision_XYZ = False
 End If

End Function

Public Function Check_4D_Object_To_Object_Collision_XYZW(Collider_Position As D3DVECTOR4, Object_Position As D3DVECTOR4, Object_Radius As Single, Optional Radius_OffsetX = 0, Optional Radius_OffsetY = 0, Optional Radius_OffsetZ = 0, Optional Radius_OffsetW = 0) As Boolean
 
  If Collider_Position.X > (Object_Position.X + Radius_OffsetX) - Object_Radius _
 And Collider_Position.X < (Object_Position.X + Radius_OffsetX) + Object_Radius _
 And Collider_Position.Y > (Object_Position.Y + Radius_OffsetY) - Object_Radius _
 And Collider_Position.Y < (Object_Position.Y + Radius_OffsetY) + Object_Radius _
 And Collider_Position.Z > (Object_Position.Z + Radius_OffsetZ) - Object_Radius _
 And Collider_Position.Z < (Object_Position.Z + Radius_OffsetZ) + Object_Radius _
 And Collider_Position.W > (Object_Position.W + Radius_OffsetW) - Object_Radius _
 And Collider_Position.W < (Object_Position.W + Radius_OffsetW) + Object_Radius Then
  Check_4D_Object_To_Object_Collision_XYZW = True
 Else
  Check_4D_Object_To_Object_Collision_XYZW = False
 End If

End Function

Public Function Get_Sphere_Radius_From_Mesh(Mesh As D3DXMesh) As Single
 On Local Error Resume Next
 Dim SR As Single: Dim SC As D3DVECTOR
 D3DX.ComputeBoundingSphereFromMesh Mesh, SC, SR
 Get_Sphere_Radius_From_Mesh = SR
End Function

Public Function Get_Sphere_Center_From_Mesh(Mesh As D3DXMesh) As D3DVECTOR
 On Local Error Resume Next
 Dim SR As Single: Dim SC As D3DVECTOR
 D3DX.ComputeBoundingSphereFromMesh Mesh, SC, SR
 Get_Sphere_Center_From_Mesh = SC
End Function

Public Function Get_Bounding_Box_Min_From_Mesh(Mesh As D3DXMesh) As D3DVECTOR
 On Local Error Resume Next
 Dim Min As D3DVECTOR: Dim Max As D3DVECTOR
 D3DX.ComputeBoundingBoxFromMesh Mesh, Min, Max
 Get_Bounding_Box_Min_From_Mesh = Min
End Function

Public Function Get_Bounding_Box_Max_From_Mesh(Mesh As D3DXMesh) As D3DVECTOR
 On Local Error Resume Next
 Dim Min As D3DVECTOR: Dim Max As D3DVECTOR
 D3DX.ComputeBoundingBoxFromMesh Mesh, Min, Max
 Get_Bounding_Box_Max_From_Mesh = Max
End Function

Public Function Get_Bounding_Box_From_MeshEX(Mesh As D3DXMesh, MinOut As D3DVECTOR, MaxOut As D3DVECTOR)
 On Local Error Resume Next
 Dim tmpMesh As D3DXMesh
 Set tmpMesh = Nothing
 Set tmpMesh = Mesh
 D3DX.ComputeBoundingBoxFromMesh Mesh, MinOut, MaxOut
 Set tmpMesh = Nothing
End Function

Public Function VectorIsInsideBox(Vector As D3DVECTOR, Box_Min As D3DVECTOR, Box_Max As D3DVECTOR) As Boolean
 If Vector.X >= Box_Min.X Or _
    Vector.X <= Box_Max.X Or _
    Vector.Y >= Box_Min.Y Or _
    Vector.Y <= Box_Max.Y Or _
    Vector.Z >= Box_Min.Z Or _
    Vector.Z <= Box_Max.Z Then
  VectorIsInsideBox = True
 Else
  VectorIsInsideBox = False
 End If
End Function

Public Function RayTri(ByRef v0 As D3DVECTOR, ByRef V1 As D3DVECTOR, ByRef V2 As D3DVECTOR, vDir As D3DVECTOR, vOrig As D3DVECTOR, t As Single, u As Single, V As Single, Optional CNormal As Boolean = False) As Long
 Dim edge1 As D3DVECTOR: Dim edge2 As D3DVECTOR
 Dim pvec As D3DVECTOR: Dim tvec As D3DVECTOR
 Dim qvec As D3DVECTOR: Dim Det As Single: Dim fInvDet As Single
 Dim CollisionNormal As D3DVECTOR
 Dim CollisionImpact As D3DVECTOR

 D3DXVec3Subtract edge1, V1, v0
 D3DXVec3Subtract edge2, V2, v0
 D3DXVec3Cross pvec, vDir, edge2
 Det = D3DXVec3Dot(edge1, pvec)
 If Abs(Det < mEPSILON) Then Exit Function
 
 D3DXVec3Subtract tvec, vOrig, v0
 u = D3DXVec3Dot(tvec, pvec)
 If (u < 0 Or u > Det) Then Exit Function
 
 D3DXVec3Cross qvec, tvec, edge1
 V = D3DXVec3Dot(vDir, qvec)
 If (V < 0 Or (u + V > Det)) Then Exit Function
 
 t = D3DXVec3Dot(edge2, qvec)
 fInvDet = 1 / Det
 t = t * fInvDet
 u = u * fInvDet
 V = V * fInvDet
 If t = 0 Then Exit Function
 
 If CNormal = True Then
  D3DXVec3Cross CollisionNormal, Vector3_Normalize(edge1), Vector3_Normalize(edge2)
  D3DXVec3Add CollisionImpact, vOrig, Vector3_Scale(vDir, t)
 End If
 
 RayTri = 1

End Function

'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|
'|þþ                    Misc. Functions                      þþ|
'|þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ|

Public Function D3D_Rect(X1 As Single, X2 As Single, Y1 As Single, Y2 As Single) As D3DRECT
 D3D_Rect.X1 = X1: D3D_Rect.X2 = X2
 D3D_Rect.Y1 = Y1: D3D_Rect.Y2 = Y2
End Function

Public Function Similar_Vertices(Vertex1 As D3DVERTEX, Vertex2 As D3DVERTEX, Offset As Single) As Boolean
 If Abs(Vertex1.X - Vertex2.X) > Offset Then Exit Function
 If Abs(Vertex1.Y - Vertex2.Y) > Offset Then Exit Function
 If Abs(Vertex1.Z - Vertex2.Z) > Offset Then Exit Function
 Similar_Vertices = True
End Function

Public Sub Fix_Prossesor_Float_Error()
 On Local Error GoTo FixIt
 Dim Pro1 As Single
 Dim Pro2 As Single
 
 Pro1 = 1#
 Pro1 = Pro1 / Pro2
FixIt:
 Pro1 = 0
End Sub

Public Function FtoDW(float As Single) As Long
 Dim Buf As D3DXBuffer
 Dim l As Long
 
 Set Buf = D3DX.CreateBuffer(4)
 D3DX.BufferSetData Buf, 0, 4, 1, float
 D3DX.BufferGetData Buf, 0, 4, 1, l
 FtoDW = l
 
 Set Buf = Nothing
End Function

Public Function GetPolygonNormal(vecPolygon1 As D3DVECTOR, vecPolygon2 As D3DVECTOR, vecPolygon3 As D3DVECTOR) As D3DVECTOR
 'Thanks to Richard Hayden for this function
 Dim Vec1 As D3DVECTOR
 Dim Vec2 As D3DVECTOR
 Dim vTemp As D3DVECTOR
 
 Vec1.X = (vecPolygon2.X - vecPolygon1.X)
 Vec1.Y = (vecPolygon2.Y - vecPolygon1.Y)
 Vec1.Z = (vecPolygon2.Z - vecPolygon1.Z)
 Vec2.X = (vecPolygon3.X - vecPolygon1.X)
 Vec2.Y = (vecPolygon3.Y - vecPolygon1.Y)
 Vec2.Z = (vecPolygon3.Z - vecPolygon1.Z)
 vTemp = Vector3_Cross_Product(Vec1, Vec2)
 GetPolygonNormal = vTemp
End Function

Public Function HyperbolicSine(Value As Double) As Double
 Dim temp As Double
 temp = Exp(Value)
 HyperbolicSine = (temp - 1 / temp) / 2
End Function

Public Function HyperbolicCosine(Value As Double) As Double
 Dim temp As Double
 temp = Exp(Value)
 HyperbolicCosine = (temp + 1 / temp) / 2
End Function

Public Function HyperbolicTangent(Value As Double) As Double
 Dim temp As Double
 temp = Exp(Value)
 HyperbolicTangent = (temp - 1 / temp) / (temp + 1 / temp)
End Function

Public Function HyperbolicCotangent(Value As Double) As Double
 Dim temp As Double
 temp = Exp(Value)
 HyperbolicCotangent = (temp + 1 / temp) / (temp - 1 / temp)
End Function

Public Function HyperbolicSecant(Value As Double) As Double
 Dim temp As Double
 temp = Exp(Value)
 HyperbolicSecant = 2 / (temp + 1 / temp)
End Function

Public Function HyperbolicCosecant(Value As Double) As Double
 Dim temp As Double
 temp = Exp(Value)
 HyperbolicCosecant = 2 / (temp - 1 / temp)
End Function

Public Function HyperbolicArcSine(Value As Double) As Double
 HyperbolicArcSine = Log(Value + Sqr(Value * Value + 1))
End Function

Public Function HyperbolicArcCosine(Value As Double) As Double
 If Value < -1 Then Exit Function
 If Value > 1 Then Exit Function
 HyperbolicArcCosine = Log(Value + Sqr(Value * Value - 1))
End Function

Public Function HyperbolicArcTangent(Value As Double) As Double
 If Value <= 0 Then Exit Function
 HyperbolicArcTangent = Log((1 / Value + 1) / (1 / Value - 1)) / 2
End Function

Public Function HyperbolicArcCotangent(Value As Double) As Double
 HyperbolicArcCotangent = Log((Value + 1) / (Value - 1)) / 2
End Function

Public Function HyperbolicArcSecant(Value As Double) As Double
 HyperbolicArcSecant = Log((Sqr(1 - Value * Value) + 1) / Value)
End Function

Public Function HyperbolicArcCosecant(Value As Double) As Double
 HyperbolicArcCosecant = Log((Sgn(Value) * Sqr(1 + Value * Value) + 1) / Value)
End Function

Public Function ArcSine(Value As Double) As Double
 If Value < -1 Then Exit Function
 If Value > 1 Then Exit Function
 If Abs(Value) <> 1 Then
  ArcSine = Atn(Value / Sqr(1 - Value * Value))
 Else
  ArcSine = 1.5707963267949 * Sgn(Value)
 End If
End Function

Public Function ArcCosine(ByVal Number As Double) As Double
 If Number < -1 Then Exit Function
 If Number > 1 Then Exit Function
 If Abs(Number) <> 1 Then
  ArcCosine = 1.5707963267949 - Atn(Number / Sqr(1 - Number * Number))
 ElseIf Number = -1 Then
  ArcCosine = 3.14159265358979
 End If
End Function

Public Function ArcCotangent(Value As Double) As Double
 If Value <= 0 Then Exit Function
 ArcCotangent = Atn(1 / Value)
End Function

Public Function ArcSecant(Value As Double) As Double
 If Value < -1 Then Exit Function
 If Value > 1 Then Exit Function
 ArcSecant = ArcCosine(1 / Value)
End Function

Public Function ArcCosecant(Value As Double) As Double
 If Value < -1 Then Exit Function
 If Value > 1 Then Exit Function
 ArcCosecant = ArcSine(1 / Value)
End Function

Public Function CotangentOfAngle(Rads As Double) As Double
 CotangentOfAngle = 1 / Tan(Rads)
End Function

Public Function SecantOfAngle(Rads As Double) As Double
 SecantOfAngle = 1 / Cos(Rads)
End Function

Public Function CosecantOfAngle(Rads As Double) As Double
 CosecantOfAngle = 1 / Sin(Rads)
End Function

Public Function NumberIsPrime(ByVal Number As Long) As Boolean
 Dim Divisor As Long
 Dim Increment As Long
 Dim MaxDivisor As Long
    
 If Number > 3 Then
  If Number Mod 2 = 0 Then Exit Function
  If Number Mod 3 = 0 Then Exit Function
 End If
    
 Divisor = 5
 Increment = 2
 MaxDivisor = Sqr(Number) + 1
    
 Do Until Divisor > MaxDivisor
  If Number Mod Divisor = 0 Then Exit Function
  Divisor = Divisor + Increment
  Increment = 6 - Increment
 Loop
    
 NumberIsPrime = True
End Function

Function Factorial(ByVal Number As Long) As Double
 Static Result(170) As Double
 Dim i As Long
    
 If Result(0) = 0 Then
  Result(0) = 1
  For i = 1 To 170
   Result(i) = Result(i - 1) * i
  Next
 End If
    
 Factorial = Result(Number)
End Function

Public Function Power2(ByVal Exponent As Long) As Long
 Static Res(0 To 31) As Long
 Dim i As Long
    
 If Exponent < 0 Or Exponent > 31 Then Exit Function
 If Res(0) = 0 Then
  Res(0) = 1
  For i = 1 To 30
   Res(i) = Res(i - 1) * 2
  Next
  Res(31) = &H80000000
 End If
    
 Power2 = Res(Exponent)
End Function

Private Sub Class_Initialize()
 If D3DX Is Nothing Then Set D3DX = New D3DX8
End Sub

Private Sub Class_Terminate()
 Set D3DX = Nothing
End Sub
