VERSION 5.00
Begin VB.Form frmKCR 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6675
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8880
   ControlBox      =   0   'False
   DrawWidth       =   3
   Icon            =   "frmKCR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmKCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim i As Long
 
 'These as here because of the boolean values
 'leaving them in the loop, the value will change to fast and may skip
 'we just want on/off
 If XI.Keyboard_KeyState(X_S) <> 0 Then SlowDown = Not SlowDown 'Do we slow down?
 If XI.Keyboard_KeyState(X_H) <> 0 Then ShowHelp = Not ShowHelp 'Do we show the help?
 If XI.Keyboard_KeyState(X_T) <> 0 Then ShowTrails = Not ShowTrails 'Do we render the trails
 
 If XI.Keyboard_KeyState(X_G) <> 0 Then 'Draw the "Ghost balls"
  For i = 0 To Num_Balls
   If Not Balls(i).Ball.Is_Transparent Then
    Balls(i).Ball.Enable_Transparency True
   Else
    Balls(i).Ball.Enable_Transparency False
   End If
  Next
 End If
 
End Sub

Private Sub Form_Load()
 Init_Engine 'Main setup(it all starts somewhere) :)
End Sub
