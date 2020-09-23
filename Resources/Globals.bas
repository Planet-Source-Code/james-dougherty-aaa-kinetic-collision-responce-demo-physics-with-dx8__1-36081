Attribute VB_Name = "Globals"
Option Explicit

Public D3DD As Direct3DDevice8
Public Declare Function GetCurrentTime Lib "kernel32" Alias "GetTickCount" () As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

