VERSION 5.00
Begin VB.Form frmSetup 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   334
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Height          =   4335
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Joystick Support"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1215
         Left            =   1800
         TabIndex        =   11
         Top             =   2880
         Width           =   3735
         Begin VB.Label lblFF 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Force Feedback - "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   600
            TabIndex        =   15
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblJoy 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Joystick - "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   600
            TabIndex        =   14
            Top             =   360
            Width           =   840
         End
         Begin VB.Label lblJoy2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Joysticks Found"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   1440
            TabIndex        =   13
            Top             =   360
            Width           =   1605
         End
         Begin VB.Label lblFF2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Not Supported"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   2040
            TabIndex        =   12
            Top             =   720
            Width           =   1185
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "Mode Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1335
         Left            =   1800
         TabIndex        =   4
         Top             =   1400
         Width           =   3735
         Begin VB.CheckBox chkMouse 
            BackColor       =   &H00000000&
            Caption         =   "Show Mouse Cursor"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1320
            TabIndex        =   19
            Top             =   1070
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.TextBox txtWin 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "500"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtWin 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "700"
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cmbModes 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1320
            TabIndex        =   7
            Top             =   720
            Width           =   2295
         End
         Begin VB.OptionButton optWinType 
            BackColor       =   &H00000000&
            Caption         =   "Fullscreen"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   750
            Width           =   1215
         End
         Begin VB.OptionButton optWinType 
            BackColor       =   &H00000000&
            Caption         =   "Windowed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2040
            TabIndex        =   10
            Top             =   360
            Width           =   165
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         Caption         =   "Default Video Card"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   975
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   3735
         Begin VB.Label lblVideo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Video"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   3
            Top             =   600
            Width           =   3240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Video Card - "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   1065
         End
      End
      Begin VB.Image ImgLogo 
         Height          =   3735
         Left            =   120
         Picture         =   "frmSetup.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Height          =   720
      Left            =   90
      TabIndex        =   16
      Top             =   4200
      Width           =   5895
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   300
         Width           =   1335
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   300
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'
  '|¶¶             © 2001-2002 Ariel Productions          ¶¶|'
  '|¶¶¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¶¶|'
  '|¶¶             Programmer - James Dougherty           ¶¶|'
  '|¶¶             Source - frmSetup.frm                  ¶¶|'
  '|¶¶             Object - UltimaX.dll                   ¶¶|'
  '|¶¶             Version - 2.1                          ¶¶|'
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'

  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'
  '|¶¶ NOTE:                                              ¶¶|'
  '|¶¶       This is part of the .dll.                    ¶¶|'
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'

Option Explicit

Private CanC As Boolean
Private OKC As Boolean
Private IsFull As Boolean
Private WWidth As Long
Private WHeight As Long
Private FWidth As Long
Private FHeight As Long

Public Function OkWasClicked() As Boolean
 OkWasClicked = OKC
End Function

Public Function CancelWasClicked() As Boolean
 CancelWasClicked = CanC
End Function

Public Function Fullscreen() As Boolean
 Fullscreen = IsFull
End Function

Public Function Get_Win_Width() As Long
 Get_Win_Width = WWidth
End Function

Public Function Get_Win_Height() As Long
 Get_Win_Height = WHeight
End Function

Public Function Get_FS_Width() As Long
 Get_FS_Width = FWidth
End Function

Public Function Get_FS_Height() As Long
 Get_FS_Height = FHeight
End Function

Public Function Get_Show_Cursor() As Boolean
 Get_Show_Cursor = chkMouse.Value
End Function

Private Sub cmdCancel_Click()
 CanC = True
 OKC = False
 Unload frmSetup
End Sub

Private Sub cmdOk_Click()
 On Local Error Resume Next
 Dim tmpStr As String

 If optWinType(0).Value Then
  WWidth = txtWin(0).Text
  WHeight = txtWin(1).Text
  IsFull = False
 ElseIf optWinType(1).Value Then
  tmpStr = cmbModes.List(cmbModes.ListIndex)
  If Len(tmpStr) = 9 Then
   FWidth = Left$(tmpStr, 3)
   FHeight = Mid$(tmpStr, 7, 3)
  ElseIf Len(tmpStr) = 10 Then
   FWidth = Left$(tmpStr, 4)
   FHeight = Mid$(tmpStr, 8, 3)
  ElseIf Len(tmpStr) = 11 Then
   FWidth = Left$(tmpStr, 4)
   FHeight = Mid$(tmpStr, 8, 4)
  End If
  IsFull = True
 End If
 
 CanC = False
 OKC = True
 DoEvents
 Unload frmSetup
End Sub

Private Sub Form_Load()
 OKC = False
 CanC = False
 IsFull = False
End Sub

Private Sub optWinType_Click(Index As Integer)
 Select Case Index
  Case 0
   txtWin(0).Enabled = True
   txtWin(1).Enabled = True
   cmbModes.Enabled = False
   chkMouse.Enabled = False
  Case 1
   txtWin(0).Enabled = False
   txtWin(1).Enabled = False
   cmbModes.Enabled = True
   chkMouse.Enabled = True
 End Select
End Sub
