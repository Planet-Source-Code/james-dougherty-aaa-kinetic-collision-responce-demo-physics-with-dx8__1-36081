Attribute VB_Name = "ERROR_CATCHER"
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'
  '|¶¶             © 2001-2002 Ariel Productions          ¶¶|'
  '|¶¶¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¶¶|'
  '|¶¶             Programmer - James Dougherty           ¶¶|'
  '|¶¶             Source - ERROR_CATCHER.bas             ¶¶|'
  '|¶¶             Object - UltimaX_SoundEngine.dll       ¶¶|'
  '|¶¶             Version - 2.1                          ¶¶|'
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'

Option Explicit

Private FSys As New FileSystemObject
Private OutStream As TextStream
Public CausedAt As String
Public WhereToLook As String

Public Sub ErrorToFile(OutputFileName As String, CausedAt As String, Optional WhereToLook As String = "")
 On Local Error Resume Next
 Dim tmpHeader As String
 Dim tmpStr1 As String
 Dim tmpStr2 As String

 Set OutStream = FSys.CreateTextFile(OutputFileName & ".txt", True, False)

 tmpHeader = "UltimaX Error - " & Format$(Time, "HH:MM:SS") & " " & Format$(Date, "MM/DD/YY") & " | "
 tmpStr1 = "Error Caused At - "
 tmpStr2 = "Please See " & WhereToLook & " For Details"
 
 OutStream.WriteLine tmpHeader & ""
 OutStream.WriteLine tmpHeader & tmpStr1 & CausedAt
 OutStream.WriteLine tmpHeader & ""
 OutStream.WriteLine tmpHeader & tmpStr2
 OutStream.WriteLine tmpHeader & ""
 OutStream.WriteLine tmpHeader & "Internal Error Source - " & Err.Source
 OutStream.WriteLine tmpHeader & ""
 OutStream.WriteLine tmpHeader & "Internal Error Description - " & Err.Description
 OutStream.WriteLine tmpHeader & ""
 OutStream.WriteLine tmpHeader & "Internal Error Number - " & Err.Number
 OutStream.WriteLine tmpHeader & ""

 Set OutStream = Nothing
End Sub
