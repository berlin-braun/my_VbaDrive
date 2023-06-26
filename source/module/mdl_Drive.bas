Attribute VB_Name = "mdl_Drive"
Option Compare Database
Option Explicit


Public Function drive_Last_Letter(Optional str_Last_Letter As String = "Z" _
                                , Optional str_First_Letter As String = "F") As String
  Dim lng_Chr       As Long
  Dim str_Ret       As String

  str_Ret = ""
  
  For lng_Chr = Asc(str_Last_Letter) To Asc(str_First_Letter) Step -1             ' iterate from last to first backwards
    
    If drive_Available(Chr(lng_Chr)) = True Then                                  ' check drive availability
      str_Ret = Chr(lng_Chr)                                                      ' letter available
      Exit For                                                                    ' letter found
    End If
    
  Next
  
  drive_Last_Letter = str_Ret

End Function

Public Function drive_Available(ByVal str_Letter As String) As Boolean
  Dim bln_Ret     As Boolean
  
  ' available = does not exist
  bln_Ret = Not drive_Exists(str_Letter)                                          ' check fso
  
  If bln_Ret = True Then                                                          ' fso says, it's available
    bln_Ret = Not drive_Share_Exists(str_Letter)                                  ' check network
  End If
  
  drive_Available = bln_Ret

End Function

Public Function drive_Exists(str_Letter As String) As Boolean
  Dim fso         As Object
  Dim bln_Ret     As Boolean
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  
  bln_Ret = fso.DriveExists(str_Letter)
  
  Set fso = Nothing
  
  drive_Exists = bln_Ret
End Function

Public Function drive_Share_Exists(str_Letter As String) As Boolean
  Dim bln_Ret     As Boolean
  Dim str_Tmp     As String
  Dim ws          As Object
  
On Error GoTo my_Error

  bln_Ret = False
  str_Tmp = ""
  
  Set ws = CreateObject("WScript.Shell")
  str_Tmp = ws.RegRead("HKEY_CURRENT_USER\Network\" & str_Letter & "\RemotePath")
  
  If Len(str_Tmp) > 0 Then
    bln_Ret = True
  End If

my_Exit:
  
  drive_Share_Exists = bln_Ret
  
  On Error GoTo 0
  Exit Function

my_Error:
  Dim str_Error  As String
  Dim lng_Error  As Long

  lng_Error = Err.Number
  str_Error = "Error " & Err.Number & " (" & Err.Description & ") in procedure drive_Share_Exists of Modul mod_Drive"

  Select Case lng_Error
    Case Is = 0: Resume Next
    Case Is = -2147024894 ': Resume Next            ' Key nor found
    Case Else:
      MsgBox str_Error
  
  End Select
  
  GoTo my_Exit

End Function



