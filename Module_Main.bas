Attribute VB_Name = "Module_Main"
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function get_current_computername() As String
  On Error Resume Next
  get_current_computername = ""
  
  Dim sBuffer As String
  sBuffer = String$(255, 0&)
  Call GetComputerName(sBuffer, 255)
  
  get_current_computername = Left$(sBuffer, InStr(vbNull, sBuffer, vbNullChar, vbBinaryCompare) - 1)

End Function

Public Sub Main()
    WriteStdOut get_current_computername()
End Sub

