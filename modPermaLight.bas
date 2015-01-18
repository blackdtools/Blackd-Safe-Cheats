Attribute VB_Name = "modPermaLight"
#Const FinalMode = 1
Option Explicit
Private LastEnabledPID As Long
Private OriginalCodeBackup As String

Public LIGHT_TRICK_ADR As Long
Public LIGHT_TRICK_CODE As String

Public Sub InitTibiaPermaLight()
    LastEnabledPID = -1
    OriginalCodeBackup = ""
End Sub

Public Sub SetTibiaPermaLight(pid As Long, active As Boolean)
    If pid = -1 Then
        Exit Sub
    End If
   
    If active = True Then
        If pid = LastEnabledPID Then
            WriteCodeAtAddress pid, LIGHT_TRICK_CODE, False
        Else
            LastEnabledPID = pid
            WriteCodeAtAddress pid, LIGHT_TRICK_CODE, True
        End If
    Else
        If pid = LastEnabledPID Then
          WriteCodeAtAddress pid, OriginalCodeBackup, False
        End If
    End If
End Sub

Private Sub WriteCodeAtAddress(ByRef pid As Long, ByVal trickCode As String, ByVal saveRecovery As Boolean)
    Dim bytepart As String
    Dim byteb As Byte
    Dim remaininglen As Long
    Dim i As Long
    Dim byteR As Byte
    i = 0
    If saveRecovery = True Then
        OriginalCodeBackup = ""
    End If
    Do
        remaininglen = Len(trickCode)
        If remaininglen > 1 Then
            bytepart = left$(trickCode, 2)
            trickCode = Right$(trickCode, remaininglen - 2)
            byteb = CByte(CLng("&H" & bytepart))
            If saveRecovery = True Then
               byteR = Memory_ReadByte(LIGHT_TRICK_ADR + i, pid)
               OriginalCodeBackup = OriginalCodeBackup & GoodHex(byteR)
            End If
            Memory_WriteByte LIGHT_TRICK_ADR + i, byteb, pid
            'Debug.Print "Wrote " & GoodHex(byteb)
            i = i + 1
        End If
    Loop Until remaininglen < 2
End Sub
