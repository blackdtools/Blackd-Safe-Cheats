Attribute VB_Name = "modMac"
#Const FinalMode = 1
Option Explicit
' This module get the MAC address of computer
' for the version that only should run in 1 computer
Private Const NCBASTAT As Long = &H33
Private Const NCBNAMSZ As Long = 16
Private Const HEAP_ZERO_MEMORY As Long = &H8
Private Const HEAP_GENERATE_EXCEPTIONS As Long = &H4
Private Const NCBRESET As Long = &H32
Public Const Trial_MAC1 = &H0 ' 1st byte of MAC address
Public Const Trial_MAC2 = &H80 ' 2nd byte of MAC address
Public Const Trial_MAC3 = &HC8 ' 3rd byte of MAC address
Public Const Trial_MAC4 = &H2C ' 4th byte of MAC address
Public Const Trial_MAC5 = &H44 ' 5th byte of MAC address
Public Const Trial_MAC6 = &H89 ' 6th byte of MAC address
Private Type NET_CONTROL_BLOCK  'NCB
   ncb_command    As Byte
   ncb_retcode    As Byte
   ncb_lsn        As Byte
   ncb_num        As Byte
   ncb_buffer     As Long
   ncb_length     As Integer
   ncb_callname   As String * NCBNAMSZ
   ncb_name       As String * NCBNAMSZ
   ncb_rto        As Byte
   ncb_sto        As Byte
   ncb_post       As Long
   ncb_lana_num   As Byte
   ncb_cmd_cplt   As Byte
   ncb_reserve(9) As Byte 'Reserved, must be 0
   ncb_event      As Long
End Type

Private Type ADAPTER_STATUS
   adapter_address(5) As Byte
   rev_major         As Byte
   reserved0         As Byte
   adapter_type      As Byte
   rev_minor         As Byte
   duration          As Integer
   frmr_recv         As Integer
   frmr_xmit         As Integer
   iframe_recv_err   As Integer
   xmit_aborts       As Integer
   xmit_success      As Long
   recv_success      As Long
   iframe_xmit_err   As Integer
   recv_buff_unavail As Integer
   t1_timeouts       As Integer
   ti_timeouts       As Integer
   Reserved1         As Long
   free_ncbs         As Integer
   max_cfg_ncbs      As Integer
   max_ncbs          As Integer
   xmit_buf_unavail  As Integer
   max_dgram_size    As Integer
   pending_sess      As Integer
   max_cfg_sess      As Integer
   max_sess          As Integer
   max_sess_pkt_size As Integer
   name_count        As Integer
End Type
   
Private Type NAME_BUFFER
   name        As String * NCBNAMSZ
   name_num    As Integer
   name_flags  As Integer
End Type

Private Type ASTAT
   adapt          As ADAPTER_STATUS
   NameBuff(30)   As NAME_BUFFER
End Type

Private Declare Function Netbios Lib "netapi32" _
   (pncb As NET_CONTROL_BLOCK) As Byte
     
Private Declare Sub CopyMemory Lib "Kernel32" _
   Alias "RtlMoveMemory" _
  (hpvDest As Any, ByVal _
   hpvSource As Long, ByVal _
   cbCopy As Long)
     
Private Declare Function GetProcessHeap Lib "Kernel32" () As Long

Private Declare Function HeapAlloc Lib "Kernel32" _
  (ByVal hHeap As Long, _
   ByVal dwFlags As Long, _
   ByVal dwBytes As Long) As Long
     
Private Declare Function HeapFree Lib "Kernel32" _
  (ByVal hHeap As Long, _
   ByVal dwFlags As Long, _
   lpMem As Any) As Long



Public Function DoMACAddressCompare() As Boolean
   'get my mac
   #If FinalMode Then
   On Error GoTo giveError
   #End If
   Dim res As Boolean
   Dim macbytes(0 To 5) As Byte
   Dim pASTAT As Long
   Dim NCB As NET_CONTROL_BLOCK
   Dim AST As ASTAT
   res = False
   NCB.ncb_command = NCBRESET
   Call Netbios(NCB)
   NCB.ncb_callname = "* "
   NCB.ncb_command = NCBASTAT
   NCB.ncb_lana_num = 0
   NCB.ncb_length = Len(AST)
   pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS _
            Or HEAP_ZERO_MEMORY, NCB.ncb_length)
   If pASTAT = 0 Then
     ' Debug.Print "memory allocation failed!"
      Exit Function
   End If
   NCB.ncb_buffer = pASTAT
   Call Netbios(NCB)
   CopyMemory AST, NCB.ncb_buffer, Len(AST)
   macbytes(0) = AST.adapt.adapter_address(0)
   macbytes(1) = AST.adapt.adapter_address(1)
   macbytes(2) = AST.adapt.adapter_address(2)
   macbytes(3) = AST.adapt.adapter_address(3)
   macbytes(4) = AST.adapt.adapter_address(4)
   macbytes(5) = AST.adapt.adapter_address(5)
   If ((macbytes(0) = Trial_MAC1) And _
       (macbytes(1) = Trial_MAC2) And _
       (macbytes(2) = Trial_MAC3) And _
       (macbytes(3) = Trial_MAC4) And _
       (macbytes(4) = Trial_MAC5) And _
       (macbytes(5) = Trial_MAC6)) Then
     res = True
   End If
   HeapFree GetProcessHeap(), 0, pASTAT
   DoMACAddressCompare = res
   Exit Function
giveError:
   DoMACAddressCompare = False
End Function

Public Function MyMac() As String
   'get my mac
   #If FinalMode Then
   On Error GoTo giveError
   #End If
   Dim tmp As String
   Dim macbytes(0 To 5) As Byte
   Dim pASTAT As Long
   Dim NCB As NET_CONTROL_BLOCK
   Dim AST As ASTAT
   NCB.ncb_command = NCBRESET
   Call Netbios(NCB)
   NCB.ncb_callname = "* "
   NCB.ncb_command = NCBASTAT
   NCB.ncb_lana_num = 0
   NCB.ncb_length = Len(AST)
   pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS _
            Or HEAP_ZERO_MEMORY, NCB.ncb_length)
   If pASTAT = 0 Then
      'Debug.Print "memory allocation failed!"
      Exit Function
   End If
   NCB.ncb_buffer = pASTAT
   Call Netbios(NCB)
   CopyMemory AST, NCB.ncb_buffer, Len(AST)
   macbytes(0) = AST.adapt.adapter_address(0)
   macbytes(1) = AST.adapt.adapter_address(1)
   macbytes(2) = AST.adapt.adapter_address(2)
   macbytes(3) = AST.adapt.adapter_address(3)
   macbytes(4) = AST.adapt.adapter_address(4)
   macbytes(5) = AST.adapt.adapter_address(5)
   tmp = GoodHex(macbytes(0)) & ":" & _
         GoodHex(macbytes(1)) & ":" & _
         GoodHex(macbytes(2)) & ":" & _
         GoodHex(macbytes(3)) & ":" & _
         GoodHex(macbytes(4)) & ":" & _
         GoodHex(macbytes(5))
   HeapFree GetProcessHeap(), 0, pASTAT
   MyMac = tmp
   Exit Function
giveError:
   MyMac = "00:00:00:00:00:03"
End Function


