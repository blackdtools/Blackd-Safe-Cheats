VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmWinsock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Winsock"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3360
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock wsPoP2 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Communication"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmWinsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit
Private Sub wsPoP2_Connect()
  Dim theIP As String
  Dim r As Long
  blnConnected = True
 ' ValidateIP wsPoP2.RemoteHost
End Sub

Private Sub wsPoP2_Close()
  blnConnected = False
 ' wsPoP.Close
  'DoEvents
End Sub

Private Sub wsPoP2_DataArrival(ByVal bytesTotal As Long)
  On Error GoTo closeit
  Dim strdata As String
  Dim justGotData As String
  If bytesTotal = 0 Then
    wsPoP2.Close
    blnConnected = False
  Else
    wsPoP2.GetData strdata
    justGotData = strdata
    webReceived = webReceived & justGotData
  End If
  Exit Sub
closeit:
  wsPoP2.Close
  blnConnected = False
End Sub
