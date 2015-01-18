VERSION 5.00
Begin VB.UserControl Socket 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   Picture         =   "Socket.ctx":0000
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   28
End
Attribute VB_Name = "Socket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Don't forget to change CSocketMaster class
'instancing property to PublicNotCreatable

'These are the same events CSocketMaster has
Public Event CloseSck()
Public Event Connect()
Public Event ConnectionRequest(ByVal requestID As Long)
Public Event DataArrival(ByVal bytesTotal As Long)
Public Event Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Public Event SendComplete()
Public Event SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

'Our socket
Private WithEvents cmSocket As CSocketMaster
Attribute cmSocket.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
'create an instance of CSocketMaster
Set cmSocket = New CSocketMaster
End Sub

Private Sub UserControl_Terminate()
'destroy instance of CSocketMaster
Set cmSocket = Nothing
End Sub

Private Sub UserControl_Resize()
'this is used to lock control size
UserControl.Width = 420
UserControl.Height = 420
End Sub


'Control properties. Every time the control is built
'the class instance cmSocket is reset, and so the
'control properties. We use these variables to make
'control properties persistent.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Me.LocalPort = PropBag.ReadProperty("LocalPort", 0)
Me.Protocol = PropBag.ReadProperty("Protocol", 0)
Me.RemoteHost = PropBag.ReadProperty("RemoteHost", "")
Me.RemotePort = PropBag.ReadProperty("RemotePort", 0)
Me.Tag = PropBag.ReadProperty("Tag", "")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "LocalPort", Me.LocalPort, 0
PropBag.WriteProperty "Protocol", Me.Protocol, 0
PropBag.WriteProperty "RemoteHost", Me.RemoteHost, ""
PropBag.WriteProperty "RemotePort", Me.RemotePort, 0
PropBag.WriteProperty "Tag", Me.Tag, ""
End Sub

'From this point we declare all the 'bridge' function
'and properties. The idea is very simple, when user
'call a function we call cmSocket function, when
'cmSocket raises an event we raise an event, when user
'set a property we set cmSocket property, when user
'retrieves a property we retrieve cmSocket property
'and pass the result to user.
'Easy, isn't it?

Private Sub cmSocket_CloseSck()
RaiseEvent CloseSck
End Sub

Private Sub cmSocket_Connect()
RaiseEvent Connect
End Sub

Private Sub cmSocket_ConnectionRequest(ByVal requestID As Long)
RaiseEvent ConnectionRequest(requestID)
End Sub

Private Sub cmSocket_DataArrival(ByVal bytesTotal As Long)
RaiseEvent DataArrival(bytesTotal)
End Sub

Private Sub cmSocket_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
RaiseEvent Error(Number, Description, sCode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub

Private Sub cmSocket_SendComplete()
RaiseEvent SendComplete
End Sub

Private Sub cmSocket_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
RaiseEvent SendProgress(bytesSent, bytesRemaining)
End Sub

Public Property Get RemotePort() As Long
RemotePort = cmSocket.RemotePort
End Property

Public Property Let RemotePort(ByVal lngPort As Long)
cmSocket.RemotePort = lngPort
End Property

Public Property Get RemoteHost() As String
RemoteHost = cmSocket.RemoteHost
End Property

Public Property Let RemoteHost(ByVal strHost As String)
cmSocket.RemoteHost = strHost
End Property

Public Property Get RemoteHostIP() As String
RemoteHostIP = cmSocket.RemoteHostIP
End Property

Public Property Get LocalPort() As Long
LocalPort = cmSocket.LocalPort
End Property

Public Property Let LocalPort(ByVal lngPort As Long)
cmSocket.LocalPort = lngPort
End Property

Public Property Get State() As SockState
State = cmSocket.State
End Property

Public Property Get LocalHostName() As String
LocalHostName = cmSocket.LocalHostName
End Property

Public Property Get LocalIP() As String
LocalIP = cmSocket.LocalIP
End Property

Public Property Get BytesReceived() As Long
BytesReceived = cmSocket.BytesReceived
End Property

Public Property Get SocketHandle() As Long
SocketHandle = cmSocket.SocketHandle
End Property

Public Property Get Tag() As String
Tag = cmSocket.Tag
End Property

Public Property Let Tag(ByVal strTag As String)
cmSocket.Tag = strTag
End Property

Public Property Get Protocol() As ProtocolConstants
Protocol = cmSocket.Protocol
End Property

Public Property Let Protocol(ByVal enmProtocol As ProtocolConstants)
cmSocket.Protocol = enmProtocol
End Property

Public Sub Accept(requestID As Long)
cmSocket.Accept requestID
End Sub

Public Sub Bind(Optional LocalPort As Variant, Optional LocalIP As Variant)
cmSocket.Bind LocalPort, LocalIP
End Sub

Public Sub CloseSck()
cmSocket.CloseSck
End Sub

Public Sub Connect(Optional RemoteHost As Variant, Optional RemotePort As Variant)
cmSocket.Connect RemoteHost, RemotePort
End Sub

Public Sub GetData(ByRef data As Variant, Optional varType As Variant, Optional maxLen As Variant)
cmSocket.GetData data, varType, maxLen
End Sub

Public Sub Listen()
cmSocket.Listen
End Sub

Public Sub PeekData(ByRef data As Variant, Optional varType As Variant, Optional maxLen As Variant)
cmSocket.PeekData data, varType, maxLen
End Sub

Public Sub SendData(data As Variant)
cmSocket.SendData data
End Sub

