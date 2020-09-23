VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FMain 
   Caption         =   "Passthrough Server"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstLog 
      Height          =   3570
      Left            =   0
      TabIndex        =   6
      Top             =   420
      Width           =   6585
   End
   Begin VB.TextBox txtProxyPort 
      Height          =   285
      Left            =   5100
      TabIndex        =   4
      Text            =   "0"
      Top             =   30
      Width           =   615
   End
   Begin VB.TextBox txtProxyServer 
      Height          =   285
      Left            =   3780
      TabIndex        =   3
      Top             =   30
      Width           =   1245
   End
   Begin VB.CheckBox chkUseProxy 
      Caption         =   "Use Proxy Server"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   60
      Width           =   1635
   End
   Begin VB.Timer tmrClient 
      Index           =   0
      Interval        =   10
      Left            =   1290
      Top             =   450
   End
   Begin VB.Timer tmrServer 
      Index           =   0
      Interval        =   10
      Left            =   870
      Top             =   450
   End
   Begin MSWinsockLib.Winsock sckClient 
      Index           =   0
      Left            =   30
      Top             =   450
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   450
      Top             =   450
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   405
      Left            =   1020
      TabIndex        =   1
      Top             =   0
      Width           =   1035
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "Listen"
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   ":"
      Height          =   285
      Left            =   5040
      TabIndex        =   5
      Top             =   60
      Width           =   195
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'GUI constants
Private Const vbTransparant = &H8000000F

'Proxy Configuration
Dim ListeningPort As Long
Dim UseProxyServer As Boolean
Dim ProxyServer As String
Dim ProxyPort As String

'Buffer Collection
Dim ServerConnection As Collection
Dim ClientConnection As Collection

Private Sub chkUseProxy_Click()
    If chkUseProxy.Value = vbChecked Then
        UseProxyServer = True
        txtProxyServer.Enabled = True
        txtProxyPort.Enabled = True
        txtProxyServer.BackColor = vbWhite
        txtProxyPort.BackColor = vbWhite
    Else
        UseProxyServer = False
        txtProxyServer.Enabled = False
        txtProxyPort.Enabled = False
        txtProxyServer.BackColor = vbTransparant
        txtProxyPort.BackColor = vbTransparant
    End If
End Sub

Private Sub cmdClear_Click()
    lstLog.Clear
End Sub

Private Sub cmdSwitch_Click()
Dim Socket As Winsock

    If cmdSwitch.Caption = "Listen" Then
        'Starting HTTP Passthrough Server.
        SetProxy txtProxyServer.Text, Val(txtProxyPort.Text)
        InitializeSocket sckServer(0)
        ListeningPort = Val(InputBox("Enter listening port", "Port required", 8080))
        sckServer(0).LocalPort = ListeningPort
        sckServer(0).Listen
        SendToLog "Server start listening on port " & sckServer(0).LocalPort
        cmdSwitch.Caption = "Stop"
        SetScreen False
    Else
        'Shutdown HTTP Passthrough Server.
        InitializeSocket sckServer(0)
        SendToLog "Shutting down Server"
        cmdSwitch.Caption = "Listen"
        SetScreen True
    
        For Each Socket In sckServer
            CloseSocket Socket.Index
        Next
        SendToLog "Server shutdown"
    End If
End Sub

Private Sub SetScreen(Flag As Boolean)
    chkUseProxy.Enabled = Flag
    txtProxyServer.Enabled = Flag
    txtProxyPort.Enabled = Flag
End Sub

Private Sub InitializeSocket(Socket As Winsock)
On Error Resume Next

    Socket.Close
    Socket.LocalPort = 0
End Sub

Private Sub SendToLog(Message As String)
    lstLog.AddItem Message
    If lstLog.ListCount > 16384 Then lstLog.Clear
End Sub

Private Sub Form_Load()
    
    chkUseProxy = vbUnchecked
    UseProxyServer = False
    txtProxyServer.Enabled = False
    txtProxyPort.Enabled = False
    txtProxyServer.BackColor = vbTransparant
    txtProxyPort.BackColor = vbTransparant
    
    'Initialize Buffer collection.
    Set ServerConnection = New Collection
    Set ClientConnection = New Collection
End Sub

Private Sub Form_Resize()
    lstLog.Width = Me.ScaleWidth
    If Me.ScaleHeight > (lstLog.Top - 100) Then
        lstLog.Height = Me.ScaleHeight - (lstLog.Top - 100)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Socket As Winsock

    'Unload Sockets.
    For Each Socket In sckServer
        CloseSocket Socket.Index
        If Socket.Index <> 0 Then Unload Socket
    Next
    
    For Each Socket In sckClient
        CloseSocket Socket.Index
        If Socket.Index <> 0 Then Unload Socket
    Next
    
    Set ServerConnection = Nothing
    Set ClientConnection = Nothing
End Sub

Private Sub SetProxy(HostName As String, HostPort As Long)
    ProxyServer = HostName
    ProxyPort = HostPort
End Sub

Private Sub sckClient_Connect(Index As Integer)
Dim Data As String

    'if connected then send data from buffer.
    If sckClient(Index).State = sckConnected Then
        Data = ClientConnection(Index).GetBuffer
        sckClient(Index).SendData Data
    End If
End Sub

Private Sub sckClient_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Data As String
Dim i As Long
Dim lpos As Long

    'receive data from server and send it to buffer.
    If Index <> 0 And sckClient(Index).State = sckConnected Then
        sckClient(Index).GetData Data
        SendToLog "Receive data from server " & sckClient(Index).RemoteHostIP & " size: " & bytesTotal & " bytes"
        ServerConnection(Index).Append Data
    End If
End Sub

Private Sub sckClient_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CloseSocket Index
End Sub

Private Sub sckServer_Close(Index As Integer)
    CloseSocket Index
End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim Socket As Winsock
    If Index = 0 Then
        'accept connection request.
        Set Socket = AvailableSocket
        Socket.Accept requestID
        SendToLog "Accept connection request from client " & Socket.RemoteHostIP
    End If
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim i As Long
Dim Data As String
Dim tmpData As String

    'receive data from client and send it to buffer.
    If Index <> 0 And sckServer(Index).State = sckConnected Then
        sckServer(Index).GetData Data
        SendToLog "Receive data from client " & sckServer(Index).RemoteHostIP & " size: " & bytesTotal & " bytes"
        ClientConnection(Index).Append Data
    End If
End Sub

Private Sub sckServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CloseSocket Index
End Sub

Private Function AvailableSocket() As Winsock
Dim ServerData As New CBuffer
Dim ClientData As New CBuffer
Dim Socket As Winsock
Dim NewSocket As Long

    'get closed socket for new connection.
    For Each Socket In sckServer
        If Socket.State = sckClosed Then
            ServerConnection(Socket.Index).Clear
            ClientConnection(Socket.Index).Clear
            Set AvailableSocket = Socket
            Exit Function
        End If
    Next
    
    'if no closed socket then load a new one.
    NewSocket = sckServer.Count
    Load sckServer(NewSocket)
    Load sckClient(NewSocket)
    Load tmrServer(NewSocket)
    Load tmrClient(NewSocket)
    ServerData.Clear
    ClientData.Clear
    ServerConnection.Add ServerData, Chr(NewSocket)
    ClientConnection.Add ClientData, Chr(NewSocket)
    Set AvailableSocket = sckServer(NewSocket)
End Function

Private Sub tmrClient_Timer(Index As Integer)
Dim i As Long
Dim Data As String

    'if there is data in the buffer then try to connect to server and send it.
    If Index <> 0 Then
        i = Index
        Data = ClientConnection(i).PeekBuffer
        If Len(Data) <> 0 Then
            DoEvents
            If sckClient(i).State <> sckConnected And sckClient(i).State <> sckConnecting Then
                If UseProxyServer Then
                    sckClient(i).Connect ProxyServer, ProxyPort
                Else
                    sckClient(i).Connect ClientConnection(i).Server, ClientConnection(i).Port
                End If
            End If
            DoEvents
            Data = ClientConnection(i).PeekBuffer
            If sckClient(i).State = sckConnected And Len(Data) <> 0 Then
                Data = ClientConnection(i).GetBuffer
                sckClient(i).SendData Data
            End If
        End If
    End If
End Sub

Private Sub tmrServer_Timer(Index As Integer)
Dim i As Long
Dim Data As String

    'if there is data in the buffer then send it.
    If Index <> 0 Then
        i = Index
        Data = ServerConnection(i).PeekBuffer
        If sckServer(i).State = sckConnected And Len(Data) <> 0 Then
            Data = ServerConnection(i).GetBuffer
            sckServer(i).SendData Data
        End If
    End If
End Sub

Private Sub CloseSocket(Index As Integer)
    InitializeSocket sckServer(Index)
    If Index <> 0 Then
        ClientConnection(Index).Clear
    End If
    
    InitializeSocket sckClient(Index)
    If Index <> 0 Then
        ServerConnection(Index).Clear
    End If
End Sub

Private Sub txtProxyPort_GotFocus()
    txtProxyPort.SelStart = 0
    txtProxyPort.SelLength = Len(txtProxyPort.Text)
End Sub

Private Sub txtProxyPort_LostFocus()
    txtProxyPort = Abs(Val(txtProxyPort))
End Sub

Private Sub txtProxyServer_GotFocus()
    txtProxyServer.SelStart = 0
    txtProxyServer.SelLength = Len(txtProxyServer.Text)
End Sub
