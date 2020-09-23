VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Buffer As String

Public Header As String
Public HeaderReceived As Boolean

Public Sub Clear()
    'Initialize Buffer
    Buffer = ""
    Header = ""
    HeaderReceived = False
End Sub

Public Sub Append(Data As String)
    'Append additional data into buffer
    If Not HeaderReceived Then
        Header = Header & Data
        If IsHTTPHeader(Header) Then
            'for http to get header information
            Buffer = Header
            Header = Left$(Header, InStr(1, Buffer, vbCrLf & vbCrLf, vbTextCompare) + 1)
            HeaderReceived = True
        End If
    Else
        Buffer = Buffer & Data
    End If
End Sub

Public Function GetBuffer() As String
    'read buffer then clear it
    GetBuffer = Buffer
    Buffer = ""
End Function

Public Function PeekBuffer() As String
    'read buffer
    PeekBuffer = Buffer
End Function

Public Function Server() As String
Dim lpos As Long
    'get host server from http header
    Server = GetHttpHeader("Host")
    lpos = InStr(1, Server, ":", vbTextCompare)
    If lpos <> 0 Then
        Server = Left$(Server, lpos - 1)
    End If
End Function

Public Function Port() As Long
Dim lpos As Long, vServer As String
    'get host port from http header
    Port = 80
    vServer = GetHttpHeader("Host")
    lpos = InStr(1, vServer, ":", vbTextCompare)
    If lpos <> 0 Then
        Port = Val(Mid$(vServer, lpos + 1))
    End If
End Function

Public Function GetHttpHeader(HeaderCaption As String) As String
Dim lpos As Long
Dim endpos As Long
    'get a specified header caption from http header
    lpos = InStr(1, Header, HeaderCaption & ":", vbTextCompare)
    If lpos <> 0 Then
        endpos = InStr(lpos + 1, Header, vbCrLf, vbTextCompare)
        GetHttpHeader = Mid$(Header, lpos + Len(HeaderCaption) + 2, endpos - (lpos + Len(HeaderCaption) + 2))
    End If
    
End Function

Public Function IsHTTPHeader(Data As String) As Boolean
Const HEADER_HTTP = "HTTP"
Const METHOD_GET = "GET"
Const METHOD_POST = "POST"
Const METHOD_HEAD = "HEAD"
Dim lpos As Long

    'checking for http header
    Select Case Trim(UCase(Left$(Data, 4)))
    Case HEADER_HTTP, METHOD_GET, METHOD_POST, METHOD_HEAD
        If InStr(1, Data, vbCrLf & vbCrLf, vbTextCompare) Then
            IsHTTPHeader = True
        End If
    End Select
End Function

