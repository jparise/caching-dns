VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Name Server Status"
   ClientHeight    =   7770
   ClientLeft      =   4485
   ClientTop       =   1815
   ClientWidth     =   5325
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   5325
   Begin VB.Timer cacheTimer 
      Interval        =   10000
      Left            =   4800
      Top             =   3480
   End
   Begin MSWinsockLib.Winsock wsDNSclient 
      Left            =   4200
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "129.21.3.17"
      RemotePort      =   53
   End
   Begin VB.CommandButton cmdClearRequests 
      Caption         =   "Clear Requests"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdClearCache 
      Caption         =   "Clear Cache"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   7320
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock wsDNSserver 
      Left            =   3600
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   53
   End
   Begin VB.CommandButton cmdConfig 
      Caption         =   "Configuration"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   7320
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid grdRequests 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5530
      _Version        =   393216
      Rows            =   13
      Cols            =   3
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdCache 
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5530
      _Version        =   393216
      Rows            =   15
      Cols            =   3
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label lblRequests 
      Caption         =   "Request Log:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblCache 
      Caption         =   "Current Cache:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   2175
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Caching DNS Server
'
' Copyright 2000, 2001 Jon Parise <jon@csh.rit.edu>.  All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions
' are met:
' 1. Redistributions of source code must retain the above copyright
'    notice, this list of conditions and the following disclaimer.
' 2. Redistributions in binary form must reproduce the above copyright
'    notice, this list of conditions and the following disclaimer in the
'    documentation and/or other materials provided with the distribution.
'
' THIS SOFTWARE IS PROVIDED BY THE AUTHOR AND CONTRIBUTORS ``AS IS'' AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
' IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
' ARE DISCLAIMED.  IN NO EVENT SHALL THE AUTHOR OR CONTRIBUTORS BE LIABLE
' FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
' DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS
' OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
' HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT
' LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY
' OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF
' SUCH DAMAGE.

Option Explicit

' The original client... this is a tremendous hack to get around the
' event-driven model.
Private Type ReplylHostType
    Waiting As Boolean
    Id As Integer
    IP As String
    Port As Integer
End Type
Dim ReplyHost As ReplylHostType

' Create a new Cache object
Dim Cache As New CCache

' These two counters record the last row entered in their respective grids.
' They're used purely for cosmetic purposes.
Dim LastRequest As Integer
Dim LastCache As Integer

'--[ Setup Procedures ]--------------------------------------------------------

Private Sub SetupRequestsGrid()
    With grdRequests
        .Clear
        .ColAlignment(0) = 1
        .TextMatrix(0, 0) = "Timestamp"
        .ColWidth(0) = 1000
        .ColAlignment(1) = 1
        .TextMatrix(0, 1) = "From"
        .ColWidth(1) = 1500
        .ColAlignment(2) = 1
        .TextMatrix(0, 2) = "Query"
        .ColWidth(2) = 2255
    End With
    LastRequest = 0
End Sub

Private Sub SetupCacheGrid()
    With grdCache
        .Clear
        .ColAlignment(0) = 1
        .TextMatrix(0, 0) = "Hostname"
        .ColWidth(0) = 1700
        .ColAlignment(1) = 1
        .TextMatrix(0, 1) = "IP Address"
        .ColWidth(1) = 1300
        .ColAlignment(2) = 1
        .TextMatrix(0, 2) = "Expiration"
        .ColWidth(2) = 1755
    End With
    LastCache = 0
End Sub

'--[ Form Routines ]-----------------------------------------------------------

Private Sub Form_Load()
    wsDNSserver.Bind 53
    
    ' Initialize the ReplyHost
    ReplyHost.Waiting = False

    ' Set up the grids
    SetupRequestsGrid
    SetupCacheGrid
End Sub

'--[ Commands ]----------------------------------------------------------------

Private Sub cmdClearCache_Click()
    Cache.Flush
    RedrawCache
End Sub

Private Sub cmdClearRequests_Click()
    ReDim Requests(0)
    SetupRequestsGrid
End Sub

Private Sub cmdConfig_Click()
    frmConfig.Show
End Sub

'--[ User Interface Housekeeping ]---------------------------------------------

Private Sub RedrawCache()
    Dim i As Integer
    
    SetupCacheGrid
    If Cache.numEntries() >= grdCache.Rows Then
        grdCache.Rows = Cache.numEntries() + 1
    End If
    
    For i = 1 To Cache.numEntries()
        grdCache.TextMatrix(i, 0) = Cache.getName(i)
        grdCache.TextMatrix(i, 1) = IPString(Cache.getAddress(i))
        grdCache.TextMatrix(i, 2) = Cache.getExpiration(i)
    Next
    
End Sub

' Adds a new entry to the to request list.
Private Sub AddRequestEntry(ByVal From As String, ByVal Query As String)
    LastRequest = LastRequest + 1
    If grdRequests.Rows <= LastRequest Then
        grdRequests.Rows = LastRequest + 1
    End If
    grdRequests.TextMatrix(LastRequest, 0) = Time()
    grdRequests.TextMatrix(LastRequest, 1) = From
    grdRequests.TextMatrix(LastRequest, 2) = Query
End Sub

'--[ Cache Routines ]----------------------------------------------------------

' Remove expired entries from the cache.
Private Sub cacheTimer_Timer()
    Cache.Expire
    RedrawCache
End Sub

' Add a new entry to the cache.
Private Sub AddCacheEntry(ByVal Name As String, ByVal Address As String, ByVal TTL As Long)
    Cache.Add Name, Address, TTL
    RedrawCache
End Sub

'--[ Winsock Routines ]--------------------------------------------------------

' Handle an incoming request.
Private Sub wsDNSserver_DataArrival(ByVal bytesTotal As Long)
    Dim packet As String
    Dim i As Long
    Dim Request As CDNSPacket
    
    ' Grab the request packet.
    wsDNSserver.GetData packet
    
    ' Break down the packet into our data structures.
    Set Request = New CDNSPacket
    Request.ParsePacket packet
    
    ' Handle the each request (questions).
    For i = 1 To Request.GetQDCount()
        ' Make sure we have an IP address A record in the query.
        If Request.GetQType(i) = 1 And Request.GetQClass(i) = 1 Then
            AddRequestEntry wsDNSserver.RemoteHostIP, Request.GetQName(i)
            ReplyHost.Id = Request.GetID()
            ReplyHost.IP = wsDNSserver.RemoteHostIP
            ReplyHost.Port = wsDNSserver.RemotePort
            ReplyHost.Waiting = True
            If Cache.hasName(Request.GetQName(i)) = True Then
                SendReply Request.GetQName(i)
            Else
                Lookup Request.GetQName(i), frmConfig!txtForwarder.Text
            End If
        End If
    Next i

    ' Destroy the request object.
    Set Request = Nothing

End Sub

' Handle an incoming response.
Private Sub wsDNSclient_DataArrival(ByVal bytesTotal As Long)
    Dim packet As String
    
    ' Grab the result from Winsock
    wsDNSclient.GetData packet

    ' Parse the packet
    ParseReply packet
End Sub

'--[ Request Handling ]--------------------------------------------------------

' Perform a lookup of the given name using the given server.
Private Sub Lookup(ByVal Name As String, ByVal server As String)
    Dim packet As String
    Dim Request As CDNSPacket
    
    ' Create a request packet.
    Set Request = New CDNSPacket
    Request.SetHeader 1, 0, 0, 0, 0, 1, 0, 0, 0
    Request.AddQD Name, 1, 1
    packet = Request.BuildPacket
    
    ' Send the request.
    wsDNSclient.RemoteHost = server
    wsDNSclient.RemotePort = 53
    wsDNSclient.SendData packet
    
    ' Destroy the request object.
    Set Request = Nothing
End Sub

' Parse the reply of the above name lookup.  Add the result(s) to the cache.
Private Sub ParseReply(ByVal packet As String)
    Dim i As Integer

    Dim Reply As CDNSPacket
    Set Reply = New CDNSPacket

    Reply.ParsePacket (packet)

    ' Check for a referral (one question, no answers, and some name servers)
    If Reply.GetQDCount = 1 And Reply.GetANCount = 0 And Reply.GetNSCount > 0 Then
        Lookup Reply.GetQName(1), Reply.GetNSRData(1)
    Else
        For i = 1 To Reply.GetANCount
            ' Make sure we have an IP address A record in the response.
            If Reply.GetANRType(i) = 1 And Reply.GetANRClass(i) = 1 Then
                AddCacheEntry Reply.GetANRName(i), Reply.GetANRData(i), _
                    Reply.GetANRTTL(i)
            End If
        Next
    End If

    If Reply.GetANCount > 0 And ReplyHost.Waiting Then
        SendReply Reply.GetANRName(Reply.GetANCount)
    End If
    
    ' Destroy the reply object.
    Set Reply = Nothing
End Sub

' Send a reply to the client based on the contents of the cache.
Private Sub SendReply(ByVal Name As String)
    Dim packet, ABlock As String
    Dim i, Answers As Integer
    Dim Response As CDNSPacket
    
    Set Response = New CDNSPacket
    
    Response.SetHeader ReplyHost.Id, 1, 0, 0, 0, 0, 0, 0, 0
    
    ' Add the question record.
    Response.AddQD Name, 1, 1
    
    ' Build the answer block.
    For i = 1 To Cache.numEntries()
        ' If the name exists in the cache ...
        If Cache.getName(i) = Name Then
            ' Add an answer from the cache to the reply packet.
            Response.AddAN Cache.getName(i), 1, 1, _
                ComputeTTL(Cache.getExpiration(i)), _
                4, Cache.getAddress(i)
        End If
    Next i
    
    packet = Response.BuildPacket

    ' Set the hostname and port
    wsDNSserver.RemoteHost = ReplyHost.IP
    wsDNSserver.RemotePort = ReplyHost.Port
    ' Send the packet
    wsDNSserver.SendData packet
    
    ' Clean up.
    ReplyHost.Waiting = False
    Set Response = Nothing
End Sub
