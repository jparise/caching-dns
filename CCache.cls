VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' DNS Caching Module
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

'--[ Cache Structures ]--------------------------------------------------------

Private mCache As Collection

'--[ Class Methods ]-----------------------------------------------------------

Private Sub Class_Initialize()
    Set mCache = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCache = Nothing
End Sub

'--[ Public Methods ]----------------------------------------------------------

Public Sub Flush()
    Set mCache = Nothing
    Set mCache = New Collection
End Sub

Public Sub Add(ByVal Name As String, ByVal Address As String, ByVal TTL As Long)
    Dim Entry As New CCacheEntry
    
    If hasName(Name) = True And hasAddress(Address) = True Then
        ' Update any existing entries with a more recent TTL
        Update Name, Address, TTL
    Else
        ' Populate the cache entry structure
        Entry.Name = Name
        Entry.Address = Address
        Entry.Expiration = calcExpiration(TTL)
    
        ' Add a new entry to the cache array
        mCache.Add Entry
    End If
End Sub

Public Sub Remove(ByVal Index As Long)
    mCache.Remove Index
End Sub

Public Sub Update(ByVal Name As String, ByVal Address As String, ByVal TTL)
    Dim i As Integer
    
    For i = 1 To mCache.count
        If mCache(i).Name = Name And mCache(i).Address = Address Then
            mCache(i).Expiration = calcExpiration(TTL)
        End If
    Next
    
End Sub

Public Function numEntries() As Long
    numEntries = mCache.count
End Function

Public Function hasName(ByVal Name As String) As Boolean
    Dim i As Integer
    Dim Hit As Boolean

    Hit = False
    
    For i = 1 To mCache.count
        If mCache(i).Name = Name Then
            Hit = True
        End If
    Next
    
    hasName = Hit
    
End Function

Public Function hasAddress(ByVal Address As String) As Boolean
    Dim i As Integer
    Dim Hit As Boolean
    
    Hit = False
    
    For i = 1 To mCache.count
        If mCache(i).Address = Address Then
            Hit = True
        End If
    Next
    
    hasAddress = Hit
    
End Function

Public Sub Expire()
    Dim i As Integer
    
    i = mCache.count
    
    While i > 0
        If mCache(i).Expiration < Now Then
            mCache.Remove (i)
        End If
        i = i - 1
    Wend

End Sub

Public Function getName(ByVal Index As Long) As String
    getName = mCache(Index).Name
End Function

Public Function getAddress(ByVal Index As Long) As String
    getAddress = mCache(Index).Address
End Function

Public Function getExpiration(ByVal Index As Long) As Date
    getExpiration = mCache(Index).Expiration
End Function
