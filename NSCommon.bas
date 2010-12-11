Attribute VB_Name = "NSCommon"
' Common Name Server Routines
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

'--[ Conversion Functions ]----------------------------------------------------

Public Function Num2Str(ByVal num As Long, ByVal numOctets As Integer) As String
    Dim ret As String
    Dim i As Integer

    ret = ""
    For i = 1 To numOctets
        ret = Chr(num Mod 256) & ret
        num = num \ 256
    Next i
    Num2Str = ret
End Function

Public Function Str2Num(ByVal S As String) As Long
    Dim ret As Long
    Dim i As Integer
    
    ret = 0
    For i = 1 To Len(S)
        ret = ret * 2 + Asc(Mid(S, i, 1))
    Next i
    Str2Num = ret
End Function

Public Function ClrBits(ByVal octet As Byte, ByVal startBit As Integer, ByVal nBits As Integer) As Byte
    
    Dim mask As Byte
    
    If startBit + nBits > 8 Then
        nBits = 8 - startBit
    End If
    
    mask = (2 ^ nBits - 1) * 2 ^ startBit
    ClrBits = octet And (Not mask)
End Function

Public Function SetBits(ByVal octet As Byte, ByVal startBit As Integer, ByVal nBits As Integer) As Byte
    
    Dim mask As Byte
    
    If startBit + nBits > 8 Then
        nBits = 8 - startBit
    End If
    
    mask = (2 ^ nBits - 1) * 2 ^ startBit
    SetBits = octet Or mask
End Function

Public Function ExtractField(ByVal octet As Byte, ByVal startBit As Integer, ByVal nBits As Integer) As Byte
   
    Dim mask As Byte
    
    If startBit + nBits > 8 Then
        nBits = 8 - startBit
    End If
    
    mask = 2 ^ nBits - 1
    ExtractField = octet \ 2 ^ startBit And mask
End Function

Public Function InsertField(ByVal octet As Byte, ByVal field As Byte, ByVal startBit As Integer, ByVal nBits As Integer)
    
    If startBit + nBits > 8 Then
        nBits = 8 - startBit
    End If

    InsertField = ClrBits(octet, startBit, nBits) _
        Or (ExtractField(field, 0, nBits) * 2 ^ startBit)
End Function

Public Function calcExpiration(ByVal TTL As Long) As Date
    calcExpiration = DateAdd("s", TTL, Now)
End Function

Public Function ComputeTTL(ByVal Expiration As Date) As Long
    ComputeTTL = DateDiff("s", Now, Expiration)
End Function

'--[ String Functions ]--------------------------------------------------------

Public Function strpos(ByVal S As String, ByVal c As String) As Long
    Dim i As Long
    Dim done As Boolean
    
    i = 0
    done = False
    While i <= Len(S) And Not done
        i = i + 1
        If (Mid(S, i, 1) = c) Then
           done = True
        End If
    Wend
    strpos = i
End Function

Public Function IPString(ByVal IP As String) As String
    Dim i As Integer
    Dim Address As String
    
    Address = ""
    
    ' Step through each of the four octets in an IP address
    For i = 1 To 4
        ' Convert the octet's value into a decimal representation
        Address = Address & Str2Num(Mid(IP, i, 1)) & "."
    Next i
    IPString = Mid(Address, 1, Len(Address) - 1)
End Function

Public Function fmtName(ByVal Query As String) As String
    Dim ret As String
    Dim i As Integer
    
    While Len(Query) > 0
        i = strpos(Query, ".")
        If Not i = 0 Then
            ret = ret & Chr(i - 1) & Mid(Query, 1, i - 1)
            Query = Mid(Query, i + 1)
        End If
    Wend
    fmtName = ret + Chr(0)
End Function

' Returns the length of the RNAME field in 'packet' starting at 'pos'.  The
' field is either null-terminated or ends in an offset pointer to elsewhere
' in the packet.  We account for the two octet pointer in our length result.
Public Function NameLengthAt(ByVal packet As String, ByVal pos As Integer) As Long
    Dim c As String
    Dim count As Integer
    
    count = 0
    c = "_"
    Do While (Not Asc(c) = 0) And (Asc(c) < &HC0)
        c = Mid(packet, pos + count, 1)
        count = count + 1
    Loop
    
    If Asc(c) >= &HC0 Then
        count = count + 1
    End If
    
    NameLengthAt = count
End Function

' Extracts a name from the packet at the given offset.  This function will
' recurse over the pointers until the entire name is decoded.
Public Function NameAt(ByVal packet As String, ByVal pos As Integer) As String
    Dim HiOctet, LoOctet, Offset As Integer
    Dim Name As String
    Dim NextLen, i As Integer
  
    If pos < Len(packet) Then
        NextLen = -1
        While Not NextLen = 0
            ' Check if we have a pointer to elsewhere in the packet.  If we do, adjust
            ' the offset accordingly.
            If Asc(Mid(packet, pos, 1)) = 192 Then
                HiOctet = Asc(Mid(packet, pos, 1))
                LoOctet = Asc(Mid(packet, pos + 1, 1))
                HiOctet = HiOctet - &HC0
                Offset = (HiOctet * 256 + LoOctet) + 1
                ' Recurse ...
                Name = Name & NameAt(packet, Offset) & "."
                NextLen = 0
            Else
                NextLen = Asc(Mid(packet, pos, 1))
                If Not NextLen = 0 Then
                    For i = 1 To NextLen
                        Name = Name & Mid(packet, pos + i, 1)
                    Next
                    Name = Name & "."
                    pos = pos + NextLen + 1
                    NextLen = -1
                End If
            End If
        Wend
    
        ' Return the name string minus the last period.
        NameAt = Mid(Name, 1, Len(Name) - 1)
    Else
        NameAt = "(error)"
    End If
End Function

' Check whether the given string already exists in the packet.  If it does,
' return a pointer string.  If it doesn't already exist, just return the
' string literal again.
Public Function AddName(ByVal packet As String, ByVal Name As String)
    Dim ret As String
    
    ' Set the result to the initial string (effectively a NOOP).
    ret = Name
    
    ' Check if the string 'name' already exists in the packet.
    If Not InStr(packet, Name) = 0 Then
        ret = Chr(&HC0) & Num2Str(InStr(packet, Name) - 1, 1)
    End If
    
    ' Return the result.
    AddName = ret
End Function

Public Function NullString(ByVal packet As String, ByVal pos As Integer) As String
    Dim nullptr As Long
    
    ' Find the position of the null.
    nullptr = strpos(Mid(packet, pos), Chr(0))
    
    ' Return the string portion before the null.
    NullString = Mid(packet, pos, nullptr - 1)
End Function
