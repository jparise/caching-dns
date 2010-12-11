VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDNSPacket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Name Server Packet Crunching Class
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

'--[ Packet Structures ]-------------------------------------------------------

' Header:
'                                    1  1  1  1  1  1
'      0  1  2  3  4  5  6  7  8  9  0  1  2  3  4  5
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+
'    |                      ID                       |
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+
'    |QR|   Opcode  |AA|TC|RD|RA|   Z    |   RCODE   |
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+
'    |                    QDCOUNT                    |
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+
'    |                    ANCOUNT                    |
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+
'    |                    NSCOUNT                    |
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+
'    |                    ARCOUNT                    |
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+

Private Type pktHeaderType
    Id As Long          ' 16-bit unique identifier
    QR As Byte          ' 1-bit field that denotes a query (0) or response (1)
    Opcode As Byte      ' 4-bit field that indicates the kind of query
    AA As Byte          ' 1-bit field indicating whether this is authoritative
    TC As Byte          ' 1-bit field indicating this message was truncated
    RD As Byte          ' 1-bit field indicating that recursion is desired
    RA As Byte          ' 1-bit field indicating that recursion is available
    Z As Byte           ' 3-bit reserved field
    Rcode As Byte       ' 4-bit response code
    QDCount As Long     ' 16-bit field specifying the number of questions
    ANCount As Long     ' 16-bit field specifying the number of answers
    NSCount As Long     ' 16-bit field specifying the number of authorities
    ARCount As Long     ' 16-bit field specifying the number of additionals
End Type

' Question Record:
'                                    1  1  1  1  1  1
'      0  1  2  3  4  5  6  7  8  9  0  1  2  3  4  5
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+
'    |                                               |
'    /                     QNAME                     /
'    /                                               /
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+
'    |                     QTYPE                     |
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+
'    |                     QCLASS                    |
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+

Private Type pktQuestionType
    QName As String     ' Formatted domain name terminated by a null
    QType As Integer    ' 16-bit field containing the query type
    QClass As Integer   ' 16-bit field containing the query class
End Type

' Resouce Record:
'                                    1  1  1  1  1  1
'      0  1  2  3  4  5  6  7  8  9  0  1  2  3  4  5
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+
'    |                                               |
'    /                                               /
'    /                      NAME                     /
'    |                                               |
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+
'    |                      TYPE                     |
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+
'    |                     CLASS                     |
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+
'    |                      TTL                      |
'    |                                               |
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+
'    |                   RDLENGTH                    |
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--|
'    /                     RDATA                     /
'    /                                               /
'    +--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+--+

Private Type pktResourceType
    RName As String     ' Formatted domain name terminated by a null
    RType As Integer    ' 16-bit field containing the meaning of RDATA
    RClass As Integer   ' 16-bit field containing the class of RDATA
    TTL As Long         ' 32-bit field specifying the TTL in seconds
    RDLength As Long    ' 16-bit field specifying the length of RDATA
    RData As String     ' Varaible length string describing the resouce
End Type

' Complete DNS Packet:

Private Type pktDNSType
    Header As pktHeaderType         ' Packet header
    QDList() As pktQuestionType     ' Array of question blocks
    ANList() As pktResourceType     ' Array of answer blocks
    NSList() As pktResourceType     ' Array of authority blocks
    ARList() As pktResourceType     ' Array of additional blocks
End Type

' Class variables:

Private pktDNS As pktDNSType


'--[ Public Methods ]----------------------------------------------------------

' Parses the given packet and populates the pktDNS structure with all of the
' values from the packet.
Public Sub ParsePacket(ByVal packet As String)
    Dim i, ptr As Long
    Dim HiOctet, LoOctet As Byte

    ' Grab the identifier field and make sure it corresponds to the query
    ' packet we just sent.
    pktDNS.Header.Id = Str2Num(Mid(packet, 1, 2))
    
    ' Read the packet status bits (high octet)
    HiOctet = Str2Num(Mid(packet, 3, 1))
    pktDNS.Header.QR = ExtractField(HiOctet, 7, 1)
    pktDNS.Header.Opcode = ExtractField(HiOctet, 3, 4)
    pktDNS.Header.AA = ExtractField(HiOctet, 2, 1)
    pktDNS.Header.TC = ExtractField(HiOctet, 1, 1)
    pktDNS.Header.RD = ExtractField(HiOctet, 0, 1)
    ' Read the packet status bits (low octet)
    LoOctet = Str2Num(Mid(packet, 4, 1))
    pktDNS.Header.RA = ExtractField(LoOctet, 7, 1)
    pktDNS.Header.Z = ExtractField(LoOctet, 4, 3)
    pktDNS.Header.Rcode = ExtractField(LoOctet, 0, 4)
    
    ' Read the record counts
    pktDNS.Header.QDCount = Str2Num(Mid(packet, 5, 2))
    pktDNS.Header.ANCount = Str2Num(Mid(packet, 7, 2))
    pktDNS.Header.NSCount = Str2Num(Mid(packet, 9, 2))
    pktDNS.Header.ARCount = Str2Num(Mid(packet, 11, 2))
    
    ' Set the pointer to the start of the question block.
    ptr = 13
    
    ' Now that we have the start of the query records, loop through them
    ' and extract the name (QNAME), type (QTYPE), and class (QCLASS).
    ReDim pktDNS.QDList(pktDNS.Header.QDCount)
    For i = 1 To pktDNS.Header.QDCount
        ' Get the name
        pktDNS.QDList(i).QName = NameAt(packet, ptr)
        ptr = ptr + NameLengthAt(packet, ptr)

        ' Get the query type and class
        pktDNS.QDList(i).QType = Str2Num(Mid(packet, ptr, 2))
        pktDNS.QDList(i).QClass = Str2Num(Mid(packet, ptr + 2, 2))
        
        ' Advance the pointer to the end of the query record.
        ptr = ptr + 4
    Next i
    
    ' Extract the contents of the answer records (NAME, TYPE, CLASS, TTL,
    ' RDLENGTH, and RDATA).
    ReDim pktDNS.ANList(pktDNS.Header.ANCount)
    For i = 1 To pktDNS.Header.ANCount
        ' Get the name
        pktDNS.ANList(i).RName = NameAt(packet, ptr)
        ptr = ptr + NameLengthAt(packet, ptr)

        ' Get the response type, response class, and TTL.
        pktDNS.ANList(i).RType = Str2Num(Mid(packet, ptr, 2))
        pktDNS.ANList(i).RClass = Str2Num(Mid(packet, ptr + 2, 2))
        pktDNS.ANList(i).TTL = Str2Num(Mid(packet, ptr + 4, 4))
        
        ' Advance the pointer to the address field response data length field (RRDLen).
        ptr = ptr + 8
        
        ' Read the RRDLen field and advance the pointer.
        pktDNS.ANList(i).RDLength = Str2Num(Mid(packet, ptr, 2))
        ptr = ptr + 2

        ' Read the RDATA field.
        pktDNS.ANList(i).RData = Mid(packet, ptr, pktDNS.ANList(i).RDLength)

        ptr = ptr + pktDNS.ANList(i).RDLength
    Next i

    ' Extract the contents of the authority records (NAME, TYPE, CLASS, TTL,
    ' RDLENGTH, and RDATA).
    ReDim pktDNS.NSList(pktDNS.Header.NSCount)
    For i = 1 To pktDNS.Header.NSCount
        ' Get the name
        pktDNS.NSList(i).RName = NameAt(packet, ptr)
        ptr = ptr + NameLengthAt(packet, ptr)

        ' Get the response type, response class, and TTL.
        pktDNS.NSList(i).RType = Str2Num(Mid(packet, ptr, 2))
        pktDNS.NSList(i).RClass = Str2Num(Mid(packet, ptr + 2, 2))
        pktDNS.NSList(i).TTL = Str2Num(Mid(packet, ptr + 4, 4))

        ' Advance the pointer to the address field response data length field (RRDLen).
        ptr = ptr + 8

        ' Read the RRDLen field and advance the pointer.
        pktDNS.NSList(i).RDLength = Str2Num(Mid(packet, ptr, 2))
        ptr = ptr + 2

        ' Read the RDATA field.
        pktDNS.NSList(i).RData = NameAt(packet, ptr)

        ptr = ptr + pktDNS.NSList(i).RDLength
    Next i

    ' Extract the contents of the resource records (NAME, TYPE, CLASS, TTL,
    ' RDLENGTH, and RDATA).
    ReDim pktDNS.ARList(pktDNS.Header.ARCount)
    For i = 1 To pktDNS.Header.ARCount
        ' Get the name
        pktDNS.ARList(i).RName = NameAt(packet, ptr)
        ptr = ptr + NameLengthAt(packet, ptr)

        ' Get the response type, response class, and TTL.
        pktDNS.ARList(i).RType = Str2Num(Mid(packet, ptr, 2))
        pktDNS.ARList(i).RClass = Str2Num(Mid(packet, ptr + 2, 2))
        pktDNS.ARList(i).TTL = Str2Num(Mid(packet, ptr + 4, 4))

        ' Advance the pointer to the address field response data length field (RRDLen).
        ptr = ptr + 8
        
        ' Read the RRDLen field and advance the pointer.
        pktDNS.ARList(i).RDLength = Str2Num(Mid(packet, ptr, 2))
        ptr = ptr + 2

        ' Read the RDATA field.
        pktDNS.ARList(i).RData = Mid(packet, ptr, pktDNS.ARList(i).RDLength)

        ptr = ptr + pktDNS.ARList(i).RDLength
    Next i

End Sub

' Builds and teturns a complete DNS packet based on the values in the pktDNS
' packet structure.
Public Function BuildPacket() As String
    Dim i As Integer
    Dim packet As String
    Dim HiOctet, LoOctet As Byte

    ' Identifier field
    packet = Num2Str(pktDNS.Header.Id, 2)
    
    ' Flags (high and low octet)
    HiOctet = InsertField(HiOctet, pktDNS.Header.QR, 7, 1)
    HiOctet = InsertField(HiOctet, pktDNS.Header.Opcode, 3, 4)
    HiOctet = InsertField(HiOctet, pktDNS.Header.AA, 2, 1)
    HiOctet = InsertField(HiOctet, pktDNS.Header.TC, 1, 1)
    HiOctet = InsertField(HiOctet, pktDNS.Header.RD, 0, 1)
    LoOctet = InsertField(LoOctet, pktDNS.Header.RA, 7, 1)
    LoOctet = InsertField(LoOctet, pktDNS.Header.Z, 4, 3)
    LoOctet = InsertField(LoOctet, pktDNS.Header.Rcode, 0, 4)
    packet = packet & Num2Str(HiOctet, 1) & Num2Str(LoOctet, 1)
    
    ' Record counts
    packet = packet & Num2Str(pktDNS.Header.QDCount, 2)
    packet = packet & Num2Str(pktDNS.Header.ANCount, 2)
    packet = packet & Num2Str(pktDNS.Header.NSCount, 2)
    packet = packet & Num2Str(pktDNS.Header.ARCount, 2)
    
    ' Add the question blocks
    For i = 1 To pktDNS.Header.QDCount
        packet = packet & fmtName(pktDNS.QDList(i).QName)
        packet = packet & Num2Str(pktDNS.QDList(i).QType, 2)
        packet = packet & Num2Str(pktDNS.QDList(i).QClass, 2)
    Next i

    ' Add the answer blocks (in a random order)
    ShuffleANList
    For i = 1 To pktDNS.Header.ANCount
        packet = packet & AddName(packet, fmtName(pktDNS.ANList(i).RName))
        packet = packet & Num2Str(pktDNS.ANList(i).RType, 2)
        packet = packet & Num2Str(pktDNS.ANList(i).RClass, 2)
        packet = packet & Num2Str(pktDNS.ANList(i).TTL, 4)
        packet = packet & Num2Str(pktDNS.ANList(i).RDLength, 2)
        packet = packet & pktDNS.ANList(i).RData
    Next i
    
    ' Add the authority blocks
    For i = 1 To pktDNS.Header.NSCount
        packet = packet & AddName(packet, fmtName(pktDNS.NSList(i).RName))
        packet = packet & Num2Str(pktDNS.NSList(i).RType, 2)
        packet = packet & Num2Str(pktDNS.NSList(i).RClass, 2)
        packet = packet & Num2Str(pktDNS.NSList(i).TTL, 4)
        packet = packet & Num2Str(pktDNS.NSList(i).RDLength, 2)
        packet = packet & pktDNS.NSList(i).RData
    Next i
    
    ' Add the resource blocks
    For i = 1 To pktDNS.Header.ARCount
        packet = packet & AddName(packet, fmtName(pktDNS.ARList(i).RName))
        packet = packet & Num2Str(pktDNS.ARList(i).RType, 2)
        packet = packet & Num2Str(pktDNS.ARList(i).RClass, 2)
        packet = packet & Num2Str(pktDNS.ARList(i).TTL, 4)
        packet = packet & Num2Str(pktDNS.ARList(i).RDLength, 2)
        packet = packet & pktDNS.ARList(i).RData
    Next i

    ' Return the packet.
    BuildPacket = packet

End Function

' Sets the header values of the packet structure to the given parameters.  We
' also initialize all of the records counts to zero.
Public Sub SetHeader(ByVal Id As Long, ByVal QR As Byte, ByVal Opcode As Byte, _
                        ByVal AA As Byte, ByVal TC As Byte, ByVal RD As Byte, _
                        ByVal RA As Byte, ByVal Z As Byte, ByVal Rcode As Byte)

    ' Build the packet header with the supplied values.
    pktDNS.Header.Id = Id
    pktDNS.Header.QR = QR
    pktDNS.Header.Opcode = Opcode
    pktDNS.Header.AA = AA
    pktDNS.Header.TC = TC
    pktDNS.Header.RD = RD
    pktDNS.Header.RA = RA
    pktDNS.Header.Z = Z
    pktDNS.Header.Rcode = Rcode
    
    ' Set all of the record counts to zero.  They'll be added elsewhere.
    pktDNS.Header.QDCount = 0
    pktDNS.Header.ANCount = 0
    pktDNS.Header.NSCount = 0
    pktDNS.Header.ARCount = 0

End Sub

' Adds a new QD record to the array.
Public Sub AddQD(ByVal QName As String, ByVal QType As Integer, ByVal QClass As Integer)

    ' Increment the header's QDCount.
    pktDNS.Header.QDCount = pktDNS.Header.QDCount + 1

    ' Increase the size of the QDList array.
    ReDim Preserve pktDNS.QDList(pktDNS.Header.QDCount)

    ' Add the new question record.
    pktDNS.QDList(pktDNS.Header.QDCount).QName = QName
    pktDNS.QDList(pktDNS.Header.QDCount).QType = QType
    pktDNS.QDList(pktDNS.Header.QDCount).QClass = QClass

End Sub

' Adds a new AN record to the array.
Public Sub AddAN(ByVal RName As String, ByVal RType As Integer, _
                    ByVal RClass As Integer, ByVal TTL As Long, _
                    ByVal RDLength As Long, ByVal RData As String)

    ' Increment the header's ANCount.
    pktDNS.Header.ANCount = pktDNS.Header.ANCount + 1

    ' Increase the size of the ANList array.
    ReDim Preserve pktDNS.ANList(pktDNS.Header.ANCount)

    ' Add the new answer record.
    pktDNS.ANList(pktDNS.Header.ANCount).RName = RName
    pktDNS.ANList(pktDNS.Header.ANCount).RType = RType
    pktDNS.ANList(pktDNS.Header.ANCount).RClass = RClass
    pktDNS.ANList(pktDNS.Header.ANCount).TTL = TTL
    pktDNS.ANList(pktDNS.Header.ANCount).RDLength = RDLength
    pktDNS.ANList(pktDNS.Header.ANCount).RData = RData

End Sub

' Adds a new NS record to the array.
Public Sub AddNS(ByVal RName As String, ByVal RType As Integer, _
                    ByVal RClass As Integer, ByVal TTL As Long, _
                    ByVal RDLength As Long, ByVal RData As String)

    ' Increment the header's NSCount.
    pktDNS.Header.NSCount = pktDNS.Header.NSCount + 1

    ' Increase the size of the NSList array.
    ReDim Preserve pktDNS.NSList(pktDNS.Header.NSCount)

    ' Add the new authority record.
    pktDNS.NSList(pktDNS.Header.NSCount).RName = RName
    pktDNS.NSList(pktDNS.Header.NSCount).RType = RType
    pktDNS.NSList(pktDNS.Header.NSCount).RClass = RClass
    pktDNS.NSList(pktDNS.Header.NSCount).TTL = TTL
    pktDNS.NSList(pktDNS.Header.NSCount).RDLength = RDLength
    pktDNS.NSList(pktDNS.Header.NSCount).RData = RData

End Sub

' Adds a new AR record to the array.
Public Sub AddAR(ByVal RName As String, ByVal RType As Integer, _
                    ByVal RClass As Integer, ByVal TTL As Long, _
                    ByVal RDLength As Long, ByVal RData As String)

    ' Increment the header's ARCount.
    pktDNS.Header.ARCount = pktDNS.Header.ARCount + 1

    ' Increase the size of the ARList array.
    ReDim Preserve pktDNS.ARList(pktDNS.Header.ARCount)

    ' Add the new resource record.
    pktDNS.ARList(pktDNS.Header.ARCount).RName = RName
    pktDNS.ARList(pktDNS.Header.ARCount).RType = RType
    pktDNS.ARList(pktDNS.Header.ARCount).RClass = RClass
    pktDNS.ARList(pktDNS.Header.ARCount).TTL = TTL
    pktDNS.ARList(pktDNS.Header.ARCount).RDLength = RDLength
    pktDNS.ARList(pktDNS.Header.ARCount).RData = RData

End Sub

'--[ Accessor Methods ]--------------------------------------------------------

' Returns the ID field from the packet header.
Public Function GetID() As Long
    GetID = pktDNS.Header.Id
End Function

' Returns the QDCOUNT from the packet header.
Public Function GetQDCount() As Long
    GetQDCount = pktDNS.Header.QDCount
End Function

' Returns the QNAME field from the requested QD record.
Public Function GetQName(ByVal Index As Long) As String
    GetQName = pktDNS.QDList(Index).QName
End Function

' Returns the QTYPE field from the requested QD record.
Public Function GetQType(ByVal Index As Long) As Integer
    GetQType = pktDNS.QDList(Index).QType
End Function

' Returns the QCLASS field from the requested QD record.
Public Function GetQClass(ByVal Index As Long) As Integer
    GetQClass = pktDNS.QDList(Index).QClass
End Function

' Returns the ANCOUNT from the packet header.
Public Function GetANCount() As Long
    GetANCount = pktDNS.Header.ANCount
End Function

' Returns the RNAME field from the requested AN record.
Public Function GetANRName(ByVal Index As Long) As String
    GetANRName = pktDNS.ANList(Index).RName
End Function

' Returns the RDATA field from the requested AN record.
Public Function GetANRData(ByVal Index As Long) As String
    GetANRData = pktDNS.ANList(Index).RData
End Function

' Returns the RTYPE field from the requested AN record.
Public Function GetANRType(ByVal Index As Long) As Integer
    GetANRType = pktDNS.ANList(Index).RType
End Function

' Returns the RCLASS field from the requested AN record.
Public Function GetANRClass(ByVal Index As Long) As Integer
    GetANRClass = pktDNS.ANList(Index).RClass
End Function

' Returns the RTTL field from the requested AN record.
Public Function GetANRTTL(ByVal Index As Long) As Long
    GetANRTTL = pktDNS.ANList(Index).TTL
End Function

' Returns the NSCOUNT from the packet header.
Public Function GetNSCount() As Long
    GetNSCount = pktDNS.Header.NSCount
End Function

' Returns the RDATA field from the requested NS record.
Public Function GetNSRData(ByVal Index As Long) As String
    GetNSRData = pktDNS.NSList(Index).RData
End Function

'--[ Private Methods ]---------------------------------------------------------

' This routine shuffles the current list of AN records.  It uses a very simple
' swapping method to randomly re-order the list.
Private Sub ShuffleANList()
    Dim i, first, second As Long
    Dim tmpElement As pktResourceType
    
    Randomize
    For i = 1 To pktDNS.Header.ANCount
        first = Int(pktDNS.Header.ANCount * Rnd()) + 1
        second = Int(pktDNS.Header.ANCount * Rnd()) + 1
        tmpElement = pktDNS.ANList(first)
        pktDNS.ANList(first) = pktDNS.ANList(second)
        pktDNS.ANList(second) = tmpElement
    Next
    
End Sub
