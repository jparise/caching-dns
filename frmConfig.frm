VERSION 5.00
Begin VB.Form frmConfig 
   Caption         =   "Configuration"
   ClientHeight    =   2085
   ClientLeft      =   7860
   ClientTop       =   4815
   ClientWidth     =   2565
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   2565
   Begin VB.TextBox txtExpiry 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "1000"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtForwarder 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "129.21.3.17"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblExpiry 
      Caption         =   "Expiration frequency (ms):"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblForwarder 
      Caption         =   "Forward request to:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Name Server Configuration Window
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

Private Sub cmdDone_Click()
    frmStatus!wsDNSclient.RemoteHost = txtForwarder.Text
    frmStatus!cacheTimer.Interval = txtExpiry.Text
    Me.Hide
End Sub

Private Sub Form_Load()
    txtForwarder.Text = frmStatus!wsDNSclient.RemoteHost
    txtExpiry.Text = frmStatus!cacheTimer.Interval
End Sub

Private Sub Form_Terminate()
    Unload Me
End Sub
