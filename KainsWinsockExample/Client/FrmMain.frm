VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Kains Winsock Example (Single Socket) Client"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock WinSock 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "Send"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox TxtSend 
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3240
      Width           =   4695
   End
   Begin VB.TextBox TxtRecieved 
      Height          =   3015
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSend_Click()

' Send the value of TxtSend to the server for analysis
Winsock.SendData TxtSend.Text

TxtSend.Text = "" ' Empty TxtSend ready to start typing again

TxtSend.SetFocus ' Set focus on txtsend so typing goes straight in

End Sub

Private Sub Form_Load()

Winsock.RemoteHost = "127.0.0.1" ' IP = LocalHost (For testing Only)
Winsock.RemotePort = 27478 ' Same port as server default

Winsock.Connect 'try to connect to server

End Sub

Private Sub TxtSend_KeyUp(KeyCode As Integer, Shift As Integer)

' If user press enter while focus is set on TxtSend
' then call CmdSend_Click
' (Send stuff and reset TxtSend)
If KeyCode = vbKeyReturn Then
Call CmdSend_Click
End If

End Sub

Private Sub WinSock_Close()

Unload Me ' Unload this form (The only form)

End Sub

Private Sub WinSock_DataArrival(ByVal bytesTotal As Long)

Winsock.GetData RData ' Add incoming data to Rdata (String)

If RData = "(Connected)" Then ' if Rdata = (Connected)
TxtSend.Locked = False ' unload txtsend (Users can type now)
TxtRecieved.Text = TxtRecieved.Text & vbNewLine & "Connected..." ' Add to txtrecieved
Exit Sub ' (Exit sub)
End If ' If Not, Continue

' Add a carriage return to txtrecieved then add Rdata (String)
TxtRecieved.Text = TxtRecieved.Text & vbCrLf & RData
' Sort TxtBox out
TxtRecieved.SelLength = TxtRecieved.Height
End Sub

