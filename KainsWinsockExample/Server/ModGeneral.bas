Attribute VB_Name = "ModGeneral"
Option Explicit

' ***** General Variables *****
Public Rdata As String ' Holds Recieved Data
Public NumUsers As Integer ' Holds how many people are connected
Public LoopC As Integer ' Servers Counter Loop

' ***** Public Constants *****
Public Const MAX_USERS As Integer = 999

' ***** Public Types *****
Public Type FreeSock
Used As Byte
UsedBy As String
End Type

Public Type Users
Name As String
Index As Integer
Connected As Byte
End Type

' ***** Public Type Arrays *****
Public FreeSocket(1 To MAX_USERS) As FreeSock
Public UserList(1 To MAX_USERS) As Users

Sub Connect(Index As Integer, RequestID As Long)


' ***** Add to UserList *****
UserList(Index).Index = Index
UserList(Index).Connected = 1
UserList(Index).Name = "Client" & UserList(Index).Index

' Load Winsock
Load FrmMain.WinSock(Index)

' Incrememnt users by 1
NumUsers = NumUsers + 1

' Accept User
FrmMain.WinSock(Index).Accept RequestID

' Let the person know there connected so as there TxtSend is unlocked
FrmMain.WinSock(Index).SendData "(Connected)"

'Add to Console
FrmMain.TxtConsole.Text = FrmMain.TxtConsole.Text & UserList(Index).Name & " Has connected!"


End Sub

Sub SendToAll(Rdata As String, Index As Integer)

For LoopC = 1 To NumUsers + 1 ' Go through all users including fact there could be a gap
If UserList(LoopC).Connected = 1 Then ' if user is connected
FrmMain.WinSock(LoopC).SendData UserList(Index).Name & ": " & Rdata ' Send data
End If ' if not, nothing
Next ' loop

End Sub
