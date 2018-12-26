Attribute VB_Name = "ModProtocol"
Option Explicit

'Used to "flash" the window when a message is received
'if the form is minimized.
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

'This is where all the data is handled.
'Different subs for different "packets".
'Keeps it simple and easy to modify/debug later on.

'Parse connection packet.
'------------------------
'The client sends this when they first connect
'telling us their nickname.
'Put their nickname into their array slot.
Public Sub ParseConnection(Index As Integer, Packet As String)
    'Packet is structured like this:
    '"CON" & Chr$(2) & Nickname
    
    'Check for packet seperator (Chr$(2)).
    If InStr(1, Packet, Chr$(2)) > 0 Then
        'Found.
        'Split packet up.
        Dim strInfo() As String, strPacket As String
        
        strInfo = Split(Packet, Chr$(2))
        
        'strInfo(0) = CON
        'strInfo(1) = Nickname
        
        'Store their nickname.
        udtUsers(Index).strNickname = strInfo(1)
        
        'Send user enters to everyone connected.
        'Build the packet.
        strPacket = "ENT" & Chr$(2) & strInfo(1) & Chr$(4)
        SendGlobalData strPacket
        
        'Show that the user has joined the room.
        AddUserEntersLeaves frmChat.rtbChat, strInfo(1), True
        
        'Add user's nickname to user list.
        'Removing it first...
        RemoveListItem frmChat.lstUsers, strInfo(1), False
        frmChat.lstUsers.AddItem strInfo(1)
        
        'Send user list to everyone.
        SendUserList Index
        
        Erase strInfo
    End If
    
End Sub

'Parse chat message packet.
'--------------------------
Public Sub ParseChatMessage(Index As Integer, Packet As String)
    'Packet is structured like this:
    '"MSG" & Chr$(2) & Nickname & Chr$(2) & Message
    
    If InStr(1, Packet, Chr$(2)) > 0 Then
        Dim strInfo() As String, strPacket As String
        
        strInfo = Split(Packet, Chr$(2))
        
        'strInfo(0) = MSG
        'strInfo(1) = Nickname
        'strInfo(2) = Message
        
        AddChatMessage frmChat.rtbChat, strInfo(1), strInfo(2)
        
        'Relay message to all other clients.
        strPacket = Packet & Chr$(4)
        SendGlobalData strPacket
        
        Erase strInfo
        
        'Flash window.
        If frmChat.WindowState = vbMinimized Then
            FlashWindow frmChat.hwnd, 1
        End If
        
    End If
    
End Sub

'Builds a user list of all connected clients.
'Each user is separated by vbCrLf (new line) (unlikely to show in username).
Public Function BuildUserList() As String
    Dim intLoop As Integer, strRet As String
    
    strRet = strMyNickname & vbCrLf
    
    For intLoop = 0 To UBUsers
        With udtUsers(intLoop)
            If Len(.strNickname) > 0 And Len(.strIP) > 0 Then
                strRet = strRet & .strNickname & vbCrLf
            End If
        End With
    Next intLoop
    
    If Len(strRet) > 0 Then
        If Right$(strRet, 2) = vbCrLf Then strRet = Mid$(strRet, 1, Len(strRet) - 2)
    End If
    
    BuildUserList = strRet
End Function

'Sends the user list to every client.
Public Sub SendUserList(Index As Integer)
    Dim strPacket As String
    
    strPacket = "LST" & Chr$(2) & BuildUserList & Chr$(4)
    
    If frmChat.sckServer(Index).State = sckConnected Then
        frmChat.sckServer(Index).SendData strPacket
    End If
    
End Sub
