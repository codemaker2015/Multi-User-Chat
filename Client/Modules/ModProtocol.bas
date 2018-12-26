Attribute VB_Name = "ModProtocol"
Option Explicit

Public Sub ParseUserEntersLeaves(Packet As String)
    'Packet is structured like this:
    '"ENT" & Chr$(2) & Nickname
    '"LEA" & Chr$(2) & Nickname
    If InStr(1, Packet, Chr$(2)) > 0 Then
        Dim strInfo() As String, bolEntering As Boolean
        
        strInfo = Split(Packet, Chr$(2))
        bolEntering = strInfo(0) = "ENT"
        AddUserEntersLeaves frmChat.rtbChat, strInfo(1), bolEntering
        
        If bolEntering Then
            frmChat.lstUsers.AddItem strInfo(1)
        Else
            RemoveListItem frmChat.lstUsers, strInfo(1)
        End If
        
    End If
    
End Sub

Public Sub ParseChatMessage(Packet As String)
    'Packet is structured like this:
    '"MSG" & Chr$(2) & Nickname & Chr$(2) & Message
    If InStr(1, Packet, Chr$(2)) > 0 Then
        Dim strInfo() As String
        
        strInfo = Split(Packet, Chr$(2))
        'strInfo(0) = MSG
        'strInfo(1) = Nickname
        'strInfo(2) = Message
        
        AddChatMessage frmChat.rtbChat, strInfo(1), strInfo(2)
        
        Erase strInfo
    End If
    
End Sub

Public Sub ParseUserList(Packet As String)
    'Packet is structured like this:
    '"LST" & Chr$(2) & User1 & vbCrLf & User2 & vbCrLf & User3
    If InStr(1, Packet, Chr$(2)) > 0 Then
        Dim strInfo() As String, strUsers() As String
        Dim intLoop As Integer
        
        strInfo() = Split(Packet, Chr$(2))
        'strInfo(0) = LST
        'strInfo(1) = User1 & vbCrLf & User2 & vbCrLf & User3
        With frmChat.lstUsers
            .Clear
            
            If InStr(1, strInfo(1), vbCrLf) > 0 Then
                strUsers = Split(strInfo(1), vbCrLf)
                
                For intLoop = 0 To UBound(strUsers)
                    If Len(strUsers(intLoop)) > 0 Then
                        .AddItem strUsers(intLoop)
                    End If
                Next intLoop
            Else
                .AddItem strInfo(1)
            End If
            
        End With
        
        Erase strInfo
        Erase strUsers
    End If
    
End Sub
