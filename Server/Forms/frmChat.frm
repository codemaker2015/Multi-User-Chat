VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   Caption         =   "Chat Server - Chat"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   529
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   615
      Left            =   6720
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtMsg 
      Height          =   645
      Left            =   120
      MaxLength       =   1024
      TabIndex        =   2
      Top             =   3480
      Width           =   6495
   End
   Begin VB.ListBox lstUsers 
      Height          =   3255
      IntegralHeight  =   0   'False
      Left            =   5520
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5741
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmChat.frx":0000
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   5640
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'JUST USED FOR FORM RESIZING.
Private Type RECT
    rctLeft As Long
    rctTop As Long
    rctRight As Long
    rctBottom As Long
End Type

'JUST USED FOR FORM RESIZING.
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'JUST USED FOR FORM RESIZING.
Private udtMyRect As RECT

Private Sub cmdSend_Click()
    If Len(Trim$(txtMsg.Text)) > 0 Then
        'Build message packet and send to everyone.
        Dim strPacket As String
        
        strPacket = "MSG" & Chr$(2) & strMyNickname & Chr$(2) & txtMsg.Text & Chr$(4)
        SendGlobalData strPacket
        
        AddChatMessage rtbChat, strMyNickname, txtMsg.Text
        txtMsg.Text = ""
        On Error Resume Next
        txtMsg.SetFocus
    End If
End Sub

'Form is resizing.
'-----------------
'Resizes according to the CLIENT AREA of the form.
'Form.Width/Form.Height/Form.ScaleWidth/Form.ScaleHeight return width + non-client area (borders, etc.).
'This provides pixel perfect resizing regardless of which windows theme/screen resolution is being used.
Private Sub Form_Resize()
    'Don't do anything if form is being minimized.
    If Me.WindowState = vbMinimized Then Exit Sub
    
    GetClientRect Me.hwnd, udtMyRect

    rtbChat.Width = udtMyRect.rctRight - 176
    rtbChat.Height = udtMyRect.rctBottom - 64
    lstUsers.Height = rtbChat.Height
    lstUsers.Left = rtbChat.Width + 15
    txtMsg.Top = rtbChat.Height + 15
    txtMsg.Width = udtMyRect.rctRight - 96
    cmdSend.Top = txtMsg.Top
    cmdSend.Left = txtMsg.Width + 15
End Sub

'End program.
Private Sub Form_Unload(Cancel As Integer)
    EndProgram
End Sub

'Contents of rtbChat have changed.
Private Sub rtbChat_Change()
    'Auto-scroll box.
    rtbChat.SelStart = Len(rtbChat.Text)
    'Sets cursor position (carrot) to end (length of text) of RTB.
End Sub

'A client has disconnected.
'--------------------------
'Send everyone else a message that this user has left the room.
'Then close the Winsock control so it will be ready for any
'new connections.
Private Sub sckServer_Close(Index As Integer)
    sckServer(Index).Close
    
    Dim strPacket As String
    
    strPacket = "LEA" & Chr$(2) & udtUsers(Index).strNickname & Chr$(4)
    SendGlobalData strPacket
    
    AddUserEntersLeaves rtbChat, udtUsers(Index).strNickname, False
    RemoveListItem lstUsers, udtUsers(Index).strNickname
    
    With udtUsers(Index)
        .strBuffer = vbNullString
        .strIP = vbNullString
        .strNickname = vbNullString
    End With
    
End Sub

'A client is attempting to connect.
'----------------------------------
'Another computer is trying to connect to the server.
'Find a socket we can use to handle the connection.
'Then load a slot in the udtUsers() array for this user.
Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim intNext As Integer
    
    intNext = ModChat.NextOpenSocket
    
    If intNext > 0 Then
        'Found a socket to use; accept connection.
        sckServer(intNext).Accept requestID
        
        'Check if there is a slot open for this connection
        'in the users array.
        If UBUsers < intNext Then
            'There isn't, load one.
            ReDim Preserve udtUsers(intNext) As CHAT_USER
        End If
        
        '(Re)set this client's info.
        With udtUsers(intNext)
            .strIP = sckServer(intNext).RemoteHostIP
            .strNickname = vbNullString
        End With
        
        'We haven't received the user's nickname yet.
        'That will happen in the DataArrival event :)
        'Once it does, we will need to let everyone know that this person joined the room.
        AddStatusMessage rtbChat, RGB(0, 0, 128), "> " & sckServer(intNext).RemoteHostIP & " connected!"
    End If
    
End Sub

'Data has been sent by the client.
'---------------------------------
'This is where we handle all data received.
'This code may seem overcomplicated compared to others but it is important to do it this way.

'This project uses the TCP protocol.
'TCP is the standard for most internet applications.
'TCP "guarentees" the data will arrive correctly and in the same order.
'But sometimes TCP will split data up and it can arrive in pieces.
'Our program needs to be able to handle this.
'A single message may get split up into different messages, or two or more messages may arrive as one.
'Example:
'--------
'Winsock.SendData "1234"
'Winsock.SendData "5678"

'You would expect that to arrive as:
'1234
'5678

'But it will usually arrive as:
'12345678

'If we find a truncated packet, we store it somewhere else
'and don't process it until the whole thing has arrived.
'Every "packet" is delimited (separated) by: Chr$(4) in this example.
'(the individual pieces of information, ie: nickname, message, etc.) are separated by Chr$(2).
'This way, our program knows where one message starts/ends, and the next begins.
Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String, strPackets() As String
    Dim strTrunc As String, bolTrunc As Boolean
    Dim lonLoop As Long, lonTruncStart As Long
    Dim lonUB As Long
    
    'Get the received data.
    sckServer(Index).GetData strData, vbString, bytesTotal
    
    With udtUsers(Index)
        'Append it to the buffer.
        .strBuffer = .strBuffer & strData
        strData = vbNullString
        
        'Check if the last byte is the packet delimiter (Chr$(4)).
        'If it is, there are no truncated packets.
        'If it isn't, the last packet got split up. >:(
        If Right$(.strBuffer, 1) <> Chr$(4) Then
        
            'Get all data to the right of the last Chr$(4) which is the truncated packet.
            bolTrunc = True
            
            'Find position of last packet delimiter.
            lonTruncStart = InStrRev(.strBuffer, Chr$(4))
            
            'Check if it was found.
            If lonTruncStart > 0 Then
                'Extract out the truncated part.
                strTrunc = Mid$(.strBuffer, lonTruncStart + 1)
            End If
            
        End If
        
        'We checked if the data was truncated.
        'If it was, we put that part away for now and set the Truncated flag to TRUE (bolTrunc).
        
        'Split up the data buffer into individual packets
        'in case we received more than 1 at a time.
        'Process them individually.
        If InStr(1, .strBuffer, Chr$(4)) > 0 Then
            strPackets() = Split(.strBuffer, Chr$(4))
            
            'Now all of the individual packets are in strPackets().
            'Loop through all of them.
            lonUB = UBound(strPackets) 'Get number of packets.
            'If the data is truncated, don't process the last one
            'because it isn't complete.
            If bolTrunc Then lonUB = lonUB - 1
            
            'Start looping through all packets.
            For lonLoop = 0 To lonUB
                'Check length of packet.
                'Each packet has a command/header,
                'the packet must be at least that length.
                'In this example, all headers are 3 bytes/characters long.
                If Len(strPackets(lonLoop)) > 3 Then
                    'Look at the header and process the packet accordingly.
                    Select Case Left$(strPackets(lonLoop), 3)
                        
                        'Packet is a chat message.
                        Case "MSG"
                            'Process message.
                            ParseChatMessage Index, strPackets(lonLoop)
                            
                        'User is connecting (sending nickname).
                        Case "CON"
                            'Process connection.
                            ParseConnection Index, strPackets(lonLoop)
                        'Add your own here! :)
                        'Case "XXX"
                            'Do something.
                        
                        'Case "YYY"
                            'Do something.
                            
                    End Select
                End If
            Next lonLoop
        
        End If
        
        'We're done processing all packets.
        Erase strPackets
        
        'Now we can erase all the data we just processed from the buffer.
        'Otherwise, it will just keep growing in size and the same data
        'will be processed over and over (which might actually be kinda cool?).
        .strBuffer = vbNullString
        
        If bolTrunc Then
            'Still have a piece of a packet left over because the data was truncated.
            'Erase the buffer then put just the truncated part back in.
            .strBuffer = strTrunc
        End If
        
        strTrunc = vbNullString
    End With
    
End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdSend_Click
        KeyAscii = 0 'Gets rid of 'beep' sound.
    End If
End Sub
