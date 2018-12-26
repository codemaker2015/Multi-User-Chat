VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   Caption         =   "Chat Client - Chat"
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   529
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   615
      Left            =   6720
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtMsg 
      Height          =   645
      Left            =   120
      MaxLength       =   1024
      TabIndex        =   0
      Top             =   3480
      Width           =   6495
   End
   Begin VB.ListBox lstUsers 
      Height          =   3255
      IntegralHeight  =   0   'False
      Left            =   5520
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5741
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmChat.frx":0000
   End
   Begin MSWinsockLib.Winsock sckClient 
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

'Received data buffer.
Private strBuffer As String

Private Sub cmdSend_Click()
    If Len(txtMsg.Text) > 0 Then
        If sckClient.State <> sckConnected Then
            AddStatusMessage rtbChat, RGB(128, 0, 0), "> Not connected! Cannot send message."
        Else
            Dim strPacket As String
            
            strPacket = "MSG" & Chr$(2) & strMyNickname & Chr$(2) & txtMsg.Text & Chr$(4)
            sckClient.SendData strPacket
            txtMsg.Text = ""
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If sckClient.State <> sckConnected Then
            frmConnect.Show
            Unload Me
        End If
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
    If Not bolRecon Then
        EndProgram
    End If
End Sub

'Contents of rtbChat have changed.
Private Sub rtbChat_Change()
    'Auto-scroll box.
    rtbChat.SelStart = Len(rtbChat.Text)
    'Sets cursor position (carrot) to end (length of text) of RTB.
End Sub

'The server has closed the connection! Connection lost.
'------------------------------------------------------
Private Sub sckClient_Close()
    sckClient.Close
    bolRecon = True
    AddStatusMessage rtbChat, RGB(128, 128, 128), "> The connection to the server was lost! Press the [ESC] key to re-connect."
End Sub

Private Sub sckClient_Connect()
    AddStatusMessage rtbChat, RGB(0, 128, 0), "> Connected!"
    
    'Send nickname to server.
    Dim strPacket As String
    
    strPacket = "CON" & Chr$(2) & strMyNickname & Chr$(4)
    sckClient.SendData strPacket
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String, strPackets() As String
    Dim strTrunc As String, bolTrunc As Boolean
    Dim lonLoop As Long, lonTruncStart As Long
    Dim lonUB As Long
    
    sckClient.GetData strData, vbString, bytesTotal
    strBuffer = strBuffer & strData
    strData = vbNullString
    
    If Right$(strBuffer, 1) <> Chr$(4) Then
        bolTrunc = True
        lonTruncStart = InStrRev(strBuffer, Chr$(4))
        If lonTruncStart > 0 Then
            strTrunc = Mid$(strBuffer, lonTruncStart + 1)
        End If
    End If
    
    If InStr(1, strBuffer, Chr$(4)) > 0 Then
        strPackets() = Split(strBuffer, Chr$(4))
        lonUB = UBound(strPackets)
        
        If bolTrunc Then lonUB = lonUB - 1
        
        For lonLoop = 0 To lonUB
            If Len(strPackets(lonLoop)) > 3 Then
                
                Select Case Left$(strPackets(lonLoop), 3)
                    
                    'Packet is a chat message.
                    Case "MSG"
                        ParseChatMessage strPackets(lonLoop)
                        
                    'User list has been sent.
                    Case "LST"
                        ParseUserList strPackets(lonLoop)
                    
                    Case "ENT", "LEA"
                        ParseUserEntersLeaves strPackets(lonLoop)
                        
                    'Add your own here! :)
                    'Case "XXX"
                        'Do something.
                    
                    'Case "YYY"
                        'Do something.
                        
                End Select
            End If
        Next lonLoop
    
    End If
    
    Erase strPackets
    
    strBuffer = vbNullString
    
    If bolTrunc Then
        strBuffer = strTrunc
    End If
    
    strTrunc = vbNullString
End Sub

Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckClient.Close
    bolRecon = True
    AddStatusMessage rtbChat, RGB(128, 0, 0), "> Error (" & Number & "): " & Description & IIf(Right$(Description, 1) = ".", "", ".")
    AddStatusMessage rtbChat, RGB(128, 0, 0), "> Press the [ESC] key to re-connect."
End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdSend_Click
        KeyAscii = 0 'Gets rid of 'beep' sound.
    End If
End Sub
