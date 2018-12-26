VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Chat Client"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3960
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
   MaxButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect »"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtNickname 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Text            =   "Client"
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   3120
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "1234"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Caption         =   "Not connected."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   3840
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nickname:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Port:"
      Height          =   195
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   3840
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConnect_Click()
    'Check input.
    txtServer.Text = Trim$(txtServer.Text)
    txtPort.Text = Trim$(txtPort.Text)
    txtNickname.Text = Trim$(txtNickname.Text)
    
    If Len(txtServer.Text) = 0 Or Len(txtPort.Text) = 0 Or _
       Len(txtNickname.Text) = 0 Then
        
        MsgBox "Please fill in all fields!", vbCritical
        Exit Sub
    ElseIf Not IsNumeric(txtPort.Text) Then
        MsgBox "Invalid port value!", vbCritical
        Exit Sub
    End If
    
    'Done with that...
    strMyNickname = txtNickname.Text
    
    With frmChat.sckClient
        .Close
        bolRecon = False
        .RemoteHost = txtServer.Text
        .RemotePort = CInt(txtPort.Text)
        .Connect
    End With
    
    Me.Hide
    frmChat.Show
    AddStatusMessage frmChat.rtbChat, RGB(128, 128, 128), "> Connecting to " & txtServer.Text & ":" & txtPort.Text & "..."
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    'Number only.
    If Not IsNumeric(Chr$(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub
