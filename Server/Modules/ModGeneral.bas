Attribute VB_Name = "ModGeneral"
'modGeneral.bas
'--------------
'General subs/functions used throughout the program.
'These are simple things like adding a message to
'a RichTextBox, and other things that aren't really
'relevant/necessary for this chat example.
Option Explicit

'Program is ending?
Public bolEnding As Boolean

Public Sub EndProgram()
    If Not bolEnding Then
        bolEnding = True
        ModChat.CloseServer
        Erase ModChat.udtUsers
        Unload frmChat
        Unload frmStart
    End If
End Sub

'Removes a ListItem from a ListBox.
Public Sub RemoveListItem(ListObject As Object, _
    ItemText As String, Optional ByVal RemoveAll As Boolean = False)
    
    'Make sure ListObject is a ListBox or ComboBox.
    If Not TypeOf ListObject Is ListBox And Not _
           TypeOf ListObject Is ComboBox Then Exit Sub
           
    Dim intLoop As Integer, strL As String
    
    'Compare lowercase values (case-insensitive search).
    strL = LCase$(ItemText)
    
    With ListObject
        'Loop through all ListItems.
        For intLoop = 0 To .ListCount - 1
            'If LowerCase(this item) equals item we're searching for...
            If LCase$(.List(intLoop)) = strL Then
                .RemoveItem intLoop
                If Not RemoveAll Then Exit For
            End If
        Next intLoop
    End With
    
End Sub

'Checks if an item exists in a ListBox.
'If it does, it returns the ListIndex.
Public Function FindListItem(ListObject As Object, _
    ItemText As String) As Integer
    
    'Make sure ListObject is a ListBox or ComboBox.
    If Not TypeOf ListObject Is ListBox And Not _
           TypeOf ListObject Is ComboBox Then Exit Function
           
    Dim intLoop As Integer, strL As String
    
    'Compare lowercase values (case-insensitive search).
    strL = LCase$(ItemText)
    
    With ListObject
        'Loop through all ListItems.
        For intLoop = 0 To .ListCount - 1
            'If LowerCase(this item) equals item we're searching for...
            If LCase$(.List(intLoop)) = strL Then
                'Return the index of the item.
                FindListItem = intLoop
                'Exit.
                Exit For
            End If
        Next intLoop
    End With
    
End Function

'Adds a status message to the chat RTB.
Public Sub AddStatusMessage(RTB As Object, _
    ByVal Color As Long, Message As String)
    
    'Make sure RTB is a RichTextBox.
    If Not TypeOf RTB Is RichTextBox Then Exit Sub
    
    With RTB
        .SelStart = Len(.Text) 'Move to end of RTB.
        .SelFontName = "Tahoma" 'Set font.
        .SelFontSize = 9 'Set font size.
        .SelBold = False
        .SelItalic = False
        .SelUnderline = False
        .SelColor = Color
        .SelText = Message & vbCrLf
    End With
    
End Sub

'Displays a user entering/leaving message.
Public Sub AddUserEntersLeaves(RTB As Object, _
    Username As String, Optional ByVal Entering As Boolean = True)
    
    If Not TypeOf RTB Is RichTextBox Then Exit Sub
    
    Dim strL As String
    
    strL = LCase$(Username)
    
    With RTB
        .SelStart = Len(.Text)
        .SelFontName = "Tahoma"
        .SelFontSize = 9
        .SelBold = True
        .SelItalic = False
        .SelUnderline = False
        .SelColor = IIf(strL = LCase$(strMyNickname), RGB(0, 0, 128), RGB(128, 0, 0))
        .SelText = Username & " "
        .SelBold = False
        .SelColor = IIf(Entering, RGB(0, 128, 0), RGB(128, 128, 128))
        .SelText = IIf(Entering, "has joined the room.", "has left the room.") & vbNewLine
    End With
    
End Sub

'Displays a chat message.
Public Sub AddChatMessage(RTB As Object, _
    Username As String, Message As String)
    
    If Not TypeOf RTB Is RichTextBox Then Exit Sub
    
    Dim strL As String
    
    strL = LCase$(Username)
    
    With RTB
        .SelStart = Len(.Text)
        .SelFontName = "Tahoma"
        .SelFontSize = 9
        .SelBold = True
        .SelItalic = False
        .SelUnderline = False
        .SelColor = IIf(strL = LCase$(strMyNickname), RGB(0, 0, 128), RGB(128, 0, 0))
        .SelText = Username & ": "
        .SelBold = False
        .SelColor = vbBlack
        .SelText = Message & vbCrLf
    End With
    
End Sub

